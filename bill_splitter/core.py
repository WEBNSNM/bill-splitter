"""读取、分类、汇总与导出；不依赖界面，不修改输入文件。"""
import csv
import json
import re
import shutil
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from uuid import uuid4

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from .report import write_report

TEACHER = ('教师', '成人', '职工', '老师')
STUDENT = ('幼儿', '学生')
CENT = Decimal('0.01')


def classify(address):
    # 仅使用收货地址；“幼儿园”是机构名称，不是学生分类备注。
    text = str(address or '').replace('幼儿园', '')
    teacher = any(k in text for k in TEACHER)
    student = any(k in text for k in STUDENT)
    if teacher == student:
        return '待核对'
    return '老师' if teacher else '学生'


def money(value, row):
    try:
        result = Decimal(str(value).replace(',', '').strip())
        if not result.is_finite():
            raise InvalidOperation
        return result.quantize(CENT, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        raise ValueError(f'第 {row} 行金额无效或公式缺少缓存值：{value!r}') from None


def delivery_date(value, row, epoch):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        value = from_excel(value, epoch)
        if isinstance(value, datetime):
            return value.date()
    text = str(value or '').strip()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f'第 {row} 行配送日期无效：{value!r}')


def safe_text(cell, value):
    cell.value = value
    # 原始文本以 = 开头时按文本写出，避免被执行为公式。
    if isinstance(value, str):
        cell.data_type = 's'


def style_sheet(ws):
    edge = Side(style='thin', color='808080')
    border = Border(left=edge, right=edge, top=edge, bottom=edge)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    for cell in ws[1]:
        cell.font = Font(name='微软雅黑', bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='234E70')
        cell.alignment = Alignment(vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 32
    for col in ws.columns:
        values = [len(str(c.value or '')) for c in list(col)[:100]]
        ws.column_dimensions[col[0].column_letter].width = min(48, max(14, max(values) * 1.5 + 2))
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = '1:1'


def write_details(ws, headers, records, amount_index, date_index):
    ws.append(headers)
    for record in records:
        r = ws.max_row + 1
        for c, value in enumerate(record['values'], 1):
            safe_text(ws.cell(r, c), value)
        ws.cell(r, amount_index + 1, float(record['amount'])).number_format = '#,##0.00'
        ws.cell(r, date_index + 1, record['date']).number_format = 'yyyy-mm-dd'
    style_sheet(ws)


def export_bills(source, output_dir, amount_column='销售金额(元)', progress=None):
    source = Path(source).resolve()
    if source.suffix.lower() != '.xlsx':
        raise ValueError('仅支持 .xlsx 文件；请先在 Excel 中将 .xls 另存为 .xlsx。')
    wb = load_workbook(source, read_only=True, data_only=True)
    groups = defaultdict(list)
    exceptions = []
    skipped = 0
    source_amount = Decimal(0)
    pending = Decimal(0)
    try:
        ws = wb.worksheets[0]
        source_sheet = ws.title
        iterator = ws.iter_rows(values_only=True)
        headers = [str(v or '').strip() for v in next(iterator)]
        required = ['客户名', '收货地址', '配送日期', amount_column]
        for label in required:
            if headers.count(label) != 1:
                raise ValueError(f'第一行须有唯一表头“{label}”。')
        ni, ai, di, mi = [headers.index(label) for label in required]
        for line, values in enumerate(iterator, 2):
            if all(v is None or v == '' for v in values):
                continue
            school = str(values[ni] or '').strip()
            if (school in ('合计', '总计', '总合计') or not school) and not values[di] and not values[ai]:
                skipped += 1
                continue
            amount = money(values[mi], line)
            day = delivery_date(values[di], line, wb.epoch)
            category = classify(values[ai])
            source_amount += amount
            reason = ''
            if not school:
                reason = '客户名为空'
            elif category == '待核对':
                reason = ('同时命中老师和学生关键词' if any(k in str(values[ai] or '') for k in TEACHER)
                          else '未命中任何分类关键词')
            record = dict(values=list(values), amount=amount, date=day, category=category)
            if school:
                groups[school].append(record)
            if reason:
                pending += amount
                exceptions.append([line, school, values[ai], day.isoformat(), str(amount), reason])
        if not groups:
            raise ValueError('没有可导出的学校数据。')
    finally:
        wb.close()
    parent = Path(output_dir).resolve()
    parent.mkdir(parents=True, exist_ok=True)
    run_dir = parent / (datetime.now().strftime('账单_%Y%m%d_%H%M%S_') + uuid4().hex[:8])
    run_dir.mkdir()  # 每次独立目录，失败则删除本次不完整输出。
    try:
        school_reports = []
        for index, (school, records) in enumerate(groups.items(), 1):
            out = Workbook()
            detail = out.active
            detail.title = '全部明细'
            summary = out.create_sheet('总计表')
            teacher = out.create_sheet('老师明细')
            student = out.create_sheet('学生明细')
            write_details(detail, headers, records, mi, di)
            write_details(teacher, headers, [r for r in records if r['category'] == '老师'], mi, di)
            write_details(student, headers, [r for r in records if r['category'] == '学生'], mi, di)
            daily = defaultdict(lambda: {'老师': Decimal(0), '学生': Decimal(0)})
            for r in records:
                daily[r['date']]
                if r['category'] in ('老师', '学生'):
                    daily[r['date']][r['category']] += r['amount']
            summary.append(['配送日期', '老师', '学生', '总计'])
            for day, amounts in sorted(daily.items()):
                row = summary.max_row + 1
                summary.append([day, float(amounts['老师']), float(amounts['学生']), f'=SUM(B{row}:C{row})'])
                summary.cell(row, 1).number_format = 'yyyy-mm-dd'
            last = summary.max_row
            summary.append(['合计', f'=SUM(B2:B{last})', f'=SUM(C2:C{last})', f'=SUM(D2:D{last})'])
            for cells in summary.iter_rows(min_row=2, min_col=2, max_col=4):
                for cell in cells:
                    cell.number_format = '#,##0.00'
            style_sheet(summary)
            for cell in summary[summary.max_row]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill('solid', fgColor='DDEBF7')
            dates = [r['date'] for r in records]
            start, end = min(dates), max(dates)
            # 使用实际配送日期范围，包含年份，跨月或跨年不产生歧义。
            safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', school).rstrip(' .')[:70] or '未命名学校'
            filename = f'{index:02d}_{safe_name}_{start:%Y.%m.%d}-{end:%Y.%m.%d}对账明细.xlsx'
            out.save(run_dir / filename)
            out.close()
            school_reports.append(dict(school=school, rows=len(records), file=filename,
                                       total=str(sum((r['amount'] for r in records), Decimal(0))),
                                       pending_rows=sum(r['category'] == '待核对' for r in records)))
            if progress:
                progress(f'已生成 {index}/{len(groups)}：{school}')
        with (run_dir / '待核对记录.csv').open('w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['源行号', '客户名', '收货地址', '配送日期', '金额', '原因'])
            for row in exceptions:
                writer.writerow([("'" + str(v)) if str(v).startswith(('=', '+', '-', '@')) else v for v in row])
        result = dict(directory=str(run_dir), schools=len(groups), exceptions=len(exceptions),
                      skipped_totals=skipped, source_amount=f'{source_amount:.2f}',
                      classified_amount=f'{source_amount-pending:.2f}', pending_amount=f'{pending:.2f}',
                      source=str(source), amount_column=amount_column, school_details=school_reports,
                      source_sheet=source_sheet, address_column=get_column_letter(ai + 1),
                      customer_column=get_column_letter(ni + 1), report_html=str(run_dir / '核对报告.html'),
                      exception_details=[dict(source_row=r[0], source_sheet=source_sheet, school=r[1],
                                              address=r[2], date=r[3], amount=r[4], reason=r[5]) for r in exceptions],
                      note='总计表仅含老师+学生；待核对记录保留在全部明细。空客户名仅在待核对清单。金额按每行四舍五入到分。')
        (run_dir / '运行报告.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        write_report(result)
        return result
    except Exception:
        shutil.rmtree(run_dir)
        raise
