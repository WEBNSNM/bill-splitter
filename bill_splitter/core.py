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
DISCOUNT = Decimal('0.957')
GUOTOU_HEADERS = ('销售价国投结算单价', '销售国投结算金额')


def product_category(value):
    return str(value or '').split('>')[0].strip()


def settlement(price, quantity):
    unit = (price * DISCOUNT).quantize(CENT, rounding=ROUND_HALF_UP)
    return unit, (unit * quantity).quantize(CENT, rounding=ROUND_HALF_UP)


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
    body_alignment = Alignment(horizontal='center')
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = body_alignment
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    for cell in ws[1]:
        cell.font = Font(name='微软雅黑', bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='234E70')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.sheet_format.defaultRowHeight = 17.25
    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].height = 17.25
    for col in ws.columns:
        values = [len(str(c.value or '')) for c in list(col)[:100]]
        ws.column_dimensions[col[0].column_letter].width = min(48, max(14, max(values) * 1.5 + 2))
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = '1:1'


def write_details(ws, headers, records, amount_index, date_index, *, discounted=False):
    ws.append(headers)
    for record in records:
        r = ws.max_row + 1
        for c, value in enumerate(record['values'], 1):
            safe_text(ws.cell(r, c), value)
        ws.cell(r, amount_index + 1, float(record['amount'])).number_format = '#,##0.00'
        ws.cell(r, date_index + 1, record['date']).number_format = 'yyyy-mm-dd'
        if discounted:
            for column, value in enumerate(record['settlement'], len(headers) - 1):
                ws.cell(r, column, float(value)).number_format = '#,##0.00'
    style_sheet(ws)


def write_summary(ws, categories, daily, *, formulas=True):
    ws.append(['配送日期', *categories, '总计'])
    last_category = get_column_letter(len(categories) + 1)
    total_column = len(categories) + 2
    totals = [Decimal(0)] * len(categories)
    for day, amounts in sorted(daily.items()):
        row = ws.max_row + 1
        amounts = [amounts[name] for name in categories]
        totals = [total + amount for total, amount in zip(totals, amounts)]
        total = f'=SUM(B{row}:{last_category}{row})' if formulas else float(sum(amounts))
        ws.append([day, *map(float, amounts), total])
        ws.cell(row, 1).number_format = 'yyyy-mm-dd'
    last = ws.max_row
    if formulas:
        ws.append(['合计', *[f'=SUM({get_column_letter(c)}2:{get_column_letter(c)}{last})'
                             for c in range(2, total_column + 1)]])
    else:
        ws.append(['合计', *map(float, totals), float(sum(totals))])
    for cells in ws.iter_rows(min_row=2, min_col=2, max_col=total_column):
        for cell in cells:
            cell.number_format = '#,##0.00'
    style_sheet(ws)
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', fgColor='DDEBF7')


def append_guotou_sheets(book, headers, records, amount_index, date_index):
    summary = book.create_sheet('鸿康-国投折扣后(95.7%)对账单')
    categories = ('老师', '学生')
    daily = defaultdict(lambda: dict.fromkeys(categories, Decimal(0)))
    for category, title in zip(categories, ('鸿康-国投教师', '鸿康-国投学生')):
        rows = [r for r in records if r['category'] == category and 'settlement' in r]
        write_details(book.create_sheet(title), [*headers, *GUOTOU_HEADERS], rows,
                      amount_index, date_index, discounted=True)
        for record in rows:
            daily[record['date']][category] += record['settlement'][1]
    write_summary(summary, categories, daily, formulas=False)


def export_bills(source, output_dir, amount_column='销售金额(元)', progress=None, *, include_guotou=False):
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
        headers = [str(v or '').strip() for v in next(iterator, ())]
        required = ['客户名', '收货地址', '配送日期', amount_column]
        for label in required:
            if headers.count(label) != 1:
                raise ValueError(f'第一行须有唯一表头“{label}”。')
        ni, ai, di, mi = [headers.index(label) for label in required]
        if headers.count('对象名称') > 1:
            raise ValueError('第一行须有唯一表头“对象名称”。')
        oi = headers.index('对象名称') if '对象名称' in headers else None
        output_indexes = [i for i, label in enumerate(headers) if label != '序号']
        output_headers = [headers[i] for i in output_indexes]
        output_mi, output_di = output_indexes.index(mi), output_indexes.index(di)
        if include_guotou:
            settlement_labels = ('品类', '销售单价(元)', '销售数量')
            for label in settlement_labels:
                if headers.count(label) != 1:
                    raise ValueError(f'生成国投版时，第一行须有唯一表头“{label}”。')
            product_index, price_index, quantity_index = map(output_headers.index, settlement_labels)
        for line, values in enumerate(iterator, 2):
            if all(v is None or v == '' for v in values):
                continue
            school = str(values[ni] or '').strip()
            if (school in ('合计', '总计', '总合计') or not school) and not values[di] and not values[ai]:
                skipped += 1
                continue
            amount = money(values[mi], line)
            day = delivery_date(values[di], line, wb.epoch)
            source_amount += amount
            record = dict(values=[values[i] for i in output_indexes], amount=amount, date=day,
                          source_row=line, school=school, address=values[ai],
                          object_name=values[oi] if oi is not None else None)
            if school:
                groups[school].append(record)
            else:
                pending += amount
                exceptions.append([line, school, values[ai], day.isoformat(), str(amount),
                                   '客户名为空', record['object_name'], '客户名'])
        for records in groups.values():
            use_object = oi is not None
            for record in records:
                reason = ''
                if use_object:
                    object_name = str(record['object_name'] or '').strip()
                    matches = [name for name in ('老师', '教师', '学生', '活动单')
                               if name in object_name]
                    category = {'教师': '老师', '老师': '老师', '学生': '学生',
                                '活动单': '活动单'}.get(object_name, '待核对')
                    if category == '待核对':
                        reason = ('对象名称同时命中多个分类' if len(matches) > 1
                                  else '对象名称未命中任何分类')
                else:
                    category = classify(record['address'])
                    if category == '待核对':
                        reason = ('同时命中老师和学生关键词'
                                  if any(k in str(record['address'] or '') for k in TEACHER)
                                  else '未命中任何分类关键词')
                record['category'] = category
                record['use_object'] = use_object
                if include_guotou and category in ('老师', '学生'):
                    values = record['values']
                    if product_category(values[product_index]) != '米面油':
                        try:
                            price = Decimal(str(values[price_index]).replace(',', '').strip())
                            quantity = Decimal(str(values[quantity_index]).replace(',', '').strip())
                            if not price.is_finite() or not quantity.is_finite():
                                raise InvalidOperation
                            record['settlement'] = settlement(price, quantity)
                        except InvalidOperation:
                            raise ValueError(f"第 {record['source_row']} 行国投版的销售单价或销售数量无效。") from None
                if reason:
                    pending += record['amount']
                    exceptions.append([record['source_row'], record['school'], record['address'],
                                       record['date'].isoformat(), str(record['amount']), reason,
                                       record['object_name'], '对象名称' if use_object else '收货地址'])
        exceptions.sort(key=lambda item: item[0])
        if not groups:
            raise ValueError('没有可导出的学校数据。')
    finally:
        wb.close()
    parent = Path(output_dir).resolve()
    parent.mkdir(parents=True, exist_ok=True)
    run_dir = parent / (datetime.now().strftime('账单_%Y%m%d_%H%M%S_') + uuid4().hex[:8])
    run_dir.mkdir()  # 每次独立目录，失败则删除本次不完整输出。
    try:
        (run_dir / '学校版').mkdir()
        if include_guotou:
            (run_dir / '国投版').mkdir()
        school_reports = []
        for index, (school, records) in enumerate(groups.items(), 1):
            out = Workbook()
            detail = out.active
            detail.title = '底稿'
            suffix = '对账单'
            sheet_school = re.sub(r'[\\/*?:\[\]\x00-\x1f]', '_', school)
            summary_title = (sheet_school[:31 - len(suffix)] + suffix).lstrip("'")
            summary = out.create_sheet(summary_title)
            out.active = summary
            detail.sheet_state = 'hidden'
            teacher = out.create_sheet('老师明细')
            student = out.create_sheet('学生明细')
            categories = ['老师', '学生'] + (['活动单'] if any(r['category'] == '活动单' for r in records) else [])
            write_details(detail, output_headers, records, output_mi, output_di)
            for name, sheet in [('老师', teacher), ('学生', student)]:
                write_details(sheet, output_headers, [r for r in records if r['category'] == name],
                              output_mi, output_di)
            if '活动单' in categories:
                activity = out.create_sheet('活动单明细')
                write_details(activity, output_headers, [r for r in records if r['category'] == '活动单'],
                              output_mi, output_di)
            daily = defaultdict(lambda: {name: Decimal(0) for name in categories})
            for r in records:
                amounts = daily[r['date']]
                if r['category'] in categories:
                    amounts[r['category']] += r['amount']
            write_summary(summary, categories, daily)
            dates = [r['date'] for r in records]
            start, end = min(dates), max(dates)
            # 使用实际配送日期范围，包含年份，跨月或跨年不产生歧义。
            safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', school).rstrip(' .')[:70] or '未命名学校'
            filename = f'{index:02d}_{safe_name}_{start:%Y.%m.%d}-{end:%Y.%m.%d}对账明细.xlsx'
            school_file = f'学校版/{filename}'
            out.save(run_dir / school_file)
            report = dict(school=school, rows=len(records), file=school_file,
                          total=str(sum((r['amount'] for r in records), Decimal(0))),
                          pending_rows=sum(r['category'] == '待核对' for r in records))
            if include_guotou:
                guotou_file = f'国投版/国投_{filename}'
                append_guotou_sheets(out, output_headers, records, output_mi, output_di)
                out.save(run_dir / guotou_file)
                report['guotou_file'] = guotou_file
            out.close()
            school_reports.append(report)
            if progress:
                progress(f'已生成 {index}/{len(groups)}：{school}')
        with (run_dir / '待核对记录.csv').open('w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['源行号', '客户名', '收货地址', '配送日期', '金额', '原因', '对象名称', '分类字段'])
            for row in exceptions:
                writer.writerow([("'" + str(v)) if str(v).startswith(('=', '+', '-', '@')) else v for v in row])
        result = dict(directory=str(run_dir), schools=len(groups), exceptions=len(exceptions),
                      skipped_totals=skipped, source_amount=f'{source_amount:.2f}',
                      classified_amount=f'{source_amount-pending:.2f}', pending_amount=f'{pending:.2f}',
                      source=str(source), amount_column=amount_column, school_details=school_reports,
                      object_column=get_column_letter(oi + 1) if oi is not None else None,
                      source_sheet=source_sheet, address_column=get_column_letter(ai + 1),
                      customer_column=get_column_letter(ni + 1), report_html=str(run_dir / '核对报告.html'),
                      exception_details=[dict(source_row=r[0], source_sheet=source_sheet, school=r[1],
                                              address=r[2], date=r[3], amount=r[4], reason=r[5],
                                              object_name=r[6], classification_field=r[7]) for r in exceptions],
                      note='有对象名称列时严格按其值分类（教师/老师、学生、活动单），仅缺少该列时按收货地址分类；每校有已分类活动单时增加活动单明细和汇总列；待核对记录保留在底稿。空客户名仅在待核对清单。金额按每行四舍五入到分。')
        (run_dir / '运行报告.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        write_report(result)
        return result
    except Exception:
        shutil.rmtree(run_dir)
        raise
