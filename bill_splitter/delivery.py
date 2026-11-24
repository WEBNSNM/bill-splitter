"""销售明细转配送单；模板内置，输入及参考文件只读。"""
import json
import re
import shutil
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from math import ceil
from pathlib import Path
from unicodedata import east_asian_width
from uuid import uuid4

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from .core import DISCOUNT, delivery_date, product_category, safe_text, settlement
HEADERS = ['分组', '配送日期', '订单编号', '客户名', '分类对象', '序号', '品类', '商品名称', '单位', '配送量', '基准价', '折扣率', '结算单价', '结算金额', '备注']
# Excel's column-width dialog excludes the padding stored in the workbook.
COLUMN_WIDTHS = {'G': 9.625, 'H': 34.625, 'I': 12, 'L': 12, 'M': 9.625, 'N': 9.625}


def wrapped_lines(value, capacity):
    return sum(max(1, ceil(sum(2 if east_asian_width(c) in ('W', 'F') else 1
                              for c in line) / capacity))
               for line in str(value or '').split('\n'))


def number(value, line, label):
    try:
        result = Decimal(str(value).replace(',', '').strip())
        if not result.is_finite():
            raise InvalidOperation
        return result
    except (InvalidOperation, ValueError):
        raise ValueError(f'第 {line} 行{label}无效或公式缺少缓存值：{value!r}') from None


def write_sheet(ws, school, day, obj, records):
    order = records[0]['order']
    ws.merge_cells('F1:O1')
    ws['F1'] = '成都鸿康农业科技开发有限公司（食材配送清单）'
    ws['F2'] = '客户：'
    ws.merge_cells('G2:H2')
    ws['G2'] = '成都市金牛国投食材配送有限公司'
    ws['I2'] = '订单编号：'
    ws.merge_cells('J2:K2')
    safe_text(ws['J2'], order)
    ws['L2'] = '配送日期：'
    ws.merge_cells('M2:O2')
    ws['M2'] = day.isoformat()
    ws.merge_cells('F3:H3')
    safe_text(ws['F3'], f'地址：{school}（{obj}）')
    ws.merge_cells('I3:K3')
    ws['I3'] = '联系人：陈粤'
    ws.merge_cells('L3:O3')
    ws['L3'] = '联系电话：13056667608'
    for col, label in enumerate(HEADERS, 1):
        ws.cell(4, col, label)
    total = Decimal(0)
    for seq, record in enumerate(records, 1):
        unit, amount = settlement(record['price'], record['quantity'])
        total += amount
        values = [day.isoformat()+school+obj, day.isoformat(), record['order'], school, obj, seq,
                  record['category'], record['name'], record['unit'], float(record['quantity']),
                  float(record['price']), float(DISCOUNT), float(unit), float(amount), record['note']]
        row = seq + 4
        for col, value in enumerate(values, 1):
            safe_text(ws.cell(row, col), value)
        for col in (11, 13, 14):
            ws.cell(row, col).number_format = '0.00'
        ws.cell(row, 12).number_format = '0.0%'
    end = len(records) + 5
    for row, label in ((end, '每页合计'), (end+1, '整单合计')):
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=7)
        ws.cell(row, 6, label)
        ws.merge_cells(start_row=row, start_column=8, end_row=row, end_column=15)
        ws.cell(row, 8, float(total)).number_format = '0.00'
    signature_row = end + 2
    ws.merge_cells(start_row=signature_row, start_column=6, end_row=signature_row, end_column=8)
    ws.merge_cells(start_row=signature_row, start_column=11, end_row=signature_row, end_column=15)
    ws.cell(signature_row, 6, '送货人：')
    ws.cell(signature_row, 11, '收货人：')
    edge = Side(style='thin', color='595959')
    border = Border(left=edge, right=edge, top=edge, bottom=edge)
    centered = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header = Alignment(horizontal='left', vertical='center', wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name='微软雅黑', size=9 if cell.column == 7 and 5 <= cell.row < end else 10.5,
                             bold=cell.row == 4)
            cell.alignment = header if cell.row in (2, 3, end, end + 1) else centered
            if 4 <= cell.row <= end + 1 and cell.column >= 6:
                cell.border = border
    ws['F1'].font = Font(name='微软雅黑', size=19.5, bold=True)
    for address in ('I2', 'J2', 'L2', 'M2', 'F3'):
        ws[address].alignment = Alignment(horizontal='left', vertical='center', shrink_to_fit=True)
    for address in ('J2', 'M2'):
        ws[address].font = Font(name='微软雅黑', size=9)
    ws.cell(signature_row, 6).alignment = Alignment(horizontal='center', vertical='top')
    ws.cell(signature_row, 11).alignment = Alignment(horizontal='left', vertical='top')
    for row in range(1, ws.max_row+1):
        ws.row_dimensions[row].height = 20
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 30
    ws.row_dimensions[signature_row].height = 100
    for row in range(5, end):
        lines = max(wrapped_lines(ws[f'{col}{row}'].value, capacity)
                    for col, capacity in (('G', 10), ('H', 30), ('I', 10), ('O', 6)))
        ws.row_dimensions[row].height = 20 * lines
    for col in range(1, 16):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = COLUMN_WIDTHS.get(letter, 8)
        if col <= 5:
            ws.column_dimensions[letter].hidden = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'J5'
    ws.print_area = f'F1:O{signature_row}'
    ws.page_margins.top = ws.page_margins.bottom = 1 / 2.54
    ws.page_margins.left = ws.page_margins.right = 1.5 / 2.54
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return total


def export_delivery(source, output_dir, progress=None):
    source = Path(source).resolve()
    if source.suffix.lower() != '.xlsx':
        raise ValueError('仅支持 .xlsx 文件。')
    groups = defaultdict(lambda: defaultdict(list))
    source_stream = source.open('rb')
    try:
        w = load_workbook(source_stream, read_only=True, data_only=True)
    except Exception:
        source_stream.close()
        raise
    count = 0
    excluded = 0
    try:
        ws = w.worksheets[0]
        rows = ws.iter_rows(values_only=True)
        headers = [str(v or '').strip() for v in next(rows, ())]
        required = ['配送日期', '订单编号', '客户名', '分类对象', '品类', '商品名', '计量单位', '销售数量', '销售单价(元)']
        for label in required:
            if headers.count(label) != 1:
                raise ValueError(f'第一行须有唯一表头“{label}”。')
        if headers.count('商品备注') > 1:
            raise ValueError('第一行须有唯一表头“商品备注”。')
        indexes = {label: headers.index(label) for label in headers}
        for line, values in enumerate(rows, 2):
            if all(v is None or v == '' for v in values):
                continue
            def get(label):
                return values[indexes[label]] if label in indexes else None
            category = product_category(get('品类'))
            if category == '米面油':
                excluded += 1
                continue
            school = str(get('客户名') or '').strip()
            if school in ('合计', '总计', '总合计') and not get('配送日期'):
                continue
            obj = str(get('分类对象') or '').strip()
            obj = '教师' if obj == '老师' else obj
            if not school or obj not in ('教师', '学生', '活动单'):
                raise ValueError(f'第 {line} 行客户名为空或分类对象未知/冲突，请修正源文件副本后重试。')
            day = delivery_date(get('配送日期'), line, w.epoch)
            order = get('订单编号')
            if order is None or not str(order).strip() or not str(get('商品名') or '').strip():
                raise ValueError(f'第 {line} 行订单编号或商品名为空。')
            if isinstance(order, float) and order.is_integer():
                order = int(order)
            order = str(order).strip()
            groups[school][day, obj].append(dict(order=order if order.startswith('CG-') else 'CG-'+order,
                category=category, name=get('商品名'),
                unit=get('计量单位'), quantity=number(get('销售数量'), line, '销售数量'),
                price=number(get('销售单价(元)'), line, '销售单价'), note=get('商品备注')))
            count += 1
    finally:
        w.close()
        source_stream.close()
    if not count:
        raise ValueError(f'源文件没有可生成的配送明细（已排除米面油 {excluded} 行）。')
    directory = Path(output_dir).resolve() / ('配送单_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid4().hex[:8])
    directory.mkdir(parents=True)
    total = Decimal(0)
    sheets = 0
    try:
        for index, (school, batches) in enumerate(groups.items(), 1):
            book = Workbook()
            book.remove(book.active)
            for (day, obj), records in sorted(batches.items(), key=lambda item: (item[0][0], ('教师', '学生', '活动单').index(item[0][1]))):
                total += write_sheet(book.create_sheet(day.isoformat()+obj), school, day, obj, records)
                sheets += 1
            filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', school).strip(' .')[:100] or '学校'
            book.save(directory / f'{index:03d}_{filename}_配送单.xlsx')
            book.close()
            if progress:
                progress(f'已生成：{school}，{len(batches)} 张配送单')
        result = dict(directory=str(directory), schools=len(groups), sheets=sheets, rows=count, excluded_rows=excluded, total=str(total))
        (directory / '配送运行报告.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        return result
    except Exception:
        shutil.rmtree(directory)
        raise
