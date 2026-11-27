"""供应商采购与销售核对。源工作簿只读，结果写入独立批次。"""
import colorsys
import json
import re
import shutil
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from hashlib import file_digest
from itertools import islice
from pathlib import Path
from uuid import uuid4

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .core import delivery_date, money, safe_text
from .delivery import number

SUPPLIER_COLUMNS = (
    ('单据日期', '采购日期', '送货日期', '日期'),
    ('单据备注', '学校', '学校名称', '校区', '客户名', '客户名称', '客户', '备注'),
    ('商品名称', '商品名', '品名', '货品名称', '存货'),
    ('数量', '采购数量', '送货数量'),
    ('单位', '计量单位', '采购单位'),
    ('金额', '采购金额', '金额(元)', '金额（元）', '采购金额(元)', '采购金额（元）'),
)
UNIT_COLUMN = 4
SUPPLIER_HEADER_LIMIT = 30
SALES_COLUMNS = ('配送日期', ('客户名', '客户名称'), ('商品名', '商品名称'),
                 '销售数量', ('计量单位', '单位'), '销售金额(元)')
RESULT_COLUMNS = ('销售名称', '销售金额', '毛利', '毛利率')
PACK_SIZE = re.compile(
    r'(?<![a-z0-9])\d+(?:\.\d+)?\s*'
    r'(?:kg|mg|ml|g|l|千克|公斤|克|毫升|升|斤|包|袋|瓶|盒|件|桶|罐|片|个)(?![a-z])'
    r'(?:\s*[×x*]\s*\d+\s*(?:包|袋|瓶|盒|件|桶|罐|片|个)?)?'
)


def normalize(value):
    return re.sub(r'[\W_]+', '', unicodedata.normalize('NFKC', str(value or '')).casefold())


def product(value):
    return normalize(re.sub(r'^\s*(?:【散】|\[散\])', '', str(value or '')))


def school_note(value):
    return normalize(re.sub(r'附批次|批检|批次|教师|老师|学生|幼儿(?!园)', '', str(value or '')))


def school_keys(value):
    value = unicodedata.normalize('NFKC', str(value))
    campus = re.search(r'\(([^()]+?)(?:校区|园区)\)', value)
    name = re.sub(r'^(?:四川省)?(?:成都市)?(?:金牛区)?', '', value)
    keys = {normalize(value), normalize(name)}
    if campus:
        keys.add(normalize(campus[1]))
    return {key for key in keys if len(key) >= 2}


def product_keywords(left, right):
    words = []
    for i, _, size in SequenceMatcher(None, left, right, autojunk=False).get_matching_blocks():
        word = left[i:i + size]
        if size >= 2 or word in ('面', '麦'):
            words.append(word)
    return words


def keyword_name(value):
    name = unicodedata.normalize('NFKC', str(value or '')).casefold()
    return product(PACK_SIZE.sub('', name))


@dataclass(frozen=True)
class Line:
    row: int
    day: object
    school: str
    name: str
    quantity: object
    unit: str
    amount: object
    spec: str


@dataclass
class Source:
    path: Path
    sheet: str
    columns: tuple
    lines: list
    digest: str
    header_row: int = 1

    @property
    def date_range(self):
        days = [line.day for line in self.lines]
        return f'{min(days)} 至 {max(days)}' if days else '无明细'


class SupplierColumnsError(ValueError):
    def __init__(self, path, message):
        self.path = path
        super().__init__(f'{path.name}：{message}')


@dataclass(frozen=True)
class SupplierMapping:
    sheet: str
    headers: tuple
    columns: tuple
    header_row: int = 1

    def validate(self, path, headers):
        if headers != self.headers:
            raise SupplierColumnsError(path, '表头已变更，请重新设置供应商列对应。')
        selected = []
        for i, (names, column) in enumerate(zip(SUPPLIER_COLUMNS, self.columns, strict=True)):
            if i == UNIT_COLUMN and column is None:
                continue
            if column is None or not 0 <= column < len(headers) or not headers[column]:
                raise SupplierColumnsError(path, f'请选择「{names[0]}」对应的列。')
            selected.append(column)
        if len(set(selected)) != len(selected):
            raise SupplierColumnsError(path, '有重复选用的 Excel 列，请为每个字段选择不同的列。')


def row_headers(row):
    return tuple(str(value or '').strip() for value in row)


def supplier_headers(path):
    with Path(path).open('rb') as stream:
        book = load_workbook(stream, read_only=True, data_only=True)
        try:
            return {ws.title: {i: row_headers(row) for i, row in enumerate(
                        islice(ws.values, SUPPLIER_HEADER_LIMIT), 1)
                        if sum(isinstance(value, str) and bool(value.strip()) for value in row) >= 5}
                    for ws in book.worksheets}
        finally:
            book.close()


def supplier_choices(headers):
    choices = []
    for names in SUPPLIER_COLUMNS:
        exact = tuple(i for i, header in enumerate(headers) if header == names[0])
        choices.append(exact or tuple(i for i, header in enumerate(headers) if header in names[1:]))
    return tuple(choices)


def source_columns(headers, required):
    columns = []
    for field in required:
        names = field if isinstance(field, tuple) else (field,)
        indexes = [index for index, header in enumerate(headers) if header in names]
        if len(indexes) != 1:
            return None
        columns.append(indexes[0])
    return tuple(columns)


def read_source(path, *, sales=False, sheet=None, mapping=None):
    path = Path(path).resolve()
    with path.open('rb') as stream:
        digest = file_digest(stream, 'sha256').hexdigest()
        stream.seek(0)
        book = load_workbook(stream, read_only=True, data_only=True)
        try:
            if mapping is not None:
                sheet = mapping.sheet
                if sheet not in book.sheetnames:
                    raise SupplierColumnsError(path, '已选工作表不存在，请重新设置供应商列对应。')
            for ws in ([book[sheet]] if sheet else book.worksheets):
                rows = ws.values
                if sales:
                    header_row = 1
                    headers = row_headers(next(rows, ()))
                    columns = source_columns(headers, SALES_COLUMNS)
                elif mapping is not None:
                    header_row = mapping.header_row
                    headers = row_headers(next(islice(rows, header_row - 1, None), ()))
                    mapping.validate(path, headers)
                    columns = mapping.columns
                else:
                    columns = None
                    for header_row, row in enumerate(islice(rows, SUPPLIER_HEADER_LIMIT), 1):
                        headers = row_headers(row)
                        choices = supplier_choices(headers)
                        if all(len(options) == 1 or (i == UNIT_COLUMN and not options)
                               for i, options in enumerate(choices)):
                            columns = tuple(options[0] if options else None for options in choices)
                            break
                if columns is None:
                    continue
                spec = headers.index('规格') if '规格' in headers else None
                lines = []
                for index, row in enumerate(rows, header_row + 1):
                    day, school, name, quantity, unit, amount = (
                        row[col] if col is not None else None for col in columns)
                    if not day and (not name or str(name).strip() in ('合计', '总计', '总合计')):
                        continue
                    if not sales:
                        label = str(day or '').strip().rstrip(':：').strip()
                        if label in ('合计', '总计', '总合计', '小计'):
                            continue
                        if label == '制表人' and all(value is None for value in (school, name, quantity, amount)):
                            continue
                    try:
                        lines.append(Line(index, delivery_date(day, index, book.epoch),
                                          str(school or '').strip(), str(name or '').strip(),
                                          number(quantity, index, '数量'), str(unit or '').strip(),
                                          None if not sales and amount in (None, '') else money(amount, index),
                                          str(row[spec] or '').strip() if spec is not None else ''))
                        if not lines[-1].name:
                            raise ValueError(f'第 {index} 行商品名为空。')
                    except ValueError as error:
                        raise ValueError(f'{path.name} / {ws.title}：{error}') from error
                return Source(path, ws.title, columns, lines, digest, header_row)
        finally:
            book.close()
    if not sales:
        raise SupplierColumnsError(path, '无法唯一识别供应商表头，请设置「供应商列对应」。')
    labels = '、'.join(field[0] if isinstance(field, tuple) else field for field in SALES_COLUMNS)
    raise ValueError(f'{path.name} 找不到首行含唯一必需表头的工作表。需要：{labels}')


def supplier_name(path):
    stem = Path(path).stem
    match = re.search(r'[（(]([^()（）]+?)供应商(?:发)?[）)]', stem)
    return match[1].strip() if match else stem


@dataclass(frozen=True)
class Candidate:
    sale: Line
    reason: str
    rank: tuple
    exact: bool


class Reconciliation:
    def __init__(self, sales, suppliers, days, aliases=None):
        self.sales = sales
        self.suppliers = suppliers
        self.days = days
        self.aliases = aliases if aliases is not None else {}
        self.lines = {(i, line.row): line for i, (_, source) in enumerate(suppliers) for line in source.lines}
        self.sale_rows = {line.row: line for line in sales.lines}
        self.matches = {}
        self.candidates = {}
        self.reasons = {}
        self._by_day = defaultdict(list)
        for sale in sales.lines:
            self._by_day[sale.day].append(sale)
        self._sale_names = {sale.row: (product(sale.name), keyword_name(sale.name)) for sale in sales.lines}
        self._schools = {sale.school: school_keys(sale.school) for sale in sales.lines if sale.school}
        proposed = {}
        owners = Counter()
        for key in self.lines:
            candidates, hits = self._find_candidates(key)
            self._set_candidates(key, candidates)
            # Consider competing ownership across every supplier before filling any row.
            owners.update(item.sale.row for item in candidates
                          if item.rank[0] or (item.rank[1] and item.sale.school in hits))
            exact = [item for item in candidates if item.exact]
            if len(exact) == 1:
                proposed[key] = exact[0].sale.row
        for key, sale_row in proposed.items():
            if owners[sale_row] == 1:
                self.matches[key] = sale_row
                self.reasons[key] = '自动匹配'

    def _find_candidates(self, key):
        purchase = self.lines[key]
        vendor = self.suppliers[key[0]][0]
        aliases = self.aliases.get(vendor, {})
        note = school_note(purchase.school)
        known_school = aliases.get('schools', {}).get(note)
        hits = {name for name, keys in self._schools.items() if note and any(k in note for k in keys)}
        if known_school in self._schools:
            hits = {known_school}
        school = len(hits) == 1
        identity = product(purchase.name)
        keywords = keyword_name(purchase.name)
        known_product = aliases.get('products', {}).get(identity)
        candidates = []
        for delta in range(self.days + 1):
            for sale in self._by_day[purchase.day + timedelta(days=delta)]:
                if school and sale.school not in hits:
                    continue
                sale_product, sale_keywords = self._sale_names[sale.row]
                exact = bool(identity and (identity == sale_product or known_product == sale_product))
                words = product_keywords(keywords, sale_keywords) if not exact else []
                if not exact and not words:
                    continue
                score = sum(len(word) ** 2 for word in words)
                unit = not purchase.unit or normalize(purchase.unit) == normalize(sale.unit)
                quantity = purchase.quantity == sale.quantity
                spec = not purchase.spec or not sale.spec or normalize(purchase.spec) == normalize(sale.spec)
                if not purchase.unit:
                    quantity_reason = ('数量一致' if quantity else '数量有差异') + '、采购单位未提供'
                else:
                    quantity_reason = '数量单位一致' if unit and quantity else '数量或单位有差异'
                reason = '、'.join(('校区明确' if school else '校区待核对',
                                  '商品名或别名吻合' if exact else '商品关键词：' + '／'.join(words),
                                  f'晚 {delta} 天' if delta else '同日',
                                  quantity_reason,
                                  '规格一致或缺省' if spec else '规格有差异'))
                candidates.append(Candidate(sale, reason,
                                            (exact, unit and quantity, delta == 1, spec, score, -delta),
                                            school and exact and unit and quantity and spec))
        candidates.sort(key=lambda item: (item.rank[0], item.rank[4], item.rank[1],
                                           item.rank[2], item.rank[3], item.rank[5]), reverse=True)
        return candidates, hits

    def _set_candidates(self, key, candidates):
        same_name = [item for item in candidates if item.rank[0]]
        self.candidates[key] = same_name or candidates
        if key not in self.matches and self.reasons.get(key) != '留空':
            self.reasons[key] = '待确认' if candidates else '无候选'

    def refresh_candidates(self, key):
        candidates, _ = self._find_candidates(key)
        self._set_candidates(key, candidates)

    def remember_aliases(self, key, *, school='', name=''):
        purchase = self.lines[key]
        vendor = self.suppliers[key[0]][0]
        note = school_note(purchase.school)
        identity = product(purchase.name)
        if note and school:
            self.aliases.setdefault(vendor, {}).setdefault('schools', {})[note] = school
        if identity and product(name):
            self.aliases.setdefault(vendor, {}).setdefault('products', {})[identity] = product(name)

    def confirm(self, key, sale_row, *, remember=False):
        purchase = self.lines[key]
        if sale_row is None:
            self.matches.pop(key, None)
            self.reasons[key] = '留空'
            return
        sale = self.sale_rows[sale_row]
        if not 0 <= (sale.day - purchase.day).days <= self.days:
            raise ValueError('所选销售行不在本次配送日期范围内。')
        if sale_row in (row for other, row in self.matches.items() if other != key):
            raise ValueError('这条销售记录已归属其他采购行。请先取消原匹配。')
        if remember:
            self.remember_aliases(key, school=sale.school, name=sale.name)
        self.matches[key] = sale_row
        self.reasons[key] = '人工确认'


def export_reconciliation(review, output):
    directory = Path(output).resolve() / f'供销核对_{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:8]}'
    directory.mkdir(parents=True)
    try:
        vendors = list(dict.fromkeys(name for name, _ in review.suppliers))
        colors = {}
        used = set()
        for index, name in enumerate(vendors):
            hue = index / len(vendors)
            while True:
                rgb = ''.join(f'{round(value * 255):02X}'
                              for value in colorsys.hsv_to_rgb(hue, .35, .98))
                if rgb not in used:
                    break
                hue = (hue + .0031) % 1
            used.add(rgb)
            colors[name] = PatternFill('solid', fgColor='FF' + rgb)
        sale_owners = {}
        for index, (name, source) in enumerate(review.suppliers):
            book = open_reviewed_workbook(source)
            try:
                ws = book[source.sheet]
                start = ws.max_column + 1
                for column, label in enumerate(RESULT_COLUMNS, start):
                    ws.cell(source.header_row, column, label)
                for line in source.lines:
                    row = review.matches.get((index, line.row))
                    if row is None:
                        continue
                    sale = review.sale_rows[row]
                    safe_text(ws.cell(line.row, start), sale.name)
                    ws.cell(line.row, start + 1, float(sale.amount)).number_format = '#,##0.00'
                    if line.amount is not None:
                        profit = sale.amount - line.amount
                        ws.cell(line.row, start + 2, float(profit)).number_format = '#,##0.00'
                        if sale.amount:
                            ws.cell(line.row, start + 3, float(profit / sale.amount)).number_format = '0.00%'
                    sale_owners[row] = name
                book.save(directory / f'{index + 1:02d}_{source.path.stem}_核对.xlsx')
            finally:
                book.close()
        book = open_reviewed_workbook(review.sales)
        try:
            ws = book[review.sales.sheet]
            last = ws.max_column + 1
            ws.cell(1, last, '供应商名称')
            for row, name in sale_owners.items():
                safe_text(ws.cell(row, last), name)
                for col in (review.sales.columns[2] + 1, review.sales.columns[3] + 1, last):
                    ws.cell(row, col).fill = colors[name]
            book.save(directory / '销售明细_核对.xlsx')
        finally:
            book.close()
        return directory
    except Exception:
        shutil.rmtree(directory)
        raise


def open_reviewed_workbook(source):
    with source.path.open('rb') as stream:
        if file_digest(stream, 'sha256').hexdigest() != source.digest:
            raise ValueError(f'{source.path.name} 已修改，请重新匹配后导出。')
        stream.seek(0)
        return load_workbook(stream)


def save_aliases(path, aliases):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(aliases, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary.replace(path)
