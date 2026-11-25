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
from pathlib import Path
from uuid import uuid4

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .core import delivery_date, money, safe_text
from .delivery import number

SUPPLIER_COLUMNS = ('单据日期', '单据备注', '商品名称', '数量', '单位', '金额')
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

    @property
    def date_range(self):
        days = [line.day for line in self.lines]
        return f'{min(days)} 至 {max(days)}' if days else '无明细'


def source_columns(headers, required):
    columns = []
    for field in required:
        names = field if isinstance(field, tuple) else (field,)
        indexes = [index for index, header in enumerate(headers) if header in names]
        if len(indexes) != 1:
            return None
        columns.append(indexes[0])
    return tuple(columns)


def read_source(path, *, sales=False, sheet=None):
    path = Path(path).resolve()
    required = SALES_COLUMNS if sales else SUPPLIER_COLUMNS
    with path.open('rb') as stream:
        digest = file_digest(stream, 'sha256').hexdigest()
        stream.seek(0)
        book = load_workbook(stream, read_only=True, data_only=True)
        try:
            for ws in ([book[sheet]] if sheet else book.worksheets):
                rows = ws.values
                headers = [str(v or '').strip() for v in next(rows, ())]
                columns = source_columns(headers, required)
                if columns is None:
                    continue
                spec = headers.index('规格') if '规格' in headers else None
                lines = []
                for index, row in enumerate(rows, 2):
                    day, school, name, quantity, unit, amount = (row[col] for col in columns)
                    if not day and (not name or str(name).strip() in ('合计', '总计', '总合计')):
                        continue
                    try:
                        lines.append(Line(index, delivery_date(day, index, book.epoch),
                                          str(school or '').strip(), str(name or '').strip(),
                                          number(quantity, index, '数量'), str(unit or '').strip(),
                                          money(amount, index), str(row[spec] or '').strip() if spec is not None else ''))
                        if not lines[-1].name:
                            raise ValueError(f'第 {index} 行商品名为空。')
                    except ValueError as error:
                        raise ValueError(f'{path.name} / {ws.title}：{error}') from error
                return Source(path, ws.title, columns, lines, digest)
        finally:
            book.close()
    labels = '、'.join(field[0] if isinstance(field, tuple) else field for field in required)
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
        by_day = defaultdict(list)
        for sale in sales.lines:
            by_day[sale.day].append(sale)
        sale_names = {sale.row: (product(sale.name), keyword_name(sale.name)) for sale in sales.lines}
        schools = {sale.school: school_keys(sale.school) for sale in sales.lines if sale.school}
        proposed = {}
        owners = Counter()
        for key, purchase in self.lines.items():
            vendor = suppliers[key[0]][0]
            aliases = self.aliases.get(vendor, {})
            note = school_note(purchase.school)
            known_school = aliases.get('schools', {}).get(note)
            hits = {name for name, keys in schools.items() if note and any(k in note for k in keys)}
            if known_school in schools:
                hits = {known_school}
            school = len(hits) == 1
            identity = product(purchase.name)
            keywords = keyword_name(purchase.name)
            known_product = aliases.get('products', {}).get(identity)
            candidates = []
            for delta in range(days + 1):
                for sale in by_day[purchase.day + timedelta(days=delta)]:
                    if school and sale.school not in hits:
                        continue
                    sale_product, sale_keywords = sale_names[sale.row]
                    exact = bool(identity and (identity == sale_product or known_product == sale_product))
                    words = product_keywords(keywords, sale_keywords) if not exact else []
                    if not exact and not words:
                        continue
                    score = sum(len(word) ** 2 for word in words)
                    unit = bool(purchase.unit and normalize(purchase.unit) == normalize(sale.unit))
                    quantity = purchase.quantity == sale.quantity
                    spec = not purchase.spec or not sale.spec or normalize(purchase.spec) == normalize(sale.spec)
                    reason = '、'.join(('校区明确' if school else '校区待核对',
                                      '商品名或别名吻合' if exact else '商品关键词：' + '／'.join(words),
                                      f'晚 {delta} 天' if delta else '同日',
                                      '数量单位一致' if unit and quantity else '数量或单位有差异',
                                      '规格一致或缺省' if spec else '规格有差异'))
                    candidates.append(Candidate(sale, reason,
                                                (exact, unit and quantity, delta == 1, spec, score, -delta),
                                                school and exact and unit and quantity and spec))
            candidates.sort(key=lambda item: (item.rank[0], item.rank[4], item.rank[1],
                                               item.rank[2], item.rank[3], item.rank[5]), reverse=True)
            same_name = [item for item in candidates if item.rank[0]]
            self.candidates[key] = same_name or candidates
            # Consider competing ownership across every supplier before filling any row.
            owners.update(item.sale.row for item in candidates
                          if item.rank[0] or (item.rank[1] and item.sale.school in hits))
            exact = [item for item in candidates if item.exact]
            if len(exact) == 1:
                proposed[key] = exact[0].sale.row
            self.reasons[key] = '待确认' if candidates else '无候选'
        for key, sale_row in proposed.items():
            if owners[sale_row] == 1:
                self.matches[key] = sale_row
                self.reasons[key] = '自动匹配'

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
            vendor = self.suppliers[key[0]][0]
            aliases = self.aliases.setdefault(vendor, {})
            note = school_note(purchase.school)
            if note and sale.school:
                aliases.setdefault('schools', {})[note] = sale.school
            if product(purchase.name) and product(sale.name):
                aliases.setdefault('products', {})[product(purchase.name)] = product(sale.name)
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
                    ws.cell(1, column, label)
                for line in source.lines:
                    row = review.matches.get((index, line.row))
                    if row is None:
                        continue
                    sale = review.sale_rows[row]
                    profit = sale.amount - line.amount
                    safe_text(ws.cell(line.row, start), sale.name)
                    ws.cell(line.row, start + 1, float(sale.amount)).number_format = '#,##0.00'
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
