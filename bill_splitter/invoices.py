import re
import shutil
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from pypdf import PdfReader


def invoice_name(pdf):
    with pdf.open('rb') as stream:
        pages = PdfReader(stream).pages
        text = pages[-1].extract_text(extraction_mode='layout')
    match = re.search(r'[（(]\s*小\s*写\s*[）)]\s*[¥￥]?\s*(-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?)(?=\s|$)', text)
    if not match:
        raise ValueError('找不到小写金额')
    area = re.split(r'开\s*票\s*人', text[match.end():], maxsplit=1)
    if len(area) != 2 or not re.search(r'备\s*注', area[0]):
        raise ValueError('无法确定备注栏')
    note = re.sub(r'备\s*注\s*[:：]?', '', area[0], count=1)
    note = ''.join(line.strip() for line in note.splitlines())
    note = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', note).strip(' .')
    if not note:
        raise ValueError('备注为空')
    amount = Decimal(match[1].replace(',', ''))
    return f'{note}（{amount:.2f}）'


def plan_invoices(paths):
    groups = defaultdict(list)
    errors = []
    for path in sorted(set(map(Path, paths))):
        if path.suffix.lower() not in {'.pdf', '.ofd', '.xml'}:
            continue
        number = next((part for part in path.stem.split('_') if part.isdecimal()), None)
        if number:
            groups[number].append(path)
        else:
            errors.append(f'{path.name}：缺少发票号')
    plans = []
    names = set()
    for number, files in groups.items():
        types = [path.suffix.lower() for path in files]
        try:
            if '.pdf' not in types or len(set(types)) != len(types):
                raise ValueError('缺少 PDF 或同一格式有多份文件')
            stem = invoice_name(files[types.index('.pdf')])
            if stem.casefold() in names:
                stem += f'_{number}'
            names.add(stem.casefold())
            plans.append([(path, stem + path.suffix) for path in files])
        except Exception as exc:
            errors.append(f'{number}：{exc}')
    return plans, errors


def export_group(group, directory):
    created = []
    try:
        for source, name in group:
            dest = Path(directory) / name
            with source.open('rb') as reader, dest.open('xb') as writer:
                created.append(dest)
                shutil.copyfileobj(reader, writer)
    except OSError:
        for path in created:
            path.unlink(missing_ok=True)
        raise
