"""供销核对窗口。"""
import json
import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from .reconcile import Reconciliation, export_reconciliation, read_source, save_aliases, supplier_name


def open_reconciliation(root):
    window = tk.Toplevel(root)
    window.title('供销核对')
    window.geometry('1100x690')
    window.minsize(850, 570)
    body = ttk.Frame(window, padding=16)
    body.pack(fill='both', expand=True)
    sales_path = tk.StringVar()
    output = tk.StringVar()
    days = tk.StringVar(value='3')
    status = tk.StringVar()
    details = tk.StringVar()
    remember = tk.BooleanVar()
    vendors = {}
    review = None
    alias_path = Path(os.environ.get('LOCALAPPDATA') or Path.home()) / '案甚财务' / '供销别名.json'
    files = [('Excel 工作簿', '*.xlsx')]

    def field(label, value, choose):
        row = ttk.Frame(body)
        row.pack(fill='x', pady=3)
        ttk.Label(row, text=label, width=13).pack(side='left')
        ttk.Entry(row, textvariable=value).pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='选择', command=choose).pack(side='left', padx=8)

    field('销售明细', sales_path, lambda: sales_path.set(
        filedialog.askopenfilename(parent=window, filetypes=files) or sales_path.get()))
    field('保存到', output, lambda: output.set(filedialog.askdirectory(parent=window) or output.get()))

    supplier_bar = ttk.Frame(body)
    supplier_bar.pack(fill='x', pady=5)
    vendor_list = tk.Listbox(supplier_bar, height=3, exportselection=False)
    vendor_list.pack(side='left', fill='x', expand=True)

    def refresh_vendors():
        vendor_list.delete(0, 'end')
        for path, name in vendors.items():
            vendor_list.insert('end', f'{name}  |  {Path(path).name}')

    def add_vendors():
        for path in filedialog.askopenfilenames(parent=window, filetypes=files):
            vendors.setdefault(path, supplier_name(path))
        refresh_vendors()

    def change_vendor(remove=False):
        if not vendor_list.curselection():
            return
        path = list(vendors)[vendor_list.curselection()[0]]
        if remove:
            del vendors[path]
        else:
            name = simpledialog.askstring('供应商名称', '供应商名称', initialvalue=vendors[path], parent=window)
            if name and name.strip():
                vendors[path] = name.strip()
        refresh_vendors()

    ttk.Button(supplier_bar, text='添加供应商表', command=add_vendors).pack(side='left')
    ttk.Button(supplier_bar, text='修改名称', command=change_vendor).pack(side='left')
    ttk.Button(supplier_bar, text='移除', command=lambda: change_vendor(True)).pack(side='left')
    actions = ttk.Frame(body)
    actions.pack(fill='x', pady=5)
    ttk.Label(actions, text='采购后最多天数').pack(side='left')
    ttk.Spinbox(actions, from_=0, to=30, width=5, textvariable=days).pack(side='left', padx=6)
    ttk.Label(body, textvariable=status, wraplength=810).pack(anchor='w')

    def table(columns, widths, height):
        frame = ttk.Frame(body)
        frame.pack(fill='both', expand=True, pady=4)
        tree = ttk.Treeview(frame, columns=columns, show='headings', height=height, selectmode='browse')
        for label, width in zip(columns, widths):
            tree.heading(label, text=label)
            tree.column(label, width=width, minwidth=40)
        vertical = ttk.Scrollbar(frame, command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        vertical.pack(side='right', fill='y')
        horizontal.pack(side='bottom', fill='x')
        tree.pack(fill='both', expand=True)
        return tree

    purchases = table(('供应商', '采购行', '单据日期', '单据备注', '采购商品', '数量 / 单位', '规格', '结果'),
                      (100, 65, 95, 155, 230, 85, 90, 125), 6)
    ttk.Label(body, textvariable=details, wraplength=810).pack(anchor='w')
    candidates = table(('销售行', '配送日期', '客户名', '销售商品', '数量 / 单位', '规格', '匹配依据'),
                       (65, 95, 220, 220, 85, 90, 360), 5)
    review_actions = ttk.Frame(body)
    review_actions.pack(fill='x')

    def current_key():
        selection = purchases.selection()
        return tuple(map(int, selection[0].split(':'))) if selection else None

    def update_purchase(key):
        reason = review.reasons[key]
        if key in review.matches:
            reason += f' · 行 {review.matches[key]}'
        purchases.set(f'{key[0]}:{key[1]}', '结果', reason)

    def show_candidates(_event=None):
        candidates.delete(*candidates.get_children())
        key = current_key()
        if review is None or key is None:
            details.set('')
            return
        purchase = review.lines[key]
        source = review.suppliers[key[0]][1]
        options = review.candidates[key]
        details.set(f'{source.path.name} / {source.sheet} · {source.date_range} · '
                    f'采购第 {purchase.row} 行 · {len(options)} 条销售候选')
        for item in options:
            sale = item.sale
            reason = item.reason
            if sale.row in review.matches.values() and review.matches.get(key) != sale.row:
                reason += '、已有归属'
            candidates.insert('', 'end', iid=str(sale.row),
                              values=(sale.row, sale.day, sale.school, sale.name,
                                      f'{sale.quantity} {sale.unit}', sale.spec, reason))

    def analyze():
        nonlocal review
        try:
            if not sales_path.get() or not vendors:
                raise ValueError('请选择销售明细和至少一份供应商表。')
            horizon = int(days.get())
            if not 0 <= horizon <= 30:
                raise ValueError('配送日期范围须为 0 至 30 天。')
            sales = read_source(sales_path.get(), sales=True)
            suppliers = [(name, read_source(path)) for path, name in vendors.items()]
            aliases = json.loads(alias_path.read_text(encoding='utf-8')) if alias_path.exists() else {}
            review = Reconciliation(sales, suppliers, horizon, aliases)
        except (OSError, ValueError, KeyError) as error:
            messagebox.showerror('无法核对', str(error), parent=window)
            return
        purchases.delete(*purchases.get_children())
        for key, line in review.lines.items():
            purchases.insert('', 'end', iid=f'{key[0]}:{key[1]}',
                             values=(review.suppliers[key[0]][0], line.row, line.day, line.school,
                                     line.name, f'{line.quantity} {line.unit}', line.spec, ''))
            update_purchase(key)
        status.set(f'销售：{review.sales.path.name} / {review.sales.sheet} · {review.sales.date_range} · '
                   f'自动匹配 {len(review.matches)}/{len(review.lines)} 行')
        first = next(iter(purchases.get_children()), None)
        if first:
            purchases.selection_set(first)
        show_candidates()

    def confirm(clear=False):
        key = current_key()
        selected = candidates.selection()
        if review is None or key is None or not (clear or selected):
            return
        try:
            review.confirm(key, None if clear else int(selected[0]), remember=remember.get())
            if remember.get() and not clear:
                save_aliases(alias_path, review.aliases)
        except (ValueError, OSError) as error:
            messagebox.showerror('无法确认', str(error), parent=window)
        update_purchase(key)
        show_candidates()

    def export():
        if review is None or not output.get():
            messagebox.showwarning('无法导出', '请先核对并选择保存目录。', parent=window)
            return
        try:
            directory = export_reconciliation(review, output.get())
        except (OSError, ValueError) as error:
            messagebox.showerror('导出失败', str(error), parent=window)
            return
        status.set(f'已导出 {len(review.matches)} 行匹配 · {directory}')
        messagebox.showinfo('已导出', str(directory), parent=window)

    ttk.Button(actions, text='开始匹配', command=analyze).pack(side='left', padx=6)
    ttk.Button(actions, text='导出核对副本', command=export).pack(side='right')
    ttk.Checkbutton(review_actions, text='记住此次学校和商品别名', variable=remember).pack(side='left')
    ttk.Button(review_actions, text='确认所选销售行', command=confirm).pack(side='right')
    ttk.Button(review_actions, text='取消匹配并留空', command=lambda: confirm(True)).pack(side='right')
    purchases.bind('<<TreeviewSelect>>', show_candidates)
