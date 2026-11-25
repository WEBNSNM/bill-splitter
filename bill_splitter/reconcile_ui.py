"""供销核对窗口。"""
import json
import os
from queue import Empty, Queue
from threading import Thread
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
    matching = False
    results = Queue()
    alias_path = Path(os.environ.get('LOCALAPPDATA') or Path.home()) / '案甚财务' / '供销别名.json'
    files = [('Excel 工作簿', '*.xlsx')]

    # Reserve space for review controls before the tables claim the remaining height.
    footer = ttk.Frame(body)
    footer.pack(side='bottom', fill='x', pady=(6, 0))
    review_actions = ttk.Frame(footer)
    review_actions.pack(fill='x')
    step_font = ('Microsoft YaHei UI', 11, 'bold')
    ttk.Label(footer, text='勾选后，确认的学校简称和商品异名会保存在本机，供同一供应商下次匹配使用。').pack(anchor='w')

    loading = None

    def position_loading(event=None):
        if loading is None or (event is not None and event.widget is not window):
            return
        shade, card, _ = loading
        x, y = window.winfo_rootx(), window.winfo_rooty()
        width, height = window.winfo_width(), window.winfo_height()
        shade.geometry(f'{width}x{height}+{x}+{y}')
        card_width, card_height = card.winfo_reqwidth(), card.winfo_reqheight()
        card.geometry(f'+{x + (width - card_width) // 2}+{y + (height - card_height) // 2}')

    def show_loading(title):
        nonlocal loading
        shade = tk.Toplevel(window, background='#44313a')
        shade.withdraw()
        shade.overrideredirect(True)
        shade.transient(window)
        shade.attributes('-alpha', 0.28)
        card = tk.Toplevel(shade)
        card.withdraw()
        card.overrideredirect(True)
        card.transient(shade)
        content = tk.Frame(card, background='#ffffff', padx=32, pady=22,
                           highlightthickness=1, highlightbackground='#e7c9d6')
        content.pack()
        tk.Label(content, text=title, background='#ffffff', foreground='#563749',
                 font=step_font).pack()
        bar = ttk.Progressbar(content, mode='indeterminate', length=176)
        bar.pack(pady=(12, 0))
        window.update_idletasks()
        loading = shade, card, bar
        position_loading()
        shade.deiconify()
        card.deiconify()
        card.lift()
        bar.start(12)
        window.update_idletasks()
        card.grab_set()
        card.focus_set()

    def hide_loading():
        nonlocal loading
        shade, card, bar = loading
        loading = None
        bar.stop()
        card.grab_release()
        shade.destroy()

    window.bind('<Configure>', position_loading, add='+')

    def field(label, value, choose):
        row = ttk.Frame(body)
        row.pack(fill='x', pady=3)
        ttk.Label(row, text=label, width=13).pack(side='left')
        ttk.Entry(row, textvariable=value).pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='选择', command=choose).pack(side='left', padx=8)

    field('销售明细', sales_path, lambda: sales_path.set(
        filedialog.askopenfilename(parent=window, filetypes=files) or sales_path.get()))
    field('保存到', output, lambda: output.set(filedialog.askdirectory(parent=window) or output.get()))

    supplier_section = ttk.LabelFrame(body, text='供应商表', padding=(8, 4))
    supplier_section.pack(fill='x', pady=5)
    ttk.Label(supplier_section, text='可一次选择多份 Excel。选中已添加的文件，可修改用于导出和保存别名的供应商名称。').pack(anchor='w')
    supplier_bar = ttk.Frame(supplier_section)
    supplier_bar.pack(fill='x', pady=5)
    vendor_list = tk.Listbox(supplier_bar, height=3, exportselection=False)
    vendor_list.pack(side='left', fill='x', expand=True)

    def refresh_vendors():
        vendor_list.delete(0, 'end')
        for path, name in vendors.items():
            vendor_list.insert('end', f'{name}  |  {Path(path).name}')
        if not vendors:
            vendor_list.insert('end', '尚未添加供应商表')

    def add_vendors():
        for path in filedialog.askopenfilenames(parent=window, filetypes=files):
            vendors.setdefault(path, supplier_name(path))
        refresh_vendors()

    def change_vendor(remove=False):
        if not vendors or not vendor_list.curselection():
            messagebox.showwarning('请选择供应商表', '请先在列表中选中一份供应商表。', parent=window)
            return
        path = list(vendors)[vendor_list.curselection()[0]]
        if remove:
            del vendors[path]
        else:
            name = simpledialog.askstring('供应商名称', '供应商名称', initialvalue=vendors[path], parent=window)
            if name and name.strip():
                vendors[path] = name.strip()
        refresh_vendors()

    ttk.Button(supplier_bar, text='选择供应商 Excel', command=add_vendors).pack(side='left')
    ttk.Button(supplier_bar, text='修改供应商名称', command=change_vendor).pack(side='left')
    ttk.Button(supplier_bar, text='移除', command=lambda: change_vendor(True)).pack(side='left')
    refresh_vendors()
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

    ttk.Label(body, text='1. 选择采购记录', font=step_font).pack(anchor='w')
    purchases = table(('供应商', '采购行', '单据日期', '单据备注', '采购商品', '数量 / 单位', '规格', '结果'),
                      (100, 65, 95, 155, 230, 85, 90, 125), 2)
    ttk.Label(body, textvariable=details, wraplength=810).pack(anchor='w')
    ttk.Label(body, text='2. 选择销售候选', font=step_font).pack(anchor='w')
    candidates = table(('销售行', '配送日期', '客户名', '销售商品', '数量 / 单位', '规格', '匹配依据'),
                       (65, 95, 220, 220, 85, 90, 360), 2)

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
            elif review.matches.get(key) == sale.row:
                reason += '、已匹配本条采购'
            candidates.insert('', 'end', iid=str(sale.row),
                              values=(sale.row, sale.day, sale.school, sale.name,
                                      f'{sale.quantity} {sale.unit}', sale.spec, reason))
        matched = review.matches.get(key)
        if matched is not None and candidates.exists(str(matched)):
            candidates.selection_set(str(matched))
            candidates.see(str(matched))

    def analyze():
        nonlocal matching
        if matching:
            return
        try:
            if not sales_path.get() or not vendors:
                raise ValueError('请选择销售明细和至少一份供应商表。')
            horizon = int(days.get())
            if not 0 <= horizon <= 30:
                raise ValueError('配送日期范围须为 0 至 30 天。')
        except ValueError as error:
            messagebox.showerror('无法核对', str(error), parent=window)
            return
        matching = True
        start_button.configure(text='正在匹配…', state='disabled')
        export_button.state(['disabled'])
        confirm_button.state(['disabled'])
        clear_button.state(['disabled'])
        show_loading('正在匹配…')
        Thread(target=match_sources, args=(sales_path.get(), tuple(vendors.items()), horizon),
               daemon=True).start()
        window.after(80, finish_matching)

    def match_sources(path, vendor_files, horizon):
        try:
            sales = read_source(path, sales=True)
            suppliers = [(name, read_source(file)) for file, name in vendor_files]
            aliases = json.loads(alias_path.read_text(encoding='utf-8')) if alias_path.exists() else {}
            results.put((Reconciliation(sales, suppliers, horizon, aliases), None))
        except Exception as error:
            results.put((None, error))

    def finish_matching():
        nonlocal review, matching
        try:
            review, error = results.get_nowait()
        except Empty:
            window.after(80, finish_matching)
            return
        matching = False
        hide_loading()
        start_button.configure(text='开始匹配', state='normal')
        confirm_button.state(['!disabled'])
        clear_button.state(['!disabled'])
        purchases.delete(*purchases.get_children())
        if error is not None:
            show_candidates()
            status.set('匹配失败，请检查文件后重试。')
            messagebox.showerror('无法核对', str(error), parent=window)
            return
        export_button.state(['!disabled'])
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

    def notify(dialog, title, message):
        dialog(title, message, parent=window)

    def confirm(clear=False):
        key = current_key()
        selected = candidates.selection()
        if review is None:
            notify(messagebox.showwarning, '请先开始匹配', '请选择销售明细和供应商表，再点「开始匹配」。')
            return
        if key is None:
            notify(messagebox.showwarning, '请选择采购记录', '请先在上表选中一条采购记录。')
            return
        if not clear and not selected:
            message = ('请在下表选中一条销售候选。' if review.candidates[key]
                       else '这条采购没有销售候选，可留空并继续核对下一条。')
            notify(messagebox.showwarning, '无法确认匹配', message)
            candidates.focus_set()
            return
        sale_row = None if clear else int(selected[0])
        try:
            review.confirm(key, sale_row, remember=remember.get())
        except ValueError as error:
            notify(messagebox.showwarning, '无法确认匹配', str(error))
            return
        update_purchase(key)
        show_candidates()
        if clear:
            notify(messagebox.showinfo, '已留空', f'采购第 {key[1]} 行已留空。')
            return
        message = f'已确认采购第 {key[1]} 行对应销售第 {sale_row} 行。'
        try:
            if remember.get():
                save_aliases(alias_path, review.aliases)
        except OSError as error:
            notify(messagebox.showwarning, '匹配成功，别名未保存', f'{message}\n别名无法保存：{error}')

    def export():
        if review is None or not output.get():
            messagebox.showwarning('无法导出', '请先核对并选择保存目录。', parent=window)
            return
        for button in (start_button, export_button, confirm_button, clear_button):
            button.state(['disabled'])
        export_button.configure(text='正在导出…')
        destination = output.get()
        show_loading('正在导出…')
        Thread(target=write_export, args=(review, destination), daemon=True).start()
        window.after(80, finish_export)

    def write_export(current_review, destination):
        try:
            results.put((export_reconciliation(current_review, destination), None))
        except Exception as error:
            results.put((None, error))

    def finish_export():
        try:
            directory, error = results.get_nowait()
        except Empty:
            window.after(80, finish_export)
            return
        hide_loading()
        export_button.configure(text='导出核对副本')
        for button in (start_button, export_button, confirm_button, clear_button):
            button.state(['!disabled'])
        if error is not None:
            status.set('导出失败，可检查保存位置后重试。')
            messagebox.showerror('导出失败', str(error), parent=window)
            return
        status.set(f'已导出 {len(review.matches)} 行匹配 · {directory}')
        messagebox.showinfo('已导出', str(directory), parent=window)

    start_button = ttk.Button(actions, text='开始匹配', command=analyze, style='Primary.TButton')
    start_button.pack(side='left', padx=6)
    export_button = ttk.Button(actions, text='导出核对副本', command=export, state='disabled')
    export_button.pack(side='right')
    ttk.Checkbutton(review_actions, text='记住此次学校和商品别名', variable=remember).pack(side='left')
    confirm_button = ttk.Button(review_actions, text='3. 确认匹配', command=confirm, style='Primary.TButton')
    confirm_button.pack(side='right')
    clear_button = ttk.Button(review_actions, text='取消匹配并留空', command=lambda: confirm(True))
    clear_button.pack(side='right')
    purchases.bind('<<TreeviewSelect>>', show_candidates)
