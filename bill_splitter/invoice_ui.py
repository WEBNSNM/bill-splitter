from tkinter import Toplevel, filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from .invoices import export_group, plan_invoices


def open_invoices(parent):
    paths = filedialog.askopenfilenames(parent=parent, title='选择发票 PDF 及配套文件',
                                        filetypes=[('电子发票', '*.pdf *.ofd *.xml')])
    if not paths:
        return
    plans, errors = plan_invoices(paths)
    win = Toplevel(parent)
    win.title('发票重命名预览')
    win.geometry('880x520')
    bar = ttk.Frame(win, padding=8)
    bar.pack(side='bottom', fill='x')
    preview = ScrolledText(win, wrap='word', padx=10, pady=10)
    preview.pack(fill='both', expand=True)
    for group in plans:
        for source, name in group:
            preview.insert('end', f'{source.name}\n  → {name}\n')
        preview.insert('end', '\n')
    for error in errors:
        preview.insert('end', f'跳过：{error}\n')
    preview.configure(state='disabled')

    def export():
        folder = filedialog.askdirectory(parent=win, title='选择重命名副本的保存位置')
        if not folder:
            return
        failures = []
        for group in plans:
            try:
                export_group(group, folder)
            except OSError as exc:
                failures.append(f'{group[0][1]}：{exc}')
        summary = f'已导出 {len(plans) - len(failures)} 组发票。'
        if failures:
            messagebox.showwarning('导出结果', summary + '\n' + '\n'.join(failures), parent=win)
        else:
            messagebox.showinfo('导出结果', summary, parent=win)

    ttk.Label(bar, text=f'可导出 {len(plans)} 组，跳过 {len(errors)} 项').pack(side='left')
    ttk.Button(bar, text='导出重命名副本', command=export,
               state='normal' if plans else 'disabled').pack(side='right')
