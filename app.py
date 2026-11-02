import os
import queue
import threading
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from bill_splitter.core import export_bills


def main():
    root = tk.Tk()
    root.title('学校账单分类工具')
    root.geometry('820x570')
    root.minsize(720, 520)
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TLabel', font=('Microsoft YaHei UI', 11))
    style.configure('TButton', font=('Microsoft YaHei UI', 10), padding=8)
    panel = ttk.Frame(root, padding=24)
    panel.pack(fill='both', expand=True)
    ttk.Label(panel, text='学校账单 · 自动分类', font=('Microsoft YaHei UI', 21, 'bold')).pack(anchor='w')
    ttk.Label(panel, text='读取第一个工作表，每所学校生成四张表；源文件保持不变。').pack(anchor='w', pady=(8, 20))
    source = tk.StringVar()
    output = tk.StringVar()
    amount = tk.StringVar(value='销售金额(元)')
    result_path = tk.StringVar()

    def field(label, variable, choose):
        frame = ttk.Frame(panel)
        frame.pack(fill='x', pady=5)
        ttk.Label(frame, text=label, width=10).pack(side='left')
        ttk.Entry(frame, textvariable=variable).pack(side='left', fill='x', expand=True)
        ttk.Button(frame, text='选择', command=choose).pack(side='left', padx=(10, 0))

    field('源账单', source, lambda: source.set(filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')]) or source.get()))
    field('输出文件夹', output, lambda: output.set(filedialog.askdirectory() or output.get()))
    options = ttk.Frame(panel)
    options.pack(fill='x', pady=10)
    ttk.Label(options, text='汇总金额', width=10).pack(side='left')
    ttk.Combobox(options, textvariable=amount, values=['销售金额(元)', '下单金额(元)'], state='readonly').pack(side='left')
    ttk.Label(panel, text='老师：教师 / 成人 / 职工 / 老师    学生：幼儿 / 学生').pack(anchor='w', pady=5)
    ttk.Label(panel, text='同时命中两类或均未命中 → 待核对，不计入分类总计。').pack(anchor='w', pady=5)
    # 先为操作栏保留空间，防止日志区域在小窗口或高 DPI 下挤掉按钮。
    actions = ttk.Frame(panel)
    actions.pack(side='bottom', fill='x', pady=(8, 0))
    log = tk.Text(panel, height=9, font=('Microsoft YaHei UI', 10), wrap='word', state='disabled')
    log.pack(fill='both', expand=True, pady=10)
    events = queue.Queue()

    def append(text):
        log.configure(state='normal')
        log.insert('end', text + '\n')
        log.see('end')
        log.configure(state='disabled')

    def worker(src, dest, col):
        try:
            result = export_bills(src, dest, col, lambda text: events.put(('log', text)))
            events.put(('done', result))
        except Exception as error:
            events.put(('error', str(error)))

    def start():
        if not source.get() or not output.get():
            messagebox.showwarning('请先选择', '请选择源账单和输出文件夹。')
            return
        button.configure(state='disabled')
        append('正在检查源数据并生成文件……')
        threading.Thread(target=worker, args=(source.get(), output.get(), amount.get()), daemon=True).start()

    def poll():
        try:
            while True:
                kind, payload = events.get_nowait()
                if kind == 'log':
                    append(payload)
                elif kind == 'done':
                    result_path.set(payload['directory'])
                    append(f"完成：{payload['schools']} 所学校，待核对 {payload['exceptions']} 行。\n源金额 {payload['source_amount']} = 已分类 {payload['classified_amount']} + 待核对 {payload['pending_amount']}\n输出：{payload['directory']}")
                    button.configure(state='normal')
                    report_button.configure(state='normal')
                    messagebox.showinfo('生成完成', '点击“查看核对报告”，可查看关键金额和待核对源行号。Excel 打开后会自动计算汇总公式。')
                else:
                    append('失败：' + payload)
                    button.configure(state='normal')
                    messagebox.showerror('未完成', payload)
        except queue.Empty:
            pass
        root.after(150, poll)

    button = ttk.Button(actions, text='生成全部学校账单', command=start)
    button.pack(side='left')
    report_button = ttk.Button(actions, text='查看核对报告', state='disabled',
                              command=lambda: webbrowser.open((Path(result_path.get()) / '核对报告.html').as_uri()))
    report_button.pack(side='left', padx=10)
    ttk.Button(actions, text='打开生成文件夹', command=lambda: os.startfile(result_path.get()) if result_path.get() else None).pack(side='right')
    root.after(150, poll)
    root.mainloop()


if __name__ == '__main__':
    main()
