import os
import queue
import threading
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from bill_splitter.core import export_bills
from bill_splitter.version import VERSION
from bill_splitter.updater import check_update, download_update, install_update
import sys


def main():
    root = tk.Tk()
    root.title('wuli娟娟 v' + VERSION)
    root.geometry('900x650')
    root.minsize(780, 600)
    root.configure(background='#fff3f7')
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TFrame', background='#fff3f7')
    style.configure('TLabel', font=('Microsoft YaHei UI', 11), background='#fff3f7', foreground='#624653')
    style.configure('Title.TLabel', background='#fff3f7', foreground='#b23b70', font=('Microsoft YaHei UI', 28, 'bold'))
    style.configure('Subtitle.TLabel', foreground='#886775', font=('Microsoft YaHei UI', 10))
    style.configure('TButton', font=('Microsoft YaHei UI', 10), padding=(14, 9), background='#f9dce8', foreground='#79344f', borderwidth=0)
    style.map('TButton', background=[('disabled', '#f2e8ec'), ('pressed', '#ecc0d3'), ('active', '#f3c8da')], foreground=[('disabled', '#a18a94')])
    style.configure('Primary.TButton', background='#c8447c', foreground='white', font=('Microsoft YaHei UI', 10, 'bold'))
    style.map('Primary.TButton', background=[('disabled', '#ead2dd'), ('pressed', '#a92f63'), ('active', '#b6386e')], foreground=[('disabled', '#957784'), ('!disabled', 'white')])
    style.configure('TEntry', fieldbackground='white', foreground='#503c46', padding=8, bordercolor='#edc8d8', lightcolor='#edc8d8', darkcolor='#edc8d8')
    style.configure('TCombobox', padding=7, foreground='#503c46', arrowcolor='#ad4772')
    style.map('TCombobox', fieldbackground=[('readonly', 'white')], selectbackground=[('readonly', '#f6dce7')], selectforeground=[('readonly', '#503c46')])
    panel = ttk.Frame(root, padding=24)
    panel.pack(fill='both', expand=True)
    heading = ttk.Frame(panel)
    heading.pack(fill='x')
    ttk.Label(heading, text='wuli娟娟', style='Title.TLabel').pack(side='left')
    update_button = ttk.Button(heading, text='v' + VERSION + ' · 检查更新')
    update_button.pack(side='right')
    ttk.Label(panel, text='把繁琐的账单，整理得明明白白。', style='Subtitle.TLabel').pack(anchor='w', pady=(5, 4))
    ttk.Label(panel, text='学校账单自动分类  ·  每校四张表，有活动单时五张表  ·  本地处理', style='Subtitle.TLabel').pack(anchor='w', pady=(0, 20))
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

    field('① 源账单', source, lambda: source.set(filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')]) or source.get()))
    field('② 保存到', output, lambda: output.set(filedialog.askdirectory() or output.get()))
    options = ttk.Frame(panel)
    options.pack(fill='x', pady=10)
    ttk.Label(options, text='③ 金额口径', width=10).pack(side='left')
    ttk.Combobox(options, textvariable=amount, values=['销售金额(元)', '下单金额(元)'], state='readonly').pack(side='left')
    ttk.Label(panel, text='有活动单：按对象名称分为老师 / 学生 / 活动单').pack(anchor='w', pady=5)
    ttk.Label(panel, text='无活动单：按收货地址分类；分类冲突或未命中时标注源行号供核对。', style='Subtitle.TLabel').pack(anchor='w', pady=5)
    # 先为操作栏保留空间，防止日志区域在小窗口或高 DPI 下挤掉按钮。
    actions = ttk.Frame(panel)
    actions.pack(side='bottom', fill='x', pady=(8, 0))
    log = tk.Text(panel, height=9, font=('Microsoft YaHei UI', 10), wrap='word', state='disabled',
                  background='#fffafd', foreground='#624653', relief='flat', borderwidth=0,
                  highlightthickness=1, highlightbackground='#efcfdd', highlightcolor='#d76a97',
                  padx=14, pady=12, selectbackground='#f4c8dc')
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
        update_button.configure(state='disabled')
        append('正在检查源数据并生成文件……')
        threading.Thread(target=worker, args=(source.get(), output.get(), amount.get()), daemon=True).start()

    def update_worker(download=None):
        try:
            events.put(('downloaded', download_update(download)) if download else ('update', check_update()))
        except Exception as error:
            events.put(('error', '检查或下载更新失败：' + str(error)))

    def begin_update():
        button.configure(state='disabled')
        update_button.configure(state='disabled')
        append('正在检查 GitHub 正式版本……')
        threading.Thread(target=update_worker, daemon=True).start()

    update_button.configure(command=begin_update)

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
                    update_button.configure(state='normal')
                    report_button.configure(state='normal')
                    messagebox.showinfo('生成完成', '点击“查看核对报告”，可查看关键金额和待核对源行号。Excel 打开后会自动计算汇总公式。')
                elif kind == 'update':
                    if payload and messagebox.askyesno('发现新版本', f"当前 v{VERSION}，最新 {payload['version']}。是否下载？\nEXE 版本下载完成后将关闭并重启。"):
                        append('正在下载新版本并校验文件……')
                        threading.Thread(target=update_worker, args=(payload,), daemon=True).start()
                    else:
                        button.configure(state='normal')
                        update_button.configure(state='normal')
                        if not payload:
                            messagebox.showinfo('检查更新', '当前已经是最新正式版本。')
                elif kind == 'downloaded':
                    if getattr(sys, 'frozen', False):
                        try:
                            install_update(payload)
                        except Exception as error:
                            button.configure(state='normal')
                            update_button.configure(state='normal')
                            messagebox.showerror('更新失败', str(error))
                        else:
                            root.destroy()
                            return
                    else:
                        button.configure(state='normal')
                        update_button.configure(state='normal')
                        os.startfile(str(payload.parent))
                        messagebox.showinfo('下载完成', '当前为源码运行模式。已打开下载目录，请运行其中的 BillSplitter.exe。')
                else:
                    append('失败：' + payload)
                    button.configure(state='normal')
                    update_button.configure(state='normal')
                    messagebox.showerror('未完成', payload)
        except queue.Empty:
            pass
        root.after(150, poll)

    button = ttk.Button(actions, text='生成全部学校账单', style='Primary.TButton', command=start)
    button.pack(side='left')
    report_button = ttk.Button(actions, text='查看核对报告', state='disabled',
                              command=lambda: webbrowser.open((Path(result_path.get()) / '核对报告.html').as_uri()))
    report_button.pack(side='left', padx=10)
    ttk.Button(actions, text='打开生成文件夹', command=lambda: os.startfile(result_path.get()) if result_path.get() else None).pack(side='right')
    append('欢迎使用 wuli娟娟。选好账单和保存位置，就可以开始整理啦。')
    root.after(150, poll)
    root.mainloop()


if __name__ == '__main__':
    main()
