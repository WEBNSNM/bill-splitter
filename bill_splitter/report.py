"""生成无需服务器、无需网络的核对报告。"""
from html import escape
from pathlib import Path
from urllib.parse import quote


def write_report(result):
    def e(value):
        return escape(str(value if value is not None else ''), quote=True)

    cards = ''.join(f'<div class="card"><small>{label}</small><strong>{e(result[key])}</strong></div>' for label, key in [
        ('学校数量', 'schools'), ('待核对行数', 'exceptions'), ('源金额（元）', 'source_amount'),
        ('已分类（元）', 'classified_amount'), ('待核对（元）', 'pending_amount')])
    rows = []
    for item in result['exception_details']:
        location = f"{item['source_sheet']} · Excel 第 {item['source_row']} 行"
        rows.append('<tr>' + ''.join(f'<td>{e(v)}</td>' for v in [location, item['school'], item['address'], item.get('object_name', ''), item.get('classification_field', '收货地址'), item['date'], item['amount'], item['reason']]) + '</tr>')
    schools = ''.join('<tr><td>' + e(s['school']) + '</td><td>' + e(s['rows']) + '</td><td>' + e(s['total']) + '</td><td>' + e(s['pending_rows']) + '</td><td><a href="' + quote(s['file']) + '">打开学校账单</a></td></tr>' for s in result['school_details'])
    document = '''<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>账单核对报告</title><style>
body{font:15px/1.65 "Microsoft YaHei",sans-serif;background:#f3f6fa;color:#21324a;margin:0}main{max-width:1400px;margin:auto;padding:32px}h1{margin:0}small{color:#607086}a{color:#175baf}nav{display:flex;gap:20px;margin:20px 0}.cards{display:flex;gap:12px;flex-wrap:wrap}.card{background:white;border:1px solid #dfe6ef;border-radius:12px;padding:18px;flex:1;min-width:150px}.card strong{display:block;font-size:24px}section{background:white;border:1px solid #dfe6ef;border-radius:12px;padding:20px;margin-top:22px}.scroll{overflow:auto;max-height:650px}table{border-collapse:collapse;width:100%;font-size:14px}th,td{padding:12px;text-align:left;border-bottom:1px solid #e5eaf0;vertical-align:top}th{background:#eaf0f8;position:sticky;top:0}td:first-child{min-width:175px;font-weight:600}input{padding:10px;width:min(90%,500px);border:1px solid #bbc9da;border-radius:6px;margin-bottom:12px}.note{padding:12px;background:#fff5df;border-radius:8px}pre{white-space:pre-wrap;overflow-wrap:anywhere}h2{margin-top:0}
</style><main><h1>账单核对报告</h1><p>本地离线报告 · 无需安装额外软件</p>
<nav><a href="#review">待核对记录</a><a href="#schools">学校账单</a><a href="#data">关键数据</a></nav>
<div class="cards">''' + cards + '</div><section><h2>来源与核对口径</h2><p>源文件：' + e(result['source']) + '</p><p>源工作表：' + e(result['source_sheet']) + '；金额字段：' + e(result['amount_column']) + '</p><p>源金额 = 已分类金额 + 待核对金额。已跳过汇总行：' + e(result['skipped_totals']) + '</p><p class="note">' + e(result['note']) + '''<br>行号为源 Excel 左侧行号，包含第 1 行表头，不是导出文件行号，也不是“序号”列。请在源表使用 Ctrl+G 输入下列单元格定位。</p></section>
<section id="review"><h2>待核对记录</h2><input id="search" placeholder="搜索学校、地址、对象名称、原因或源行号" aria-label="搜索待核对记录"><p id="count"></p><div class="scroll"><table><thead><tr><th>源数据位置</th><th>学校</th><th>收货地址原文</th><th>对象名称原文</th><th>分类字段</th><th>配送日期</th><th>金额（元）</th><th>具体原因</th></tr></thead><tbody id="issues">''' + (''.join(rows) if rows else '<tr><td colspan="8">没有待核对记录。</td></tr>') + '</tbody></table></div><p>定位单元格：收货地址列 ' + e(result['address_column']) + '，对象名称列 ' + e(result.get('object_column') or '无（使用地址分类）') + '，客户名列 ' + e(result['customer_column']) + '''。例如第 20 行地址输入对应列字母加 20。</p></section>
<section id="schools"><h2>学校账单</h2><div class="scroll"><table><thead><tr><th>学校</th><th>明细行数</th><th>全部金额（元）</th><th>待核对行数</th><th>文件</th></tr></thead><tbody>''' + schools + '''</tbody></table></div></section>
<section id="data"><h2>关键数据</h2><p><a href="待核对记录.csv">打开原始 CSV</a> · <a href="运行报告.json">打开完整 JSON</a></p><details><summary>展开查看运行数据</summary><pre>''' + e(__import__('json').dumps({k:v for k,v in result.items() if k not in ('exception_details','school_details')}, ensure_ascii=False, indent=2)) + '''</pre></details></section></main>
<script>const search=document.getElementById('search');const rows=Array.from(document.querySelectorAll('#issues tr'));function filter(){let n=0;for(const row of rows){let show=row.textContent.toLowerCase().includes(search.value.toLowerCase());row.hidden=!show;if(show)n++;}document.getElementById('count').textContent='当前显示 '+n+' 条';}search.addEventListener('input',filter);filter();</script></html>'''
    Path(result['report_html']).write_text(document, encoding='utf-8')
