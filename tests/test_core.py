import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from bill_splitter.core import classify, export_bills


class BillTests(unittest.TestCase):
    def test_keywords_and_ambiguity(self):
        for word in ['教师', '成人', '职工', '老师']:
            self.assertEqual(classify('地址（' + word + '）'), '老师')
        for word in ['幼儿', '学生']:
            self.assertEqual(classify(word), '学生')
        self.assertEqual(classify('教师和学生'), '待核对')
        self.assertEqual(classify(None), '待核对')
        self.assertEqual(classify('成都市机关第三幼儿园总园（成人）'), '老师')
        self.assertEqual(classify('金牛区第三幼儿园（教师）'), '老师')
        self.assertEqual(classify('某幼儿园（幼儿）'), '学生')
        self.assertEqual(classify('某幼儿园'), '待核对')

    def test_export_reconciles_and_never_overwrites(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source = root / 'source.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.append(['客户名', '收货地址', '配送日期', '销售金额(元)'])
            ws.append(['学校A', '教师', '2026-10-01', 0.1])
            ws.append(['学校A', '幼儿', '2026-10-01', 0.2])
            ws.append(['学校A', '未知', '2026-10-02', 2])
            ws.append(['学校B', '成人学生', '2026-10-01', 3])
            ws.append([None, '职工', '2026-10-01', 4])
            ws.append(['合计', None, None, '=SUM(D2:D6)'])
            wb.save(source)
            result = export_bills(source, root / 'out')
            self.assertEqual(result['schools'], 2)
            self.assertEqual(result['exceptions'], 3)
            self.assertEqual(result['skipped_totals'], 1)
            self.assertEqual(result['source_amount'], '9.30')
            self.assertEqual(result['classified_amount'], '0.30')
            self.assertEqual(result['pending_amount'], '9.00')
            html = Path(result['report_html']).read_text(encoding='utf-8')
            self.assertIn('Excel 第 4 行', html)
            self.assertIn('未命中任何分类关键词', html)
            self.assertIn('同时命中老师和学生关键词', html)
            self.assertIn('9.30', html)
            self.assertEqual(result['exception_details'][0]['source_row'], 4)
            book = load_workbook(next(Path(result['directory']).glob('*学校A*.xlsx')), data_only=False)
            self.assertEqual(book.sheetnames, ['全部明细', '总计表', '老师明细', '学生明细'])
            self.assertEqual(book['全部明细'].max_row, 4)
            self.assertEqual(book['总计表']['D2'].value, '=SUM(B2:C2)')
            self.assertEqual(book['老师明细'].max_row, 2)
            self.assertEqual(book['学生明细'].max_row, 2)
            for sheet in book:
                for row in sheet:
                    for cell in row:
                        for edge in ('left', 'right', 'top', 'bottom'):
                            self.assertEqual(getattr(cell.border, edge).style, 'thin')
            book.close()
            second = export_bills(source, root / 'out')
            self.assertNotEqual(result['directory'], second['directory'])

    def test_invalid_amount_and_date_are_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            src = Path(td) / 'bad.xlsx'
            wb = Workbook()
            wb.active.append(['客户名', '收货地址', '配送日期', '销售金额(元)'])
            wb.active.append(['A', '老师', '坏日期', '坏金额'])
            wb.save(src)
            with self.assertRaises(ValueError):
                export_bills(src, Path(td) / 'out')


if __name__ == '__main__':
    unittest.main()
