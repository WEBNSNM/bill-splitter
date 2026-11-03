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

    def test_activity_mode_is_per_school_and_preserves_source(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source = root / 'source.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.append(['序号', '客户名', '收货地址', '对象名称', '配送日期', '销售金额(元)'])
            for row in [
                [1, 'A', '学生', '老师', '2026-10-01', '1.005'],
                [2, 'A', '老师', '学生', '2026-10-01', 2],
                [3, 'A', '老师', '活动单', '2026-10-02', 3],
                [4, 'A', '老师', '老师学生', '2026-10-02', 4],
                [5, 'A', '老师', None, '2026-10-02', 5],
                [6, 'A', '老师', '活动单老师', '2026-10-02', 6],
                [7, 'B', '教师', '学生', '2026-10-01', 7],
            ]:
                ws.append(row)
            wb.save(source)
            wb.close()
            original = source.read_bytes()
            result = export_bills(source, root / 'out')
            self.assertEqual(source.read_bytes(), original)
            self.assertEqual(result['source_amount'], '28.01')
            self.assertEqual(result['classified_amount'], '13.01')
            self.assertEqual(result['pending_amount'], '15.00')
            self.assertEqual(result['exceptions'], 3)
            for school in result['school_details']:
                book = load_workbook(Path(result['directory']) / school['file'])
                self.assertEqual(len(book.sheetnames), 5 if school['school'] == 'A' else 4)
                for sheet in book:
                    self.assertNotIn('序号', [c.value for c in sheet[1]])
                    for row in range(1, sheet.max_row + 1):
                        self.assertEqual(sheet.row_dimensions[row].height, 17.25)
                self.assertEqual(book['全部明细']['E2'].value, 1.01 if school['school'] == 'A' else 7)
                self.assertEqual(book['老师明细'].max_row, 2 if school['school'] == 'A' else 1)
                if school['school'] == 'A':
                    self.assertEqual(book['老师明细']['C2'].value, '老师')
                    self.assertEqual(book['活动单明细']['E2'].value, 3)
                    summary = book['总计表']
                    self.assertEqual([c.value for c in summary[1]], ['配送日期', '老师', '学生', '活动单', '总计'])
                    self.assertEqual(summary['E2'].value, '=SUM(B2:D2)')
                    self.assertEqual(summary['D3'].value, 3)
                    self.assertEqual(summary['D4'].value, '=SUM(D2:D3)')
                    self.assertEqual(summary['E4'].value, '=SUM(E2:E3)')
                else:
                    self.assertEqual(book['学生明细'].max_row, 2)
                book.close()
            html = Path(result['report_html']).read_text(encoding='utf-8')
            self.assertIn('对象名称', html)
            self.assertEqual(result['object_column'], 'D')

    def test_object_column_always_controls_classification_without_activity(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source = root / 'source.xlsx'
            wb = Workbook()
            wb.active.append(['客户名', '收货地址', '对象名称', '配送日期', '销售金额(元)'])
            for address, name, amount in [('普通地址', '学生', 1), ('普通地址', '教师', 2),
                                          ('老师学生', ' 学生 ', 3), ('老师', None, 4),
                                          ('学生', '未知', 5), ('老师', '学生餐', 6)]:
                wb.active.append(['A', address, name, '2026-10-01', amount])
            wb.save(source)
            wb.close()
            result = export_bills(source, root / 'out')
            self.assertEqual(result['classified_amount'], '6.00')
            self.assertEqual(result['pending_amount'], '15.00')
            self.assertEqual([r['source_row'] for r in result['exception_details']], [5, 6, 7])
            self.assertTrue(all(r['classification_field'] == '对象名称' for r in result['exception_details']))
            book = load_workbook(next(Path(result['directory']).glob('*.xlsx')))
            self.assertEqual(len(book.sheetnames), 4)
            self.assertEqual(book['学生明细'].max_row, 3)
            self.assertEqual(book['老师明细']['E2'].value, 2)
            self.assertEqual(book['总计表']['B2'].value, 2)
            self.assertEqual(book['总计表']['C2'].value, 4)
            book.close()

    def test_serial_column_in_middle_keeps_date_amount_mapping(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src = root / 'source.xlsx'
            wb = Workbook()
            wb.active.append(['客户名', '序号', '配送日期', '收货地址', '下单金额(元)'])
            wb.active.append(['A', 99, '2026-10-01', '老师', '2.345'])
            wb.save(src)
            wb.close()
            result = export_bills(src, root / 'out', '下单金额(元)')
            book = load_workbook(next(Path(result['directory']).glob('*.xlsx')))
            self.assertEqual(book['全部明细'].max_column, 4)
            self.assertEqual(book['全部明细']['D2'].value, 2.35)
            self.assertEqual(book['全部明细']['B2'].number_format, 'yyyy-mm-dd')
            book.close()

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
