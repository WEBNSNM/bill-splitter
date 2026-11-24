import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from bill_splitter.delivery import export_delivery


class DeliveryTests(unittest.TestCase):
    def source(self, root, objects=('教师', '学生', '教师')):
        p = root / 'source.xlsx'
        w = Workbook()
        s = w.active
        s.append(['配送日期', '订单编号', '客户名', '分类对象', '品类', '商品名', '计量单位', '销售数量', '销售单价(元)', '商品备注'])
        for i, obj in enumerate(objects):
            s.append(['2026-10-30', str(100+i), '测试学校', obj, '面点豆制品>豆腐', '=商品', '斤', 3, 1.005, '备注'])
            s.cell(s.max_row, 6).data_type = 's'
        w.save(p)
        w.close()
        return p

    def test_groups_prices_layout_and_source_preserved(self):
        with tempfile.TemporaryDirectory() as t:
            root = Path(t)
            p = self.source(root)
            original = p.read_bytes()
            result = export_delivery(p, root)
            self.assertEqual(result['schools'], 1)
            self.assertEqual(result['sheets'], 2)
            w = load_workbook(next(Path(result['directory']).glob('*.xlsx')))
            self.assertEqual(w.sheetnames, ['2026-10-30教师', '2026-10-30学生'])
            s = w.worksheets[0]
            self.assertEqual(s['J2'].value, 'CG-100')
            self.assertEqual(s['C6'].value, 'CG-102')
            self.assertEqual([s['F5'].value, s['F6'].value], [1, 2])
            self.assertEqual(s['G5'].value, '面点豆制品')
            self.assertEqual(s['H5'].data_type, 's')
            self.assertEqual(s['M5'].value, .96)
            self.assertEqual(s['N5'].value, 2.88)
            self.assertTrue(s.column_dimensions['A'].hidden)
            self.assertEqual(s.column_dimensions['G'].width, 9.625)
            self.assertEqual(s.column_dimensions['H'].width, 34.625)
            for col in ('M', 'N'):
                self.assertEqual(s.column_dimensions[col].width, 9.625)
            self.assertEqual(s['F1'].font.sz, 19.5)
            for cell in (*s[4], *s[5]):
                self.assertEqual(cell.alignment.horizontal, 'center')
            self.assertTrue(all(cell.font.bold for cell in s[4]))
            for address in ('F2', 'G2', 'I2', 'J2', 'L2', 'M2', 'F3', 'I3', 'L3'):
                self.assertEqual(s[address].alignment.horizontal, 'left')
            for address in ('I2', 'J2', 'L2', 'M2'):
                self.assertFalse(s[address].alignment.wrap_text)
            self.assertEqual(s['J2'].font.sz, 9)
            self.assertEqual(s['M2'].font.sz, 9)
            for margin in ('top', 'bottom'):
                self.assertAlmostEqual(getattr(s.page_margins, margin) * 2.54, 1)
            for margin in ('left', 'right'):
                self.assertAlmostEqual(getattr(s.page_margins, margin) * 2.54, 1.5)
            self.assertEqual(s.row_dimensions[5].height, 20)
            self.assertEqual(s['G5'].font.sz, 9)
            self.assertTrue(s['G5'].alignment.wrap_text)
            for row in s.iter_rows(min_row=1, max_row=3, min_col=6, max_col=15):
                for cell in row:
                    for side in (cell.border.left, cell.border.right, cell.border.top, cell.border.bottom):
                        self.assertIsNone(side and side.style)
            self.assertEqual(s['F4'].border.top.style, 'medium')
            self.assertEqual(s.cell(s.max_row, 6).value, '送货人：')
            self.assertEqual(s.cell(s.max_row, 11).value, '收货人：')
            self.assertGreaterEqual(s.row_dimensions[s.max_row].height, 75)
            for cell in s[s.max_row][5:15]:
                border = cell.border
                for side in (border.left, border.right, border.top, border.bottom):
                    self.assertIsNone(side and side.style)
            self.assertEqual(s.cell(s.max_row - 1, 6).border.bottom.style, 'medium')
            self.assertEqual(s.page_setup.paperSize, int(s.PAPERSIZE_LETTER))
            self.assertEqual(s.page_setup.orientation, 'portrait')
            self.assertTrue(s.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual((s.page_setup.fitToWidth, s.page_setup.fitToHeight), (1, 0))
            self.assertIsNone(s.print_title_rows)
            self.assertEqual(original, p.read_bytes())
            w.close()
            again = export_delivery(p, root)
            self.assertNotEqual(result['directory'], again['directory'])

    def test_unknown_or_conflicting_objects_rejected(self):
        for obj in ('', '教师学生', '未知'):
            with self.subTest(obj=obj), tempfile.TemporaryDirectory() as t:
                root = Path(t)
                with self.assertRaisesRegex(ValueError, '2'):
                    export_delivery(self.source(root, (obj,)), root)
                self.assertEqual(list(root.glob('配送单_*')), [])

    def test_nonfinite_and_missing_headers_rejected(self):
        with tempfile.TemporaryDirectory() as t:
            root = Path(t)
            p = self.source(root)
            w = load_workbook(p)
            w.active['H2'] = 'NaN'
            w.save(p)
            w.close()
            with self.assertRaisesRegex(ValueError, '2'):
                export_delivery(p, root)


if __name__ == '__main__':
    unittest.main()
