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
            self.assertIn('CG-100', s['J2'].value)
            self.assertIn('CG-102', s['J2'].value)
            self.assertEqual([s['F5'].value, s['F6'].value], [1, 2])
            self.assertEqual(s['G5'].value, '面点豆制品')
            self.assertEqual(s['H5'].data_type, 's')
            self.assertEqual(s['M5'].value, .96)
            self.assertEqual(s['N5'].value, 2.88)
            self.assertTrue(s.column_dimensions['A'].hidden)
            self.assertEqual(s.column_dimensions['H'].width, 34)
            self.assertEqual(s.row_dimensions[5].height, 20)
            self.assertEqual(s['G5'].font.sz, 9)
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
