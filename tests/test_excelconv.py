import unittest
import excelconv_core as exc

'''エクセル行列名変換モジュールの単体テストクラス'''
class TestExcelConv(unittest.TestCase):

# 各テストメソッドはtest_で始まること
# test_で始まるメソッドはテスト実行時に自動的にコールされる


    # 列名から行への変換
    def test_column_to_row(self):

        # 最小,+1
        self.assertEqual(exc.convert_excel_axis('A'), '1')
        self.assertEqual(exc.convert_excel_axis('B'), '2')

        # 最大,-1
        self.assertEqual(exc.convert_excel_axis('FXSHRXW'), '2147483647')
        self.assertEqual(exc.convert_excel_axis('FXSHRXV'), '2147483646')

        # 中間
        self.assertEqual(exc.convert_excel_axis('DAA'),'2731')

        # 異常値(例外の発生を確認)
        # 空
        self.assertRaises(exc.ExcelConvError, lambda: exc.convert_excel_axis(''))

        # 異常な文字
        self.assertRaises(exc.ExcelConvError, lambda: exc.convert_excel_axis('AA[/[-0'))

        # 範囲外
        self.assertRaises(exc.ExcelConvError, lambda: exc.convert_excel_axis('FXSHRXX'))
        self.assertRaises(exc.ExcelConvError, lambda: exc.convert_excel_axis('ABCSEFGHIJKLMN'))


    # 行から列名への変換
    def test_row_to_column(self):
        # 最小,+1
        self.assertEqual(exc.convert_excel_axis('1'), 'A')
        self.assertEqual(exc.convert_excel_axis('2'), 'B')

        # 最大,-1
        self.assertEqual(exc.convert_excel_axis('2147483647'), 'FXSHRXW')
        self.assertEqual(exc.convert_excel_axis('2147483646'), 'FXSHRXV')

        # 中間
        self.assertEqual(exc.convert_excel_axis('2731'), 'DAA')

        # 異常値(例外の発生を確認)
        # 0
        self.assertRaises(exc.ExcelConvError, lambda: exc.convert_excel_axis('0'))

        # 負数
        self.assertRaises(exc.ExcelConvError, lambda: exc.convert_excel_axis('-1'))

        # 範囲外
        self.assertRaises(exc.ExcelConvError, lambda: exc.convert_excel_axis('2147483648'))
        self.assertRaises(exc.ExcelConvError, lambda: exc.convert_excel_axis('12354634231'))


if __name__ == "__main__":
    unittest.main()
