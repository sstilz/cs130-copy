import unittest, context
from sheets.cell_error_type import CellError, CellErrorType
from sheets.lark_parser import LarkParser
from sheets.formula_evaluator import FormulaEvaluator
from sheets.workbook import Workbook
from decimal import Decimal
from sheets.cell import Cell

class TestFormulaEvaluator(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        (self.index, self.name) = self.wb.new_sheet()
        self.e = FormulaEvaluator(self, self.name)
        self.p = LarkParser()

    def test_add_with_num(self):
        refs, tree, error = self.p.parse_formula("=1+ 2")
        v = self.e.evaluate(tree)
        self.assertEqual(v, Decimal('3'))

    def test_add_with_error(self):
        refs, tree, error = self.p.parse_formula("=#ERROR! + 2")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.PARSE_ERROR)


    def test_add_with_string(self):
        refs, tree, error = self.p.parse_formula("= \"hello\" + 2")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.TYPE_ERROR)

    def test_mul_with_num(self):
        refs, tree, error = self.p.parse_formula("= 3 * 2")
        v = self.e.evaluate(tree)
        self.assertEqual(v, Decimal('6'))

    def test_add_with_error(self):
        refs, tree, error = self.p.parse_formula("=\"#ERROR!\" * 2")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.PARSE_ERROR)

    def test_mul_with_string(self):
        refs, tree, error = self.p.parse_formula("= \"hello\"  * 2")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.TYPE_ERROR)

    def test_div_with_zero(self):
        refs, tree, error = self.p.parse_formula("= 2/0")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.DIVIDE_BY_ZERO)

    def test_div_zero_with_error(self):
        refs, tree, error = self.p.parse_formula("= \"#ERROR!\"/0")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.PARSE_ERROR)

    def test_concat_strings(self):
        refs, tree, error = self.p.parse_formula("=\"a\"&\"b\"")
        v = self.e.evaluate(tree)
        self.assertEqual(v, "ab")

    def test_concat_num(self):
        refs, tree, error = self.p.parse_formula("=2&\"b\"")
        v = self.e.evaluate(tree)
        self.assertEqual(v, "2b")

    def test_concat_errors(self):
        refs, tree, error = self.p.parse_formula("=\"#ERROR!\"&\"#DIV/0\"")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.PARSE_ERROR)

    def test_unary_num(self):
        refs, tree, error = self.p.parse_formula("=+100")
        v = self.e.evaluate(tree)
        self.assertEqual(v, Decimal('100'))

    def test_unary_string(self):
        refs, tree, error = self.p.parse_formula("=-\"hello\"")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.TYPE_ERROR)

    def test_unary_error(self):
        refs, tree, error = self.p.parse_formula("=-\"#ERROR!\"")
        v = self.e.evaluate(tree)
        self.assertTrue(isinstance(v, CellError))
        self.assertEqual(v.get_type(), CellErrorType.PARSE_ERROR)

    # implicit conversion
    def test_add_with_string_implicit(self):
        refs, tree, error = self.p.parse_formula("= \"5\" + 2")
        v = self.e.evaluate(tree)
        self.assertEqual(v, Decimal('7'))

    def test_mul_with_string_implicit(self):
        refs, tree, error = self.p.parse_formula("= \"3\" * 2")
        v = self.e.evaluate(tree)
        self.assertEqual(v, Decimal('6'))

    def test_unary_with_string_implicit(self):
        refs, tree, error = self.p.parse_formula("= +\"8\"")
        v = self.e.evaluate(tree)
        self.assertEqual(v, Decimal('8'))
    
    def test_refer_to_big_cell_almost_beyond_max_extent(self):
        # This shouldn't cause any errors.
        self.wb.set_cell_contents('Sheet1', 'A1', '=ZZZZ9999')
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1'), Decimal('0'))

    def test_reference_to_cell_outside_extent_raises_exception(self):
        # Should give bad reference error because beyond max extent.
        self.wb.set_cell_contents('Sheet1', 'A1', '=ZZZZ10000')
        self.assertTrue(self.wb.get_cell_value('Sheet1', 'A1'), '#REF!')

    def tearDown(self):
        del self.wb



    # empty cell tests are in workbook because they depend on get_cell_value
    
if __name__ == '__main__':
    unittest.main()