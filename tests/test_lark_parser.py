import unittest, context
from decimal import Decimal
from sheets.cell import Cell
from sheets.lark_parser import LarkParser

class TestLarkParser(unittest.TestCase):

    def setUp(self):
        self.parser = LarkParser()

    def test_sheet_name_with_spaces(self):
        formula = "=Sheet 1!A1"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertEqual(error, "#ERROR!")

        formula = "='Sheet 1'!A1"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertIsNone(error)

    def test_parse_formula_basic(self):
        formula = "=A1+B2"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertIsNone(error)
        self.assertEqual(refs, [('A1', ), ('B2', )])

    def test_parse_formula_cross_sheet(self):
        formula = "=Sheet1!Z20+Sheet2!AA299 - sHeEt3!dd88"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertIsNone(error)
        self.assertEqual(refs, [('SHEET1', 'Z20'), ('SHEET2', 'AA299'), ('SHEET3','DD88')])

    def test_parse_formula_complex(self):
        formula = "=A1*A1+B2+C2+(H1 - B6) - Sheet1!A1"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertIsNone(error)
        self.assertEqual(refs, [('A1', ), ('A1', ), ('B2', ), ('H1', ), ('B6', ), ('C2', ), ('SHEET1', 'A1')])

    def test_parse_formula_single_cell(self):
        formula = "=C3"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertIsNone(error)
        self.assertEqual(refs, [('C3', )])

    def test_parse_formula_invalid_syntax(self):
        formula = "=A1+"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertIsNotNone(error)
        self.assertIsNone(refs)
        self.assertIsNone(tree)

    def test_parse_formula_invalid_syntax(self):
        formula = "=++AZZ8"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertIsNotNone(error)
        self.assertIsNone(refs)
        self.assertIsNone(tree)

    def test_parse_formula_error_handling(self):
        formula = "=INVALID_FORMULA"
        refs, tree, error = self.parser.parse_formula(formula)
        self.assertIsNotNone(error)
        self.assertIsNone(refs)
        self.assertIsNone(tree)

if __name__ == '__main__':
    unittest.main()
