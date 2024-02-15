import context
from decimal import Decimal
from sheets import *
from sheets.cell import Cell
from sheets.value_type import ValueType
from sheets.cell_error_type import CellError, CellErrorType
import lark

import unittest

class TestCell(unittest.TestCase):

    def test_empty_cell(self):
        cell = Cell("")
        self.assertIsNone(cell.content)
        self.assertIsNone(cell.value)

        cell = Cell("'   ")
        self.assertIsNone(cell.content)
        self.assertIsNone(cell.value)

        cell = Cell("   '")
        self.assertIsNone(cell.content)
        self.assertIsNone(cell.value)

        cell = Cell("   ")
        self.assertIsNone(cell.content)
        self.assertIsNone(cell.value)

        cell = Cell("  ' ")
        self.assertIsNone(cell.content)
        self.assertIsNone(cell.value)

    def test_string_cell(self):
        cell = Cell("'Hello")
        self.assertEqual(cell.type, ValueType.STRING)
        self.assertEqual(cell.value, "Hello")

        cell = Cell(" '''")
        self.assertEqual(cell.type, ValueType.STRING)
        self.assertEqual(cell.value, "''")

        cell = Cell("'   hello")
        self.assertEqual(cell.type, ValueType.STRING)
        self.assertEqual(cell.value, "   hello")

        cell = Cell("'0")
        self.assertEqual(cell.type, ValueType.STRING)
        self.assertEqual(cell.value, "0")

    def test_formula_cell(self):
        cell = Cell("=A1+B1")
        self.assertEqual(cell.type, ValueType.FORMULA)

        cell = Cell("=A1+B1-C1*2/3")
        self.assertEqual(cell.type, ValueType.FORMULA)

    def test_decimal_cell(self):
        cell = Cell("123.4500")
        self.assertEqual(cell.type, ValueType.NUMBER)
        self.assertEqual(cell.value, Decimal('123.45'))

    def test_nan_infinity(self):
        cell = Cell("NaN")
        self.assertEqual(cell.type, ValueType.STRING)
        self.assertEqual(cell.value, "NaN")

        cell = Cell("Infinity")
        self.assertEqual(cell.type, ValueType.STRING)
        self.assertEqual(cell.value, "Infinity")

        cell = Cell("-Infinity")
        self.assertEqual(cell.type, ValueType.STRING)
        self.assertEqual(cell.value, "-Infinity")

    def test_strip_trailing_zeros(self):
        cell = Cell('0.1000')
        self.assertEqual(cell.value, Decimal('0.1'))

        cell = Cell('100.')
        self.assertEqual(cell.value, Decimal('100'))

        cell = Cell('0.000')
        self.assertEqual(cell.value, Decimal('0'))

    def test_check_type(self):
        cell_formula = Cell('=A1+A10')
        self.assertEqual(cell_formula.type, ValueType.FORMULA)

        cell_string = Cell("'This is a string'")
        self.assertEqual(cell_string.type, ValueType.STRING)

        cell_literal = Cell('A literal value that is not a decimal is a string')
        self.assertEqual(cell_literal.type, ValueType.STRING)

        cell_literal = Cell('0')
        self.assertEqual(cell_literal.type, ValueType.NUMBER)

    def test_check_error_type(self):
        cell_error_parse = Cell("#ERROR!")
        self.assertEqual(cell_error_parse.type, ValueType.ERROR)
        self.assertIsInstance(cell_error_parse.get_value(), CellError)
        self.assertEqual(cell_error_parse.get_value().get_type(), CellErrorType.PARSE_ERROR)

        cell_error_circ = Cell("#CIRCREF!")
        self.assertEqual(cell_error_circ.type, ValueType.ERROR)
        self.assertIsInstance(cell_error_circ.value, CellError)
        self.assertEqual(cell_error_circ.value.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        cell_error_ref = Cell("#REF!")
        self.assertEqual(cell_error_ref.type, ValueType.ERROR)
        self.assertIsInstance(cell_error_ref.value, CellError)
        self.assertEqual(cell_error_ref.value.get_type(), CellErrorType.BAD_REFERENCE)

        cell_error_name = Cell("#NAME?")
        self.assertEqual(cell_error_name.type, ValueType.ERROR)
        self.assertIsInstance(cell_error_name.value, CellError)
        self.assertEqual(cell_error_name.value.get_type(), CellErrorType.BAD_NAME)

        cell_error_type = Cell("#VALUE!")
        self.assertEqual(cell_error_type.type, ValueType.ERROR)
        self.assertIsInstance(cell_error_type.value, CellError)
        self.assertEqual(cell_error_type.value.get_type(), CellErrorType.TYPE_ERROR)

        cell_error_divzero = Cell("#DIV/0!")
        self.assertEqual(cell_error_divzero.type, ValueType.ERROR)
        self.assertIsInstance(cell_error_divzero.value, CellError)
        self.assertEqual(cell_error_divzero.value.get_type(), CellErrorType.DIVIDE_BY_ZERO)

    def test_is_formula(self):
        cell = Cell('=A1+B1')
        self.assertTrue(cell.is_formula())

        cell_not_formula = Cell('A1+B1')
        self.assertFalse(cell_not_formula.is_formula())

    def test_is_string(self):
        cell = Cell("'Hello'")
        self.assertTrue(cell.is_string())

        cell_literal_string = Cell('Hello')
        self.assertTrue(cell_literal_string.is_string())

    def test_literal(self):
        cell_int = Cell('100')
        self.assertEqual(cell_int.value, Decimal(100))

        cell_float = Cell('100.50')
        self.assertEqual(cell_float.value, Decimal(100.5))

        cell_string = Cell('Test')
        self.assertEqual(cell_string.value, 'Test')

    def test_evaluate_content(self):
        cell_string = Cell("'Hello'")
        self.assertEqual(cell_string.value, "Hello'")

        cell_decimal = Cell('100.50')
        self.assertEqual(cell_decimal.value, 100.5)

        cell_literal = Cell('Just a literal value')
        self.assertEqual(cell_literal.value, 'Just a literal value')

    def test_update_cell(self):
        cell = Cell("123")
        self.assertEqual(cell.type, ValueType.NUMBER)
        self.assertEqual(cell.value, Decimal("123"))

        cell.update("'Hello")
        self.assertEqual(cell.type, ValueType.STRING)
        self.assertEqual(cell.value, "Hello")

        cell.update("=A1+B1")
        self.assertEqual(cell.type, ValueType.FORMULA)
        # Assuming current implementation sets the tree for formulas:
        self.assertIsNotNone(cell.tree)

        cell.update("#ERROR!")
        self.assertEqual(cell.type, ValueType.ERROR)
        self.assertIsInstance(cell.value, CellError)
        self.assertEqual(cell.value.get_type(), CellErrorType.PARSE_ERROR)

    def test_string_cell_with_special_characters(self):
        special_chars = ["!@#$%^&*()_+{}:\<>?[];',./", "`~", "€£¥₹", "🥐"]
        for chars in special_chars:
            with self.subTest(chars=chars):
                cell = Cell(f"'{chars}")
                self.assertEqual(cell.type, ValueType.STRING)
                self.assertEqual(cell.value, chars)

    def test_cell_reference_parsing(self):
        cell = Cell("=A1+B2")
        expected_refs = [('A1', ), ('B2', )]
        self.assertListEqual(sorted(cell.get_refs()), sorted(expected_refs))

    def test_get_string_from_error_type(self):
        self.assertEqual(CellError.get_string_from_error_type(CellErrorType.CIRCULAR_REFERENCE), "#CIRCREF!")

    def test_is_error_string(self):
        self.assertEqual(CellError.is_error_string("#ERROR!"), True)
        self.assertEqual(CellError.is_error_string("#E!"), False)

    def test_get_error_type_from_string(self):
        self.assertEqual(CellError.get_error_type_from_string("#ERROR!"), CellErrorType.PARSE_ERROR)

if __name__ == '__main__':
    unittest.main()
