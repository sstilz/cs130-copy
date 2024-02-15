import unittest, context
from sheets.lark_parser import LarkParser
from sheets.formula_evaluator import FormulaEvaluator
from sheets.workbook import Workbook
from decimal import Decimal
from sheets.cell import Cell
from sheets.cell_error_type import CellErrorType, CellError

class TestFormulaEvaluator(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        (self.index, self.name) = self.wb.new_sheet("Sheet1")

    def test_formulas_with_addition_involving_error_literals(self):
        # "=4 + #REF!" is a valid formula and should evaluate to a BAD_REFERENCE
        self.wb.set_cell_contents(self.name, "A1", "=4 + #ERROR!")
        self.wb.set_cell_contents(self.name, "A2", "=4 + #CIRCREF!")
        self.wb.set_cell_contents(self.name, "A3", "=4 + #REF!")
        self.wb.set_cell_contents(self.name, "A4", "=4 + #NAME?")
        self.wb.set_cell_contents(self.name, "A5", "=4 + #VALUE!")
        self.wb.set_cell_contents(self.name, "A6", "=4 + #DIV/0!")

        assert isinstance(self.wb.get_cell_value(self.name, "A1"), CellError)
        assert self.wb.get_cell_value(self.name, "A1").get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A2"), CellError)
        assert self.wb.get_cell_value(self.name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A3"), CellError)
        assert self.wb.get_cell_value(self.name, "A3").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.BAD_NAME
        assert isinstance(self.wb.get_cell_value(self.name, "A5"), CellError)
        assert self.wb.get_cell_value(self.name, "A5").get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A6"), CellError)
        assert self.wb.get_cell_value(self.name, "A6").get_type() == CellErrorType.DIVIDE_BY_ZERO
    
    def test_formulas_with_subtraction_involving_error_literals(self):
        # "=4 - #REF!" is a valid formula and should evaluate to a BAD_REFERENCE
        self.wb.set_cell_contents(self.name, "A1", "=4 - #ERROR!")
        self.wb.set_cell_contents(self.name, "A2", "=4 - #CIRCREF!")
        self.wb.set_cell_contents(self.name, "A3", "=4 - #REF!")
        self.wb.set_cell_contents(self.name, "A4", "=4 - #NAME?")
        self.wb.set_cell_contents(self.name, "A5", "=4 - #VALUE!")
        self.wb.set_cell_contents(self.name, "A6", "=4 - #DIV/0!")

        assert isinstance(self.wb.get_cell_value(self.name, "A1"), CellError)
        assert self.wb.get_cell_value(self.name, "A1").get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A2"), CellError)
        assert self.wb.get_cell_value(self.name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A3"), CellError)
        assert self.wb.get_cell_value(self.name, "A3").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.BAD_NAME
        assert isinstance(self.wb.get_cell_value(self.name, "A5"), CellError)
        assert self.wb.get_cell_value(self.name, "A5").get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A6"), CellError)
        assert self.wb.get_cell_value(self.name, "A6").get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_formulas_with_division_involving_error_literals(self):
        # "=4 / #REF!" is a valid formula and should evaluate to a BAD_REFERENCE
        self.wb.set_cell_contents(self.name, "A1", "=4 / #ERROR!")
        self.wb.set_cell_contents(self.name, "A2", "=4 / #CIRCREF!")
        self.wb.set_cell_contents(self.name, "A3", "=4 / #REF!")
        self.wb.set_cell_contents(self.name, "A4", "=4 / #NAME?")
        self.wb.set_cell_contents(self.name, "A5", "=4 / #VALUE!")
        self.wb.set_cell_contents(self.name, "A6", "=4 / #DIV/0!")

        assert isinstance(self.wb.get_cell_value(self.name, "A1"), CellError)
        assert self.wb.get_cell_value(self.name, "A1").get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A2"), CellError)
        assert self.wb.get_cell_value(self.name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A3"), CellError)
        assert self.wb.get_cell_value(self.name, "A3").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.BAD_NAME
        assert isinstance(self.wb.get_cell_value(self.name, "A5"), CellError)
        assert self.wb.get_cell_value(self.name, "A5").get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A6"), CellError)
        assert self.wb.get_cell_value(self.name, "A6").get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_formulas_with_multiplication_involving_error_literals(self):
        # "=4 * #REF!" is a valid formula and should evaluate to a BAD_REFERENCE
        self.wb.set_cell_contents(self.name, "A1", "=4 * #ERROR!")
        self.wb.set_cell_contents(self.name, "A2", "=4 * #CIRCREF!")
        self.wb.set_cell_contents(self.name, "A3", "=4 * #REF!")
        self.wb.set_cell_contents(self.name, "A4", "=4 * #NAME?")
        self.wb.set_cell_contents(self.name, "A5", "=4 * #VALUE!")
        self.wb.set_cell_contents(self.name, "A6", "=4 * #DIV/0!")

        assert isinstance(self.wb.get_cell_value(self.name, "A1"), CellError)
        assert self.wb.get_cell_value(self.name, "A1").get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A2"), CellError)
        assert self.wb.get_cell_value(self.name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A3"), CellError)
        assert self.wb.get_cell_value(self.name, "A3").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.BAD_NAME
        assert isinstance(self.wb.get_cell_value(self.name, "A5"), CellError)
        assert self.wb.get_cell_value(self.name, "A5").get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A6"), CellError)
        assert self.wb.get_cell_value(self.name, "A6").get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_formulas_with_parentheses_involving_error_literals(self):
        # "=(#REF!)" is a valid formula and should evaluate to a BAD_REFERENCE
        self.wb.set_cell_contents(self.name, "A1", "=(#ERROR!)")
        self.wb.set_cell_contents(self.name, "A2", "= (#CIRCREF!)")
        self.wb.set_cell_contents(self.name, "A3", "=(#REF!)")
        self.wb.set_cell_contents(self.name, "A4", "=(#NAME?)")
        self.wb.set_cell_contents(self.name, "A5", "=(#VALUE!)")
        self.wb.set_cell_contents(self.name, "A6", "=(#DIV/0!)")

        assert isinstance(self.wb.get_cell_value(self.name, "A1"), CellError)
        assert self.wb.get_cell_value(self.name, "A1").get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A2"), CellError)
        assert self.wb.get_cell_value(self.name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A3"), CellError)
        assert self.wb.get_cell_value(self.name, "A3").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.BAD_NAME
        assert isinstance(self.wb.get_cell_value(self.name, "A5"), CellError)
        assert self.wb.get_cell_value(self.name, "A5").get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A6"), CellError)
        assert self.wb.get_cell_value(self.name, "A6").get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_formulas_involving_error_literals(self):
        # "=#REF!" is a valid formula and should evaluate to a BAD_REFERENCE
        self.wb.set_cell_contents(self.name, "A1", "=#ERROR!")
        self.wb.set_cell_contents(self.name, "A2", "=#CIRCREF!")
        self.wb.set_cell_contents(self.name, "A3", "=#REF!")
        self.wb.set_cell_contents(self.name, "A4", "=#NAME?")
        self.wb.set_cell_contents(self.name, "A5", "=#VALUE!")
        self.wb.set_cell_contents(self.name, "A6", "=#DIV/0!")
        assert isinstance(self.wb.get_cell_value(self.name, "A1"), CellError)
        assert self.wb.get_cell_value(self.name, "A1").get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A2"), CellError)
        assert self.wb.get_cell_value(self.name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A3"), CellError)
        assert self.wb.get_cell_value(self.name, "A3").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.BAD_NAME
        assert isinstance(self.wb.get_cell_value(self.name, "A5"), CellError)
        assert self.wb.get_cell_value(self.name, "A5").get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A6"), CellError)
        assert self.wb.get_cell_value(self.name, "A6").get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_formulas_with_unary_involving_error_literals(self):
        # "=-#REF!" is a valid formula and should evaluate to a BAD_REFERENCE
        self.wb.set_cell_contents(self.name, "A1", "=-#ERROR!")
        self.wb.set_cell_contents(self.name, "A2", "=-#CIRCREF!")
        self.wb.set_cell_contents(self.name, "A3", "=-#REF!")
        self.wb.set_cell_contents(self.name, "A4", "=-#NAME?")
        self.wb.set_cell_contents(self.name, "A5", "=-#VALUE!")
        self.wb.set_cell_contents(self.name, "A6", "=-#DIV/0!")

        assert isinstance(self.wb.get_cell_value(self.name, "A1"), CellError)
        assert self.wb.get_cell_value(self.name, "A1").get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A2"), CellError)
        assert self.wb.get_cell_value(self.name, "A2").get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A3"), CellError)
        assert self.wb.get_cell_value(self.name, "A3").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.BAD_NAME
        assert isinstance(self.wb.get_cell_value(self.name, "A5"), CellError)
        assert self.wb.get_cell_value(self.name, "A5").get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(self.wb.get_cell_value(self.name, "A6"), CellError)
        assert self.wb.get_cell_value(self.name, "A6").get_type() == CellErrorType.DIVIDE_BY_ZERO


    def test_bad_reference_add_sheet(self):
        self.wb.set_cell_contents(self.name, "A1", "=Sheet2!A2")
        self.wb.set_cell_contents(self.name, "A2", "=Sheet2!A1 + 4")
        self.wb.set_cell_contents(self.name, "A3", "=Sheet2!A1 * 4")
        self.wb.set_cell_contents(self.name, "A4", "=4/Sheet2!A1")
        self.wb.set_cell_contents(self.name, "A5", "=(Sheet2!A1)")
        self.wb.set_cell_contents(self.name, "A6", "=Sheet2!A1 - 4")
        self.wb.set_cell_contents(self.name, "A7", "=+Sheet2!A1")
        self.wb.set_cell_contents(self.name, "A8", "=Sheet2!A1&\"hello\"")
        
        assert isinstance(self.wb.get_cell_value(self.name, "A1"), CellError)
        assert self.wb.get_cell_value(self.name, "A1").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A2"), CellError)
        assert self.wb.get_cell_value(self.name, "A2").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A3"), CellError)
        assert self.wb.get_cell_value(self.name, "A3").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A5"), CellError)
        assert self.wb.get_cell_value(self.name, "A5").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A6"), CellError)
        assert self.wb.get_cell_value(self.name, "A6").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A7"), CellError)
        assert self.wb.get_cell_value(self.name, "A7").get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(self.wb.get_cell_value(self.name, "A8"), CellError)
        assert self.wb.get_cell_value(self.name, "A8").get_type() == CellErrorType.BAD_REFERENCE

        (self.index2, self.name2) = self.wb.new_sheet("Sheet2")

        assert self.wb.get_cell_value(self.name, "A1") == Decimal('0')
        assert self.wb.get_cell_value(self.name, "A2") == Decimal('4')
        assert self.wb.get_cell_value(self.name, "A3") == Decimal('0')
        assert isinstance(self.wb.get_cell_value(self.name, "A4"), CellError)
        assert self.wb.get_cell_value(self.name, "A4").get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert self.wb.get_cell_value(self.name, "A5") == Decimal('0')
        assert self.wb.get_cell_value(self.name, "A6") == Decimal('-4')
        assert self.wb.get_cell_value(self.name, "A7") == Decimal('0')
        assert self.wb.get_cell_value(self.name, "A8") == "hello"

    def test_unset_cells_none_content_value(self):
        # A1 is an unset cell, it's content and value should be None
        assert self.wb.get_cell_contents(self.name, "A1") == None
        assert self.wb.get_cell_value(self.name, "A1") == None
        
        # A2 is an unset cell that should be implicity created with None content and value
        # A1 should have value 0, as A2 is unset and is interpreted as 0
        self.wb.set_cell_contents(self.name, "A1", "=A2")
        assert self.wb.get_cell_contents(self.name, "A2") == None
        assert self.wb.get_cell_value(self.name, "A2") == None
        assert self.wb.get_cell_value(self.name, "A1") == Decimal('0')

    def test_remove_trailing_zeros_when_concat_num_string(self):
        self.wb.set_cell_contents(self.name,"A1", "5.0000000")
        self.wb.set_cell_contents(self.name, "A2", "' hello")
        self.wb.set_cell_contents(self.name, "A3", "=A1&A2")
        self.wb.set_cell_contents(self.name, "A10", "=A2&A1")

        # Check that concatenating num+str OR str+num removes trailing zeros
        assert self.wb.get_cell_value(self.name, "A3") == "5 hello"
        assert self.wb.get_cell_value(self.name, "A10") == " hello5"

        # Check that trailing zeros are only removed for numbers and not strings
        self.wb.set_cell_contents(self.name, "A4", "'5.0000000")
        self.wb.set_cell_contents(self.name, "A5", "=A4&A2")
        assert self.wb.get_cell_value(self.name, "A5") == "5.0000000 hello"

        # Check that concatenating two strings that could implicity be converted to numbers does not remove trailing zeros
        self.wb.set_cell_contents(self.name, "A6", "'5.0000000")
        self.wb.set_cell_contents(self.name, "A7", "'0")
        self.wb.set_cell_contents(self.name, "A8", "=A6&A7")
        assert self.wb.get_cell_value(self.name, "A8") == "5.00000000"

        self.wb.set_cell_contents(self.name, "A9", "=(A8)")
        c = self.wb._get_cell(self.name, "A9")
        assert self.wb.get_cell_value(self.name, "A9") == Decimal('5')

    def tearDown(self):
        del self.wb


if __name__ == '__main__':
    unittest.main()