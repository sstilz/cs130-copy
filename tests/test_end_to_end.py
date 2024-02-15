import context
import unittest
from decimal import Decimal
from sheets.cell_error_type import CellErrorType, CellError
from sheets.workbook import Workbook
from sheets.formula_evaluator import FormulaEvaluator

class TestWorkbookExtended(unittest.TestCase):
    def test_workbook_implicit_conversion(self):
        wb = Workbook()
        (index, name) = wb.new_sheet("Sheet1")

        wb.set_cell_contents(name, 'a1', "'123")
        wb.set_cell_contents(name, 'b1', "5.3")
        wb.set_cell_contents(name, 'c1', "=a1*b1")

        value = wb.get_cell_value(name, 'c1')
        self.assertEqual(value, Decimal('651.9'))

        wb.set_cell_contents(name, 'a1', "'     123")
        value = wb.get_cell_value(name, 'c1')
        self.assertEqual(value, Decimal('651.9'))

        wb.set_cell_contents(name, 'd1', "")
        wb.set_cell_contents(name, 'e1', "=d1+5")
        value = wb.get_cell_value(name, 'e1')
        self.assertEqual(value, Decimal('5'))

        wb.set_cell_contents(name, 'f1', "=d1&a1")
        value = wb.get_cell_value(name, 'f1')
        self.assertEqual(value, "     123")

        wb.set_cell_contents(name, 'a1', None)
        value = wb.get_cell_value(name, 'a1')
        self.assertEqual(value, None)

        wb.set_cell_contents(name, 'b1', "=a1")
        value = wb.get_cell_value(name, 'b1')
        self.assertEqual(value, Decimal('0'))    

    def test_workbook_circular_reference(self):
        wb = Workbook()
        (index, name) = wb.new_sheet("Sheet2")

        wb.set_cell_contents(name, 'a1', "=b1")
        wb.set_cell_contents(name, 'b1', "=c1")
        wb.set_cell_contents(name, 'c1', "=b1/0")

        value = wb.get_cell_value(name, 'a1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        value = wb.get_cell_value(name, 'b1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        value = wb.get_cell_value(name, 'c1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

    def test_order_of_evaluation(self):
        wb = Workbook()
        (index, name) = wb.new_sheet("Sheet3")

        wb.set_cell_contents(name, 'a1', "=b1+c1")
        wb.set_cell_contents(name, 'b1', "=c1")
        wb.set_cell_contents(name, 'c1', "1")

        assert wb.get_cell_value(name, 'a1') == Decimal('2')
        assert wb.get_cell_value(name, 'b1') == Decimal('1')

        wb.set_cell_contents(name, 'c1', "2")
        assert wb.get_cell_value(name, 'a1') == Decimal('4')
        assert wb.get_cell_value(name, 'b1') == Decimal('2')

        # regardless of the order these cell's contents are set, they should 
        # evaluate in the correct order such that A1 has a value of 2
        # and B1 has a value of 1
        
        wb.set_cell_contents(name, 'b2', "=c2")
        wb.set_cell_contents(name, 'a2', "=b2+c2")
        wb.set_cell_contents(name, 'c2', "1")

        assert wb.get_cell_value(name, 'a2') == Decimal('2')
        assert wb.get_cell_value(name, 'b2') == Decimal('1')

        wb.set_cell_contents(name, 'c1', "2")
        assert wb.get_cell_value(name, 'a1') == Decimal('4')
        assert wb.get_cell_value(name, 'b1') == Decimal('2')

    def test_workbook_propagate_errors(self):
        wb = Workbook()
        (index, name) = wb.new_sheet("Sheet4")

        wb.set_cell_contents(name, 'a1', "=b1+5")
        wb.set_cell_contents(name, 'b1', "=15/0")

        value = wb.get_cell_value(name, 'a1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.DIVIDE_BY_ZERO

        value = wb.get_cell_value(name, 'b1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_circular_sheets(self):
        wb = Workbook()
        (index1, name1) = wb.new_sheet("Sheet1")
        (index2, name2) = wb.new_sheet("Sheet2")
        (index3, name3) = wb.new_sheet("Sheet3")

        # Add numbers, strings, and formulas
        wb.set_cell_contents(name1, 'a1', "=2")
        wb.set_cell_contents(name1, 'b1', "=Sheet1!a1")
        wb.set_cell_contents(name2, 'a1', '=Sheet1!b1')
        wb.set_cell_contents(name1, 'a1', '=Sheet2!a1')
        
        assert isinstance(wb.get_cell_value(name1, 'a1'), CellError)
        assert wb.get_cell_value(name1, 'a1').get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(wb.get_cell_value(name1, 'b1'), CellError)
        assert wb.get_cell_value(name1, 'b1').get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(wb.get_cell_value(name2, 'a1'), CellError)
        assert wb.get_cell_value(name2, 'a1').get_type() == CellErrorType.CIRCULAR_REFERENCE


    def test_circular_sheet_break(self):
        wb = Workbook()
        (index1, name1) = wb.new_sheet("Sheet1")
        (index2, name2) = wb.new_sheet("Sheet2")

        # Add numbers, strings, and formulas
        wb.set_cell_contents(name1, 'a1', "=2")
        wb.set_cell_contents(name1, 'b1', "=Sheet1!a1")
        wb.set_cell_contents(name2, 'a1', '=Sheet1!b1')
        wb.set_cell_contents(name1, 'a1', '=Sheet2!a1')
        
        assert isinstance(wb.get_cell_value(name1, 'a1'), CellError)
        assert wb.get_cell_value(name1, 'a1').get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(wb.get_cell_value(name1, 'b1'), CellError)
        assert wb.get_cell_value(name1, 'b1').get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(wb.get_cell_value(name2, 'a1'), CellError)
        assert wb.get_cell_value(name2, 'a1').get_type() == CellErrorType.CIRCULAR_REFERENCE


        wb.set_cell_contents(name1, 'a1', "=2")
        assert wb.get_cell_value(name1, 'a1') == Decimal('2')
        assert wb.get_cell_value(name1, 'b1') == Decimal('2')
        assert wb.get_cell_value(name2, 'a1') == Decimal('2')


    def test_circular_sheet_break_with_literal(self):
        wb = Workbook()
        (index1, name1) = wb.new_sheet("Sheet1")
        (index2, name2) = wb.new_sheet("Sheet2")

        # Add numbers, strings, and formulas
        wb.set_cell_contents(name1, 'a1', "=2")
        wb.set_cell_contents(name1, 'b1', "=Sheet1!a1")
        wb.set_cell_contents(name2, 'a1', '=Sheet1!b1')
        wb.set_cell_contents(name1, 'a1', '=Sheet2!a1')
        
        assert isinstance(wb.get_cell_value(name1, 'a1'), CellError)
        assert wb.get_cell_value(name1, 'a1').get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(wb.get_cell_value(name1, 'b1'), CellError)
        assert wb.get_cell_value(name1, 'b1').get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(wb.get_cell_value(name2, 'a1'), CellError)
        assert wb.get_cell_value(name2, 'a1').get_type() == CellErrorType.CIRCULAR_REFERENCE

        wb.set_cell_contents(name1, 'a1', " ' hello")
        assert wb.get_cell_value(name1, 'a1') == " hello"
        assert wb.get_cell_value(name1, 'b1') == " hello"
        assert wb.get_cell_value(name2, 'a1') == " hello"

    def test_propagate_parse_error(self):
        wb = Workbook()
        (index1, name1) = wb.new_sheet("Sheet1")
        (index2, name2) = wb.new_sheet("Sheet2")

        # Add numbers, strings, and formulas
        wb.set_cell_contents(name1, 'a1', "=2")
        wb.set_cell_contents(name1, 'b1', "=Sheet1!a1")
        wb.set_cell_contents(name2, 'a1', '=Sheet1!b1')
        wb.set_cell_contents(name1, 'a1', '==+')
        
        assert isinstance(wb.get_cell_value(name1, 'a1'), CellError)
        assert wb.get_cell_value(name1, 'a1').get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(wb.get_cell_value(name1, 'b1'), CellError)
        assert wb.get_cell_value(name1, 'b1').get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(wb.get_cell_value(name2, 'a1'), CellError)
        assert wb.get_cell_value(name2, 'a1').get_type() == CellErrorType.PARSE_ERROR

        wb.set_cell_contents(name1, 'a1', "2")
        assert wb.get_cell_value(name1, 'a1') == Decimal('2')
        assert wb.get_cell_value(name1, 'b1') == Decimal('2')
        assert wb.get_cell_value(name2, 'a1') == Decimal('2')

    def test_concat_with_nums(self):
        wb = Workbook()
        (index, name) = wb.new_sheet("Sheet1")
        wb.set_cell_contents(name, "a1", "=1")
        wb.set_cell_contents(name, "a2", "two")
        wb.set_cell_contents(name, "a3", "=A1&A2")
        assert wb.get_cell_value(name, "a3") == "1two"

    def test_all(self):
        wb = Workbook()
        (index1, name1) = wb.new_sheet("Sheet1")
        (index2, name2) = wb.new_sheet("Sheet2")
        (index3, name3) = wb.new_sheet("Sheet3")

        # Add numbers, strings, and formulas

        wb.set_cell_contents(name1, 'a1', "'123")
        wb.set_cell_contents(name1, 'b1', "5.30000")
        wb.set_cell_contents(name1, 'e1', "=a1+b1")
        assert wb.get_cell_value(name1, 'e1') == Decimal('128.3')

        wb.set_cell_contents(name2, 'a1', "'   123")
        wb.set_cell_contents(name2, 'b1', None)
        wb.set_cell_contents(name2, 'c1', "")

        wb.set_cell_contents(name3, "a1", "=Sheet1!a1+Sheet1!b1+Sheet1!e1")
        assert wb.get_cell_value(name3, "a1") == Decimal('256.6')
        wb.set_cell_contents(name3, "a2", "=-Sheet3!a1 + (Sheet2!a1 + Sheet2!b1)*Sheet2!c1")
        assert wb.get_cell_value(name3, "a2") == Decimal('-256.6')

        # Test circular reference
        wb.set_cell_contents(name1, 'e1', "=Sheet3!a2")

        assert isinstance(wb.get_cell_value(name1, 'e1'), CellError)
        assert wb.get_cell_value(name1, 'e1').get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(wb.get_cell_value(name3, 'a1'), CellError)
        assert wb.get_cell_value(name3, 'a1').get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(wb.get_cell_value(name3, 'a2'), CellError)
        assert wb.get_cell_value(name3, 'a2').get_type() == CellErrorType.CIRCULAR_REFERENCE

        # Test breaking circular reference
        wb.set_cell_contents(name1, 'e1', " ' hello")
        assert wb.get_cell_value(name1, 'e1') == " hello"
        
        assert isinstance(wb.get_cell_value(name3, 'a1'), CellError)
        assert isinstance(wb.get_cell_value(name3, 'a2'), CellError)
        assert wb.get_cell_value(name3, 'a1').get_type() == CellErrorType.TYPE_ERROR
        assert wb.get_cell_value(name3, 'a2').get_type() == CellErrorType.TYPE_ERROR
    
if __name__ == '__main__':
    unittest.main()
