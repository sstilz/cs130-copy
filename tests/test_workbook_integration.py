import unittest
import context
from sheets.workbook import Workbook
from sheets.cell_error_type import CellError, CellErrorType

class TestWorkbookIntegration(unittest.TestCase):
    def setUp(self):
        self.workbook = Workbook()
        self.workbook.new_sheet("Sheet1")
        self.workbook.new_sheet("Sheet2")
    
    def test_negative_circular_reference(self):
        """
        If A is a circref error, and B=-A, then B is also a circref error.
        """
        self.workbook.set_cell_contents("Sheet1", "A1", "=Sheet1!A1")
        value_a1 = self.workbook.get_cell_value("Sheet1", "A1")
        self.assertIsInstance(value_a1, CellError)
        self.assertEqual(value_a1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        self.workbook.set_cell_contents("Sheet1", "B1", "=-Sheet1!A1")
        value_b1 = self.workbook.get_cell_value("Sheet1", "B1")
        self.assertIsInstance(value_b1, CellError)
        self.assertEqual(value_b1.get_type(), CellErrorType.CIRCULAR_REFERENCE)
    
    def test_multiple_errors_circular_reference(self):
        """
        When a cell can be multiple errors, but it is part of a cycle, it should be a circular reference error.
        """
        # Set A1 to a formula that would cause multiple errors
        self.workbook.set_cell_contents("Sheet1", "A1", "=Sheet1!B1/0")
        value_a1 = self.workbook.get_cell_value("Sheet1", "A1")
        self.assertIsInstance(value_a1, CellError)
        self.assertEqual(value_a1.get_type(), CellErrorType.DIVIDE_BY_ZERO)

        # Set B1 to A1, creating a cycle
        self.workbook.set_cell_contents("Sheet1", "B1", "=Sheet1!A1")

        # Check if A1 is a circular reference error
        value_a1 = self.workbook.get_cell_value("Sheet1", "A1")
        self.assertIsInstance(value_a1, CellError)
        self.assertEqual(value_a1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        # Check if B1 is also a circular reference error
        value_b1 = self.workbook.get_cell_value("Sheet1", "B1")
        self.assertIsInstance(value_b1, CellError)
        self.assertEqual(value_b1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

    def test_multiple_errors_circular_reference_complex(self):
        # Note, we don't include parse errors because if we can't parse the formula, we can't check for cycles
        self.workbook.set_cell_contents("Sheet1", "A1", "= Sheet1!C1/0") # Divide by zero error
        self.workbook.set_cell_contents("Sheet1", "B1", "= Sheet1!C1 + Sheet3!B1 ") # Bad Reference
        self.workbook.set_cell_contents("Sheet1", "C1", "=  \"hello\" +  Sheet1!D1") # Value Error

        # Set E1 to A1, creating a cycle
        self.workbook.set_cell_contents("Sheet1", "D1", "=Sheet1!A1")

        # Check if A1 is a circular reference error
        value_a1 = self.workbook.get_cell_value("Sheet1", "A1")
        self.assertIsInstance(value_a1, CellError)
        self.assertEqual(value_a1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        # Check if B1 is also a circular reference error
        value_b1 = self.workbook.get_cell_value("Sheet1", "B1")
        self.assertIsInstance(value_b1, CellError)
        self.assertEqual(value_b1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        # Check if C1 is also a circular reference error
        value_c1 = self.workbook.get_cell_value("Sheet1", "C1")
        self.assertIsInstance(value_c1, CellError)
        self.assertEqual(value_c1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        # Check if D1 is also a circular reference error
        value_d1 = self.workbook.get_cell_value("Sheet1", "D1")
        self.assertIsInstance(value_d1, CellError)
        self.assertEqual(value_d1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

    def test_cycle_detection_and_error_propagation(self):
        self.workbook.set_cell_contents("Sheet1", "B1", "=Sheet2!A1 * 2")
        self.workbook.set_cell_contents("Sheet2", "A1", "=Sheet1!A1 + 5")

        # Introduce a circular reference 
        # - creating cycle Sheet1!B1 <- Sheet2!A1 <- Sheet1!A1 <- Sheet1!B1
        self.workbook.set_cell_contents("Sheet1", "A1", "=Sheet1!B1")

        # Check for cycle detection
        value_a1 = self.workbook.get_cell_value("Sheet1", "A1")
        self.assertIsInstance(value_a1, CellError)
        self.assertEqual(value_a1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        # Check if the circular reference error propagates to other 2 cells
        value_b1 = self.workbook.get_cell_value("Sheet1", "B1")
        self.assertIsInstance(value_b1, CellError)
        self.assertEqual(value_b1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        value_s2_a1 = self.workbook.get_cell_value("Sheet2", "A1")
        self.assertIsInstance(value_s2_a1, CellError)
        self.assertEqual(value_s2_a1.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        # Introduce an error in Sheet2 (e.g., division by zero), breaking the cycle
        self.workbook.set_cell_contents("Sheet2", "A1", "=1/0")

        # Check if the error propagates
        value_b1_after_error = self.workbook.get_cell_value("Sheet1", "B1")
        self.assertIsInstance(value_b1_after_error, CellError)
        self.assertEqual(value_b1_after_error.get_type(), CellErrorType.DIVIDE_BY_ZERO)

        value_s1_a1_after_error = self.workbook.get_cell_value("Sheet1", "A1")
        self.assertIsInstance(value_s1_a1_after_error, CellError)
        self.assertEqual(value_s1_a1_after_error.get_type(), CellErrorType.DIVIDE_BY_ZERO)


    def test_deleting_sheet_and_impact(self):
        self.workbook.set_cell_contents("Sheet1", "A1", "10")
        self.workbook.set_cell_contents("Sheet2", "A1", "=Sheet1!A1 + 5")
        self.workbook.set_cell_contents("Sheet1", "B1", "=Sheet2!A1 * 2")
        
        # Delete Sheet2 and check the impact on Sheet1
        self.workbook.del_sheet("Sheet2")

        # Check if Sheet1's cell referencing Sheet2 now has a reference error
        value_b1_post_deletion = self.workbook.get_cell_value("Sheet1", "B1")
        self.assertIsInstance(value_b1_post_deletion, CellError)
        self.assertEqual(value_b1_post_deletion.get_type(), CellErrorType.BAD_REFERENCE)

    def tearDown(self):
        del self.workbook

if __name__ == '__main__':
    unittest.main()