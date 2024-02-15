import context
import unittest
from decimal import Decimal
from sheets.cell_error_type import CellErrorType, CellError
from sheets.workbook import Workbook
from sheets.formula_evaluator import FormulaEvaluator

class TestWorkbookReorderCopy(unittest.TestCase):
    def setUp(self):
        self.workbook = Workbook()

    # Test reorder sheets
    def test_reorder_sheets(self):
        self.workbook.new_sheet("Sheet1")
        self.workbook.new_sheet("Sheet2")
        self.workbook.move_sheet("Sheet1", 1)

        sheets = self.workbook.list_sheets()
        self.assertEqual(sheets, ["Sheet2", "Sheet1"])

    def test_move_sheet_to_front_and_back(self):
        for i in range(5):
            self.workbook.new_sheet(f"Sheet{i+1}")
        self.workbook.move_sheet("Sheet3", 0)
        sheets = self.workbook.list_sheets()
        self.assertEqual(sheets, ["Sheet3", "Sheet1", "Sheet2", "Sheet4", "Sheet5"])
        self.workbook.move_sheet("Sheet3", 4)
        sheets = self.workbook.list_sheets()
        self.assertEqual(sheets, ["Sheet1", "Sheet2", "Sheet4", "Sheet5", "Sheet3"])

    def test_move_sheet_to_same_position(self):
        for i in range(3):
            self.workbook.new_sheet(f"Sheet{i+1}")

        for i in range(3):
            self.workbook.move_sheet(f"Sheet{i+1}", i)
            sheets = self.workbook.list_sheets()
            self.assertEqual(sheets[i], f"Sheet{i+1}")

    def test_move_sheet_out_of_bounds(self):
        self.workbook.new_sheet("Sheet1")
        with self.assertRaises(IndexError):
            self.workbook.move_sheet("Sheet1", -1)
        with self.assertRaises(IndexError):
            self.workbook.move_sheet("Sheet1", 1)

    def test_move_nonexistent_sheet(self):
        # No sheet have been created yet
        with self.assertRaises(KeyError):
            self.workbook.move_sheet("NonExistentSheet", 0)

    def test_move_sheet_with_dependencies(self):
        self.workbook.new_sheet("DataSheet")
        self.workbook.new_sheet("ReferenceSheet")
        self.workbook.set_cell_contents("DataSheet", "A1", "5")
        self.workbook.set_cell_contents("ReferenceSheet", "A1", "=DataSheet!A1 * 2")
        self.assertEqual(self.workbook.get_cell_value("ReferenceSheet", "A1"), Decimal("10"))

        # Move DataSheet to a new position
        self.workbook.move_sheet("DataSheet", 1)

        self.assertEqual(self.workbook.get_cell_value("DataSheet", "A1"), Decimal("5"))
        self.assertEqual(self.workbook.get_cell_contents("ReferenceSheet", "A1"), "=DataSheet!A1 * 2")

        # Verify that the formula in ReferenceSheet still works correctly
        self.assertEqual(self.workbook.get_cell_value("DataSheet", "A1"), Decimal("5"))
        self.assertEqual(self.workbook.get_cell_value("ReferenceSheet", "A1"), Decimal("10"))

    def test_move_multiple_sheets(self):
        for i in range(4):
            self.workbook.new_sheet(f"Sheet{i+1}")
        self.workbook.move_sheet("Sheet3", 1)
        self.workbook.move_sheet("Sheet4", 0)
        expected_order = ["Sheet4", "Sheet1", "Sheet3", "Sheet2"]
        self.assertEqual(self.workbook.list_sheets(), expected_order)

    def test_move_sheet_case_insensitive_name(self):
        self.workbook.new_sheet("TestSheet")
        self.workbook.new_sheet("testsheet2")
        self.workbook.move_sheet("testsheet", 1)
        sheets = self.workbook.list_sheets()
        self.assertEqual(sheets, ["testsheet2", "TestSheet"])
        self.workbook.move_sheet("TESTSHEET2", 1)
        sheets = self.workbook.list_sheets()
        self.assertEqual(sheets, ["TestSheet", "testsheet2"])

    def test_copy_sheets(self):
        # assert cell contents are the same
        # assert copied sheet independent of original
        # assert changes made to sheet don't affect original
        # check correct naming of copied sheet [original sheet_ name]_1
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", "Test")
        _, name = self.workbook.copy_sheet("Sheet1")

        self.assertEqual(self.workbook.get_cell_contents(name, "A1"), "Test")
        self.workbook.set_cell_contents("Sheet1", "A1", "Changed")
        self.assertNotEqual(self.workbook.get_cell_contents(name, "A1"), "Changed")
        self.assertEqual(name, "Sheet1_1")
        self.workbook.set_cell_contents(name, "A1", "Change in Copy")
        self.assertNotEqual(self.workbook.get_cell_contents("Sheet1", "A1"), "Change in Copy")

    def test_copy_is_last_sheet(self):
        # assert copied sheet is last sheet
        self.workbook.new_sheet("Sheet1")
        self.workbook.new_sheet("Sheet2")
        self.workbook.new_sheet("Sheet3")
        _, name = self.workbook.copy_sheet("Sheet1")
        sheets = self.workbook.list_sheets()
        self.assertEqual(sheets[-1], "Sheet1_1")

    def test_copy_changes_other_sheets(self):
        # After changing the contents in original sheet S1, this properly
        # updates sheets that depend on S1, even the sheets that depend on S1
        # because they were copied from original sheets that depended on S1.
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", "5")
        self.workbook.new_sheet("Sheet2")
        self.workbook.set_cell_contents("Sheet2", "A1", "=Sheet1!A1 * 2")

        self.workbook.copy_sheet("Sheet2")
        self.workbook.set_cell_contents("Sheet1", "A1", "10")

        self.assertEqual(self.workbook.get_cell_value("Sheet2", "A1"), 20)
        self.assertEqual(self.workbook.get_cell_value("Sheet2_1", "A1"), 20)

    def test_sheet_independence_changing_original(self):
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", "5")
        self.workbook.new_sheet("Sheet2")
        self.workbook.set_cell_contents("Sheet2", "A1", "=Sheet1!A1 * 2")

        self.workbook.copy_sheet("Sheet2")
        self.assertEqual(self.workbook.get_cell_value("Sheet2_1", "A1"), 10)
        
        # Changing the OG sheet's contents after the copy should not affect 
        # the copied sheet.
        self.workbook.set_cell_contents("Sheet2", "A1", "=Sheet1!A1")
        self.assertEqual(self.workbook.get_cell_value("Sheet2", "A1"), 5)
        self.assertEqual(self.workbook.get_cell_value("Sheet2_1", "A1"), 10)

    def test_sheet_independence_changing_copy(self):
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", "5")
        self.workbook.new_sheet("Sheet2")
        self.workbook.set_cell_contents("Sheet2", "A1", "=Sheet1!A1 * 2")

        self.workbook.copy_sheet("Sheet2")
        self.assertEqual(self.workbook.get_cell_value("Sheet2_1", "A1"), 10)
        self.workbook.set_cell_contents("Sheet2_1", "A1", "5")
        self.assertEqual(self.workbook.get_cell_value("Sheet2", "A1"), 10)
        self.assertEqual(self.workbook.get_cell_value("Sheet2_1", "A1"), 5)

    def test_copy_delete_naming(self):
        # If Sheet1 is copied -> should create Sheet1_1
        # Then Sheet1_1 is deleted -> then if Sheet1 is copied again
        # then Sheet1_1 should be created again
        self.workbook.new_sheet("Sheet1")
        _, name = self.workbook.copy_sheet("Sheet1")
        self.assertEqual(name, "Sheet1_1")

        self.workbook.del_sheet(name)
        _, new_name = self.workbook.copy_sheet("Sheet1")
        self.assertEqual(new_name, "Sheet1_1")

    def test_copy_twice_naming(self):
        # If Sheet1 is copied -> creates Sheet1_1
        # If Sheet1_1 is copied -> creates Sheet1_1_1
        self.workbook.new_sheet("Sheet1")
        _, name1 = self.workbook.copy_sheet("Sheet1")
        self.assertEqual(name1, "Sheet1_1")

        _, name2 = self.workbook.copy_sheet(name1)
        self.assertEqual(name2, "Sheet1_1_1")

    def test_name_iteration(self):
        # If Sheet1 is copied -> creates Sheet1_1
        # If Sheet1 is copied -> creates Sheet1_2
        # If Sheet1 is copied -> creates Sheet1_3
        # ... x 10
        self.workbook.new_sheet("Sheet1")
        for i in range(10):
            _, name = self.workbook.copy_sheet("Sheet1")
            self.assertEqual(name, f"Sheet1_{i+1}")
        
        #If we delete Sheet1_2 and Sheet1_4 and copy Sheet1 again twice, Sheet1_2 and Sheet1_4 should be recreated
        self.workbook.del_sheet("Sheet1_2")
        self.assertFalse("Sheet1_2" in self.workbook.list_sheets())
        self.workbook.del_sheet("Sheet1_4")
        self.assertFalse("Sheet1_4" in self.workbook.list_sheets())

        self.workbook.copy_sheet("Sheet1")
        self.assertTrue("Sheet1_2" in self.workbook.list_sheets())
        self.workbook.copy_sheet("Sheet1")
        self.assertTrue("Sheet1_4" in self.workbook.list_sheets())

    def test_copy_sheet_with_formulas(self):
        self.workbook.new_sheet("Original")
        self.workbook.set_cell_contents("Original", "A1", "10")
        self.workbook.set_cell_contents("Original", "B1", "=A1 * 2")
        _, copied_sheet_name = self.workbook.copy_sheet("Original")
    
        # Check if the formula in the copied sheet references its own cells
        self.assertEqual(self.workbook.get_cell_contents(copied_sheet_name, "B1"), "=A1 * 2")
        self.assertEqual(self.workbook.get_cell_value(copied_sheet_name, "B1"), Decimal("20"))

        self.workbook.set_cell_contents(copied_sheet_name, "A1", "7")
        self.assertEqual(self.workbook.get_cell_value(copied_sheet_name, "B1"), Decimal("14"))

    def test_preserve_original_sheet_case(self):
        # When copying a sheet, the original sheet's name should be preserved
        self.workbook.new_sheet("SheeT1")
        _, copied_sheet_name = self.workbook.copy_sheet("SheeT1")
        self.assertEqual(copied_sheet_name, "SheeT1_1")
        self.assertNotEqual(copied_sheet_name, "Sheet1_1")

    def test_copy_sheet_case_insensitive_name(self):
        self.workbook.new_sheet("TestSheet")
        _, copied_sheet_name = self.workbook.copy_sheet("testsheet")  # Using different case
        self.assertEqual(copied_sheet_name, "TestSheet_1")
        _, copied_sheet_name = self.workbook.copy_sheet("TESTSHEET")
        self.assertEqual(copied_sheet_name, "TestSheet_2")

    def test_uniqueness_case_insensitive(self):
        # If Sheet1 is copied, since SHEET1_1 already exists, and 
        # uniqueness is case insensitive, the copied sheet should be named
        # Sheet1_2
        self.workbook.new_sheet("Sheet1")
        self.workbook.new_sheet("SHEET1_1")
        _, copied_sheet_name = self.workbook.copy_sheet("Sheet1")
        self.assertEqual(copied_sheet_name, "Sheet1_2")

    def test_if_sheet_not_found_keyerror_is_raised(self):
        # If the sheet to be copied does not exist, a KeyError should be raised
        with self.assertRaises(KeyError):
            self.workbook.copy_sheet("Sheet1")

    def test_copy_returns_index_and_sheet_name(self):
        # If Sheet1 is copied, the index of the copied sheet and the name of the
        # copied sheet should be returned as a tuple. Index should be 2 as
        # sheet should be added to end of workbook, which is 0-indexed.
        self.workbook.new_sheet("Sheet1")
        self.workbook.new_sheet("Sheet2")
        index, name = self.workbook.copy_sheet("Sheet1")
        self.assertEqual(index, 2)
        self.assertEqual(name, "Sheet1_1")

    def test_copy_resolves_bad_reference_error_creates_circular_ref(self):
        # It is possible to construct sheets where, when copied, the values of 
        # the copy end up different from the values of the original sheet, even
        # though the contents of the two sheets are identical.

        # Consider case where Original!A1 refers to Original_1!A1. After copying
        # Original, the copied sheet Original_1 should have a circular reference
        # error in A1 and resolve the bad reference error in Original!A1.

        self.workbook.new_sheet("Original")
        # Should be a bad reference error since Original_1 does not exist
        self.workbook.set_cell_contents("Original", "A1", "=Original_1!A1")
        self.assertEqual(self.workbook.get_cell_value("Original", "A1").get_type(), CellErrorType.BAD_REFERENCE)
        
        _, copied_sheet_name = self.workbook.copy_sheet("Original")
        
        # Copied sheet at A1's = Original_1!A1 after copying -> self circular reference
        self.assertEqual(copied_sheet_name, "Original_1")
        self.assertEqual(self.workbook.get_cell_contents(copied_sheet_name, "A1"), "=Original_1!A1")
        self.assertIsInstance(self.workbook.get_cell_value(copied_sheet_name, "A1"), CellError)
        self.assertEqual(self.workbook.get_cell_value(copied_sheet_name, "A1").get_type(), CellErrorType.CIRCULAR_REFERENCE)
        
        # Original!A1 should now be a circular reference error since it refers to Original_1!A1, which references itself
        self.assertIsInstance(self.workbook.get_cell_value("Original", "A1"), CellError)
        self.assertEqual(self.workbook.get_cell_value("Original", "A1").get_type(), CellErrorType.CIRCULAR_REFERENCE)

    def test_copy_resolves_bad_reference(self):
        # As with adding and deleting sheets, copying a sheet could result in 
        # changes to the values in other sheets.

        # Copying a sheet could resolve some bad reference errors in other sheets

        self.workbook.new_sheet("Sheet1")
        self.workbook.new_sheet("Sheet2")
        self.workbook.set_cell_contents("Sheet1", "A1", "=Sheet2_1!A1")
        # Should be a bad reference error since Sheet2_1 does not exist
        self.assertIsInstance(self.workbook.get_cell_value("Sheet1", "A1"), CellError)
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A1").get_type(), CellErrorType.BAD_REFERENCE)

        # Copying Sheet2 should resolve the bad reference error in Sheet1
        self.workbook.copy_sheet("Sheet2")
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A1"), 0)

    def test_nested_bad_refs_fixed_after_multiple_copies(self):
        # Test nested bad references. Cell A1 in original sheet S1 depends on a 
        # copied sheet S2_1. Cell A2 in original sheet S1 depends on A1 and 
        # another copied sheet S2_1_1.

        self.workbook.new_sheet("Sheet1")
        self.workbook.new_sheet("Sheet2")
        # Refers to inexistent sheet
        self.workbook.set_cell_contents("Sheet1", "A1", "=Sheet2_1!A1")
        # Refers to above cell AND another inexistent sheet
        self.workbook.set_cell_contents("Sheet1", "A2", "=Sheet1!A1 + Sheet2_1_1!A1 + 5")
        
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A1").get_type(), CellErrorType.BAD_REFERENCE)
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A2").get_type(), CellErrorType.BAD_REFERENCE)

        _, copied_sheetname = self.workbook.copy_sheet("Sheet2")
        # Bad reference fixed for Sheet1!A1
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A1"), Decimal('0'))
        # Bad reference still not fixed since its 2nd reference still refers to inexistent sheet
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A2").get_type(), CellErrorType.BAD_REFERENCE)

        self.workbook.copy_sheet(copied_sheetname)
        # Bad reference finally fixed
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A2"), Decimal('5'))

    def test_copy_turns_bad_ref_into_div_error(self):
        # Test if a bad reference error in a copied sheet turns into a divide by
        # zero error in the original sheet
        self.workbook.new_sheet("Sheet1")
        self.workbook.new_sheet("Sheet2")
        self.workbook.set_cell_contents("Sheet1", "A1", "=2/Sheet2_1!A1")
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A1").get_type(), CellErrorType.BAD_REFERENCE)

        # Copying Sheet2 resolves the bad reference error in Sheet1
        # Sheet1!A1 should now be a divide by zero error
        self.workbook.copy_sheet("Sheet2")
        self.assertEqual(self.workbook.get_cell_value("Sheet1", "A1").get_type(), CellErrorType.DIVIDE_BY_ZERO)

    def tearDown(self):
        del self.workbook

if __name__ == '__main__':
    unittest.main()
