import context
import unittest
from decimal import Decimal
from sheets.cell_error_type import CellErrorType, CellError
from sheets.workbook import Workbook
from sheets.formula_evaluator import FormulaEvaluator

class TestRenameSheets(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()

    def test_rename_to_same_sheet_raises_error(self):
        self.wb.new_sheet('Sheet1')
        # Renaming a sheet to itself should raise a ValueError because it's not unique.
        with self.assertRaises(ValueError):
            self.wb.rename_sheet('Sheet1', 'Sheet1')
    
    def test_rename_to_existent_sheet_raises_error(self):
        # Can't rename to sheet name that's already used.
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')

        # Case sensitive rename
        with self.assertRaises(Exception):
            self.wb.rename_sheet('Sheet1', 'Sheet2')
        
        # Case insensitive rename
        with self.assertRaises(Exception):
            self.wb.rename_sheet('Sheet1', 'sheet2')

    def test_error_raised_if_sheet_name_not_found(self):
        with self.assertRaises(KeyError):
            self.wb.rename_sheet('Sheet1', 'Sheet2')
    
    def test_error_raised_after_accessing_sheetname_before_rename(self):
        self.wb.new_sheet('Sheet1')
        self.wb.rename_sheet('Sheet1', 'Sheet2')

        with self.assertRaises(Exception):
            self.wb.set_cell_contents('Sheet1', 'A5', '5')
    
    def test_bad_reference_disappears_after_creating_new_sheet(self):
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')

        # Reference Sheet3 which doesn't exist.
        self.wb.set_cell_contents('Sheet1', 'A1', '=Sheet3!A1')
        cell_error_object = self.wb.get_cell_value('Sheet1', "A1")
        assert(isinstance(cell_error_object, CellError))

        # Bad reference
        assert(cell_error_object.get_error_type_string() == "#REF!")

        # Renaming should fix the cell's bad reference.
        self.wb.rename_sheet('Sheet2', 'Sheet3')

        # Since '=Sheet3!A1' refers to an empty cell and doesn't specify how
        # it should be used, its value should be Decimal('0'). Note that it's
        # no longer a bad reference.
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), '=Sheet3!A1')
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1'), Decimal('0'))
        # Set Sheet3!A1 to 3 to check that the reference is working.
        self.wb.set_cell_contents('Sheet3', 'A1', "=1")
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1'), Decimal('1'))
    
    def test_rename_to_invalid_sheet_raises_error(self):
        # These renamed sheet names should already be caught by the new_sheet 
        # method, but just in case, we'll check again here.
        self.wb.new_sheet('Sheet1')

        # Whitespace
        with self.assertRaises(ValueError):
            self.wb.rename_sheet('Sheet1', ' Sheet 2 ')
        
        # Invalid punctuation
        with self.assertRaises(ValueError):
            self.wb.rename_sheet('Sheet1', 'Sheet<2>')

        # Empty string
        with self.assertRaises(ValueError):
            self.wb.rename_sheet('Sheet1', '')
    
    def test_rename_preserves_case_of_new_name(self):
        # Test the case of the new name is preserved by the workbook
        self.wb.new_sheet('Sheet1')
        self.wb.rename_sheet('Sheet1', 'sHeet2B')
        self.assertEqual(self.wb.list_sheets(), ['sHeet2B'])
    
    def test_sheet_list_reflects_multiple_renames(self):
        # After renaming a sheet multiple times, only the final name of the
        # sheet should be in the list
        self.wb.new_sheet('Sheet1')
        self.wb.rename_sheet('Sheet1', 'Sheet2')
        self.wb.rename_sheet('Sheet2', 'Sheet3')
        self.assertEqual(self.wb.list_sheets(), ['Sheet3'])
    
    def test_single_references_updated_with_new_name(self):
        # If Sheet1 references Sheet2!A1, and Sheet2 is renamed to RenamedSheet2,
        # then Sheet1 should now reference RenamedSheet2!A1
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1')
        self.wb.rename_sheet('Sheet2', 'RenamedSheet2')
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), '=RenamedSheet2!A1')

    def test_multiple_references_updated_with_new_name(self):
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A5', '5')
        self.wb.set_cell_contents('Sheet2', 'A6', '6')
        self.wb.set_cell_contents('Sheet1', 'A1', "='Sheet1'!A5+'Sheet2'!A6")

        self.wb.rename_sheet('Sheet2', 'SheetBla')

        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), "=Sheet1!A5+SheetBla!A6")
    
    def test_reference_to_same_sheet_updated_with_new_name(self):
        self.wb.new_sheet('Sheet1')

        # It's weird for a sheet to reference a cell in the same sheet with this
        # notation, but it's technically valid, so test it.
        self.wb.set_cell_contents('Sheet1', 'A1', '=Sheet1!A2')
        self.wb.rename_sheet('Sheet1', 'RenamedSheet1')

        try:
            content = self.wb.get_cell_contents('RenamedSheet1', 'A1')
        except KeyError:
            self.fail("RenamedSheet1 does not exist")
        self.assertEqual(content, '=RenamedSheet1!A2')
    
    def test_rename_reflected_in_only_necessary_cell_references(self):
        # If a sheet has multiple references, we should only update the references
        # that involve the renamed sheet.
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2') # will rename this later
        self.wb.new_sheet('Sheet3')

        self.wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1')
        self.wb.set_cell_contents('Sheet1', 'A2', '=Sheet3!A1')

        self.wb.set_cell_contents('Sheet2', 'A3', '=Sheet2!A1 * Sheet3!A1')
        self.wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1 + Sheet3!A1')

        # since A4's contents are updated to not include Sheet2, 
        # A4's reference should not be updaetd after the rename.
        self.wb.set_cell_contents('Sheet3', 'A4', '=Sheet2!A1+Sheet3!A1')
        self.wb.set_cell_contents('Sheet3', 'A4', '=Sheet1!A1 + Sheet3!A1')

        self.wb.rename_sheet('Sheet2', 'RenamedSheet2')
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1').replace(" ", ""), '=RenamedSheet2!A1')
        self.assertEqual(self.wb.get_cell_contents('RenamedSheet2', 'A3').replace(" ", ""), '=RenamedSheet2!A1*Sheet3!A1')
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A2').replace(" ", ""), '=Sheet3!A1')
        self.assertEqual(self.wb.get_cell_contents('RenamedSheet2', 'A1').replace(" ", ""), '=Sheet1!A1+Sheet3!A1')
        self.assertEqual(self.wb.get_cell_contents('Sheet3', 'A4').replace(" ", ""), '=Sheet1!A1+Sheet3!A1')  

    def test_single_quotes_not_added_to_sheetname(self):
        self.wb.new_sheet('Sheet1')
        self.wb.rename_sheet('Sheet1', 'Sheet-Bla')
        first_sheet_object = self.wb.worksheet_order[0] # the first WS object

        # Don't add single quotes to sheet names with hyphens.
        # We only add quotes for formulas.
        self.assertNotEqual(first_sheet_object.sheet_name, "'Sheet-Bla'")
    
    def test_preserve_parentheses_simple(self):
        # Renaming a sheet should preserve the parentheses in formulas
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        formula = "=(Sheet2!A1 + Sheet2!A2) * Sheet2!A3"
        self.wb.set_cell_contents('Sheet1', 'A1', formula)
        
        self.wb.rename_sheet('Sheet2', 'SheetBla')        
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1').replace(" ", ""), formula.replace(" ", "").replace("Sheet2", "SheetBla"))

    def test_preserve_parentheses_complex(self):
        # Renaming a sheet should preserve the parentheses in formulas
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        formula = "=(Sheet2!A1 + Sheet2!A2) * (Sheet2!A3 + (Sheet2!A4 + Sheet2!A5) * Sheet2!A6 + Sheet2!A7 + Sheet2!A8) * Sheet2!A9"
        self.wb.set_cell_contents('Sheet1', 'A1', formula)
        
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1').replace(" ", ""), formula.replace(" ", "").replace("Sheet2", "SheetBla"))

    def test_preserve_parentheses_nested(self):
        # Renaming a sheet should preserve the parentheses in nested formulas
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.new_sheet('Sheet3')
        formula = "=((Sheet2!A1) + ((Sheet3!A1) + (Sheet3!A2)) * (Sheet3!A3))"
        self.wb.set_cell_contents('Sheet1', 'A1', formula)
        
        self.wb.rename_sheet('Sheet3', 'SheetNested')
        
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1').replace(" ", ""), formula.replace(" ", "").replace("Sheet3", "SheetNested"))

    def test_preserve_parentheses_multiple_sheets(self):
        # Renaming multiple sheets should preserve the parentheses in formulas
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.new_sheet('Sheet3')
        formula = "=(Sheet2!A1 + (Sheet3!A1 + Sheet3!A2)) * Sheet3!A3"
        self.wb.set_cell_contents('Sheet1', 'A1', formula)
        
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        self.wb.rename_sheet('Sheet3', 'SheetNested')
        
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), formula.replace("Sheet2", "SheetBla").replace("Sheet3", "SheetNested").replace(" ", ""))

    def test_preserve_parentheses_multiple_sheets_complex(self):
        # Renaming multiple sheets should preserve the parentheses in complex formulas
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.new_sheet('Sheet3')
        formula = "=((Sheet2!A1 + (Sheet3!A1 + Sheet3!A2)) * (Sheet3!A3 + (Sheet2!A4 + Sheet2!A5) * Sheet2!A6 + Sheet2!A7 + Sheet2!A8) * Sheet3!A9)"
        self.wb.set_cell_contents('Sheet1', 'A1', formula)
        
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        self.wb.rename_sheet('Sheet3', 'SheetNested')
        
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), formula.replace("Sheet2", "SheetBla").replace("Sheet3", "SheetNested").replace(" ", ""))

    def test_unparseable_formula(self):
        # Do not bother updating unparseable formulas
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A1', "=^-Sheet2!A1+")
        
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        
        # assert that cell A1's contents are not 
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.PARSE_ERROR)
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), "=^-Sheet2!A1+")

    def test_unparseable_formula_with_multiple_sheets(self):
        # Do not bother renaming a sheet if the formula is unparseable
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.new_sheet('Sheet3')
        self.wb.new_sheet('Sheet4')
        self.wb.set_cell_contents('Sheet1', 'A1', "=^*Sheet2!A1 + Sheet3!A1 + Sheet4!A1")
        
        self.wb.rename_sheet('Sheet4', 'SheetBla')
        
        # assert that cell A1's contents are not 
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.PARSE_ERROR)
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), "=^*Sheet2!A1 + Sheet3!A1 + Sheet4!A1")
    
    def test_unparseable_formula_with_multiple_sheets_complex(self):
        # Do not bother renaming a sheet if the formula is unparseable
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.new_sheet('Sheet3')
        self.wb.new_sheet('Sheet4')
        
        formula = "=Sheet2!A1) + (Sheet3!A1 + Sheet4!A1) * (Sheet1!A1 * Sheet2!A2 + Sheet3!A3 + Sheet4!A4"

        self.wb.set_cell_contents('Sheet1', 'A1', formula)
        self.wb.rename_sheet('Sheet4', 'SheetBla')
        
        # assert that cell A1's contents are not 
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.PARSE_ERROR)
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), formula)

    def test_false_sheet_name(self):
        # Do not bother renaming a sheet if the formula is unparseable
        self.wb.new_sheet('sheet1!a2')
        self.wb.new_sheet('Sheet2')
        self.wb.new_sheet('Sheet1')
        self.wb.set_cell_contents('Sheet2', 'A1', "='sheet1!a2'!A1")
        
        # Renaming Sheet1 should not affect 'sheet1!a2' because it's not the same sheet.
        self.wb.rename_sheet('Sheet1', 'SheetBla')
        self.assertEqual(self.wb.get_cell_contents('Sheet2', 'A1'), "='sheet1!a2'!A1")

        self.wb.rename_sheet('sheet1!a2', 'sheet2!B1_sheet1!a2')
        self.assertEqual(self.wb.get_cell_contents('Sheet2', 'A1'), "='sheet2!B1_sheet1!a2'!A1")

    def test_rename_fixes_bad_references(self):
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A1', "='Sheet-2'!A1 + 0")

        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.BAD_REFERENCE)

        self.wb.rename_sheet('Sheet2', 'Sheet-2')

        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1').replace(" ", ""), "='Sheet-2'!A1+0")
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1'), Decimal(0))

    def test_rename_fixes_bad_references_in_multiple_formulas(self):

        formula = "=ASDF!A1"  # Reference to a non-existent sheet

        self.wb.new_sheet('newsheet')

        for i in range(10):
            self.wb.new_sheet(f'Sheet{i}')
            self.wb.set_cell_contents(f'Sheet{i}', 'A1', formula)

        for i in range(10):
            self.assertIsInstance(self.wb.get_cell_value(f'Sheet{i}', 'A1'), CellError)
            self.assertEqual(self.wb.get_cell_value(f'Sheet{i}', 'A1').get_type(), CellErrorType.BAD_REFERENCE)

        self.wb.rename_sheet('newsheet', 'ASDF')  # Renaming a sheet to ASDF

        for i in range(10):
            self.assertEqual(self.wb.get_cell_contents(f'Sheet{i}', 'A1'), formula)
            self.assertEqual(self.wb.get_cell_value(f'Sheet{i}', 'A1'), Decimal(0))

    def test_rename_does_not_remove_necessary_quotes_from_other_references(self):
        name = "Sheet-1"
        name2 = "Sheet-2"
        self.wb.new_sheet(name)
        self.wb.new_sheet(name2)
        self.wb.set_cell_contents(name, "A5", "5")
        self.wb.set_cell_contents(name2, "A6", "6")
        # Here the references require single quotes.
        self.wb.set_cell_contents(name, "A1", "='Sheet-1'!A5+'Sheet-2'!A6")

        # This rename should not remove the necessary single quotes around 
        # 'Sheet-1'.
        self.wb.rename_sheet(name2, "SheetBla")
        assert(self.wb.get_cell_contents(name, "A1") == "='Sheet-1'!A5+SheetBla!A6")
    

    def test_cells_not_explicitly_referring_to_renamed_sheet_remain_unchanged(self):
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')

        # Referring to cell in same sheet.
        self.wb.set_cell_contents('Sheet1', 'A1', '=B1')

        # Reference to cell in a different sheet that's about to be renamed.
        self.wb.set_cell_contents('Sheet1', 'C1', '=Sheet2!A2')
        self.wb.rename_sheet('Sheet2', 'NewSheet')

        # Check that the contents of cell A1 in Sheet1 have not changed.
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), '=B1')

    def test_new_single_quotes(self):
        # sheet's old name does not require single-quotes but "new name" does
        # check that we only use single quotes if necessary
        pass

    def test_allowed_to_add_sheet_with_old_name_used_before_rename(self):
        # If a sheet is renamed, then a new sheet with the old name should be allowed.
        self.wb.new_sheet('Sheet1')
        self.wb.rename_sheet('Sheet1', 'Sheet2')
        self.wb.new_sheet('Sheet1')
        self.assertEqual(self.wb.list_sheets(), ['Sheet2', 'Sheet1'])

    def test_formula_literal_looks_like_cell_ref_is_not_udpated(self):
        # If a formula looks like a cell reference, but is actually a literal,
        # then it should not be updated.
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A1', '=\"Sheet2!A1\"')
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        # The formula should not be updated as it is a literal, not a cell reference.
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), '=\"Sheet2!A1\"')

    def test_adding_sheet_with_old_name_fixes_bad_ref_errors(self):
        # If a sheet is renamed, then a new sheet with the old name should fix
        # bad references to the original sheet.
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        self.wb.set_cell_contents('Sheet1', 'A2', '=Sheet2!A1')
        # The reference to Sheet2 should be a bad reference as Sheet2 no longer exists.
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A2'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A2').get_type(), CellErrorType.BAD_REFERENCE)
        # Adding a new sheet with the old name should fix the bad reference.
        self.wb.new_sheet('Sheet2')
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A2'), Decimal('0'))

    def test_renaming_creates_circular_reference_error(self):
        # Consider the case where renaming creates a circular reference
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A1', '=SheetBla!A1')
        self.wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1')
        # Sheet1 references SheetBla, which doesn't exist, so it's a bad reference.
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.BAD_REFERENCE)
        # Sheet2!A1 references Sheet1!A1 so it's also a bad reference error
        self.assertIsInstance(self.wb.get_cell_value('Sheet2', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet2', 'A1').get_type(), CellErrorType.BAD_REFERENCE)

        # Renaming Sheet2 should create a circular reference error
        # because Sheet2!A1 references Sheet1!A1, which references SheetBla!A1.
        # Once Sheet2 is renamed, it will create a cycle
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), '=SheetBla!A1')
        self.assertEqual(self.wb.get_cell_contents('SheetBla', 'A1'), '=Sheet1!A1')

        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(self.wb.get_cell_value('SheetBla', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('SheetBla', 'A1').get_type(), CellErrorType.CIRCULAR_REFERENCE)

    def test_renaming_creates_divide_by_zero_error(self):
        # Consider the case where renaming creates a divide by zero error
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A1', '=2/SheetBla!A1')

        # Sheet1 references SheetBla, which doesn't exist, so it's a bad reference.
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.BAD_REFERENCE)

        # Renaming Sheet2 should create a divide by zero error
        # because Sheet1!A1 references 2/SheetBla!A1, which references 2/0.
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), '=2/SheetBla!A1')
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.DIVIDE_BY_ZERO)

class TestRenameSheetSingleQuotes(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()

    def test_sheet_no_quotes_renamed_to_something_needing_quotes(self):
        # No quotes -> needs quotes
        self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents("Sheet1", "A1", "=Sheet1!A2")     # no quotes
        self.wb.rename_sheet("Sheet1", "Sheet-1")   # needs quotes bc hyphen

        # Reference must include single quote since rename included hyphen.
        self.assertEqual(self.wb.get_cell_contents("Sheet-1", "A1"), "='Sheet-1'!A2")
    
    def test_sheet_no_quotes_renamed_to_something_not_needing_quotes(self):
         # No quotes -> doesn't need quotes
        self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents("Sheet1", "A1", "=Sheet1!A2")     # no quotes
        self.wb.rename_sheet("Sheet1", "Sheet2")   # doesn't need quotes

        # Reference must include single quote since rename included hyphen.
        self.assertEqual(self.wb.get_cell_contents("Sheet2", "A1"), "=Sheet2!A2")
                
    def test_sheet_unnecessary_quotes_renamed_to_something_needing_quotes(self):
        self.wb.new_sheet("Sheet1")

        # unnecessary quotes
        self.wb.set_cell_contents("Sheet1", "A1", "='Sheet1'!A2")
        self.wb.rename_sheet("Sheet1", "Sheet-1")   # needs quote

        self.assertEqual(self.wb.get_cell_contents("Sheet-1", "A1"), "='Sheet-1'!A2")
    
    def test_sheet_with_unnecessary_renamed_to_something_not_needing_quotes(self):
        self.wb.new_sheet("Sheet1")

        # unnecessary quotes
        self.wb.set_cell_contents("Sheet1", "A1", "='Sheet1'!A2")
        self.wb.rename_sheet("Sheet1", "Sheet2")   # doesn't need quotes

        self.assertEqual(self.wb.get_cell_contents("Sheet2", "A1"), "=Sheet2!A2") # quotes should be removed
    
    def test_sheet_with_necessary_quotes_renamed_to_something_needing_quotes(self):
        self.wb.new_sheet("Sheet-1")
        self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents("Sheet1", "A1", "='Sheet-1'!A2")
        self.wb.rename_sheet("Sheet-1", "Sheet-2")

        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet-2'!A2")

    def test_sheet_with_necessary_quotes_renamed_to_something_not_needing_quotes(self):
        self.wb.new_sheet("Sheet-1")
        self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents("Sheet1", "A1", "='Sheet-1'!A2")
        self.wb.rename_sheet("Sheet-1", "Sheet_1")

        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "=Sheet_1!A2")
    
    def test_unrelated_references_with_unnecessary_quotes_not_changed(self):
        self.wb.new_sheet("Sheet1")
        self.wb.new_sheet("Sheet2")

        # Unnecessary quotes, but shouldn't be changed in the below rename
        # since it's not related.
        self.wb.set_cell_contents("Sheet1", "A1", "='Sheet1'!A2")

        # Unnecessary quotes, and SHOULD be affected by rename.
        self.wb.set_cell_contents("Sheet1", "A2", "='Sheet2'!A2")

        # References cell in same sheet - shouldn't be affected by rename
        self.wb.set_cell_contents("Sheet1", "A4", "=Sheet1!A3")
        
        self.wb.rename_sheet("Sheet2", "NewSheet2")

        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet1'!A2")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A2"), "=NewSheet2!A2")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A4"), "=Sheet1!A3")

    def test_addition_of_single_quotes_applies_to_all_relevant_references(self):
        self.wb.new_sheet("Sheet1")
        self.wb.new_sheet("Sheet2")
        self.wb.new_sheet("Sheet3")

        self.wb.set_cell_contents("Sheet3", "A2", "='Sheet1'!A5 + 'Sheet2'!A6")
        self.wb.rename_sheet("Sheet2", "SheetBla")

        self.assertEqual(self.wb.get_cell_contents("Sheet3", "A2").replace(" ", ""), "=Sheet1!A5+SheetBla!A6")

    def test_addition_or_removal_of_single_quotes_applied_individually_to_cell_references(self):
        # Example from Piazza Post
        # If we have a formula ='Sheet1'!A9 + 'Sheet-2'!B4 - Sheet3!C6
        # Then "Sheet3" was renamed to "Sheet THREE!"
        # Formula is updated to =Sheet1!A9 + 'Sheet-2'!B4 - 'Sheet THREE!'!C6

        self.wb.new_sheet("Sheet1")
        self.wb.new_sheet("Sheet-2")
        self.wb.new_sheet("Sheet3")
        self.wb.set_cell_contents("Sheet1", "A1", "='Sheet1'!A9 + 'Sheet-2'!B4 - Sheet3!C6")
        self.wb.rename_sheet("Sheet3", "Sheet THREE!")

        # Should remove uncessary quotes around Sheet1
        # Keep necessary quotes around Sheet-2
        # Add necessary quotes around Sheet THREE! (note that missing space between Sheet and THREE because of the replace method)
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1").replace(" ", ""), "=Sheet1!A9+'Sheet-2'!B4-'SheetTHREE!'!C6")

    def test_single_quotes_added_when_necessary(self):
        # Single quotes are necessary if a sheet name doesn't start with alphabetical character
        # or underscore
        # Single quotes are necessary if a sheet contains spaces or any other characters besides
        # A-Z, a-z, 0-9, and underscore.
        # Can use .?!,:;!@#$%^&*()-

        # Maybe unnecessary, should be checked in test_cell, but testing more extensively, just in case
        self.wb.new_sheet("Sheet1")
        self.wb.new_sheet("Sheet2")

        self.wb.set_cell_contents("Sheet1", "A1", "=Sheet2!A1")
        
        self.wb.rename_sheet("Sheet2", "1Sheet2")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='1Sheet2'!A1")
        self.wb.rename_sheet("1Sheet2", "Sheet2.")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2.'!A1")
        self.wb.rename_sheet("Sheet2.", "Sheet2?")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2?'!A1")
        self.wb.rename_sheet("Sheet2?", "Sheet2,")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2,'!A1")
        self.wb.rename_sheet("Sheet2,", "Sheet2:")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2:'!A1")
        self.wb.rename_sheet("Sheet2:", "Sheet2;")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2;'!A1")
        self.wb.rename_sheet("Sheet2;", "Sheet2!")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2!'!A1")
        self.wb.rename_sheet("Sheet2!", "Sheet2@")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2@'!A1")
        self.wb.rename_sheet("Sheet2@", "Sheet2#")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2#'!A1")
        self.wb.rename_sheet("Sheet2#", "Sheet2$")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2$'!A1")
        self.wb.rename_sheet("Sheet2$", "Sheet2%")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2%'!A1")
        self.wb.rename_sheet("Sheet2%", "Sheet2^")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2^'!A1")
        self.wb.rename_sheet("Sheet2^", "Sheet2&")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2&'!A1")
        self.wb.rename_sheet("Sheet2&", "Sheet2*")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2*'!A1")
        self.wb.rename_sheet("Sheet2*", "Sheet2(")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2('!A1")
        self.wb.rename_sheet("Sheet2(", "Sheet2-")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2-'!A1")
        self.wb.rename_sheet("Sheet2-", "Sheet2)")
        self.assertEqual(self.wb.get_cell_contents("Sheet1", "A1"), "='Sheet2)'!A1")

if __name__ == '__main__':
    unittest.main()
