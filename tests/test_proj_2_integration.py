import context
import unittest
from decimal import Decimal
from sheets.cell_error_type import CellErrorType, CellError
from sheets.workbook import Workbook
from sheets.formula_evaluator import FormulaEvaluator
import os
import json

class TestIntegration(unittest.TestCase):
    def setUp(self):
         self.called_cells = []
         self.called_cells2 = []
         self.called_cells3 = []
         self.called_cells4 = []
         self.order_ntfy_lst = []
         self.wb = Workbook()
         self.wb.notify_cells_changed(self.on_cells_changed)
         self.script_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "example_workbooks") 

    def tearDown(self):
        del self.wb

    def on_cells_changed(self, workbook, changed_cells):
        '''
        This function gets called when cells change in the workbook that the
        function was registered on.  The changed_cells argument is an iterable
        of tuples; each tuple is of the form (sheet_name, cell_location).
        '''
        self.called_cells.extend(changed_cells)

    def on_cells_changed2(self, workbook, changed_cells):
        if ('Sheet1', 'A1') not in changed_cells:
            raise Exception("A1 not in cells changed")
        self.called_cells2.extend(changed_cells)
    
    def on_cells_changed3(self, workbook, changed_cells):
        # every time this function is called it adds a tuple for each of the changed
        # cells to the called_cells3 list in the form (sheet_name, cell_location, cell_value)
        # This way we can keep track of cell values at the time of notification
        for cell in changed_cells:
            cell_updated = (cell[0], cell[1], workbook.get_cell_value(cell[0], cell[1]))
            self.called_cells3.append(cell_updated)

    def on_cells_changed4(self, workbook, changed_cells):
        # Updates called_cells4 to only include the most recent cell change
        self.called_cells4 = changed_cells

    # Testing Cell Notifications for Most Complicated Rename Tests
    def test_bad_reference_disappears_after_creating_new_sheet(self):
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')

        # Reference Sheet3 which doesn't exist, creating a bad reference    
        self.wb.set_cell_contents('Sheet1', 'A1', '=Sheet3!A1')
        cell_error_object = self.wb.get_cell_value('Sheet1', "A1")
        assert(isinstance(cell_error_object, CellError))
        assert(cell_error_object.get_type() == CellErrorType.BAD_REFERENCE)

        # Creating Sheet1!A1 should create a cell notification for Sheet1!A1
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        # Renaming should fix the cell's bad reference.
        self.wb.rename_sheet('Sheet2', 'Sheet3')

        # Sheet1!A1 should now be a good reference, so should be a changed cell
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

        # Since '=Sheet3!A1' refers to an empty cell and doesn't specify how
        # it should be used, its value should be Decimal('0'). Note that it's
        # no longer a bad reference.
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), '=Sheet3!A1')
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1'), Decimal('0'))
        # Set Sheet3!A1 to 3 to check that the reference is working.
        self.wb.set_cell_contents('Sheet3', 'A1', "=1")
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1'), Decimal('1'))

        # Setting Sheet3!A1 should create a cell notification for Sheet3!A1 and Sheet1!A1 (through reference)
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1'), ('Sheet3', 'A1'), ('Sheet1', 'A1')])

    def test_rename_reflected_in_only_necessary_cell_references(self):

        # If a sheet has multiple references, we should only update the references
        # that involve the renamed sheet.
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2') # will rename this later
        self.wb.new_sheet('Sheet3')

        self.wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1')
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        self.wb.set_cell_contents('Sheet1', 'A2', '=Sheet3!A1')
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A2')])

        self.wb.set_cell_contents('Sheet2', 'A3', '=Sheet2!A1 * Sheet3!A1')
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A2'), ('Sheet2', 'A3')])

        self.wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1 + Sheet3!A1')

        # Creates a cycle between Sheet1!A1 and Sheet2!A1.
        # Sheet1!A1 and Sheet2!A1 should be notified because they become 
        # part of a cycle. Sheet2!A3 should be notified because it references
        # Sheet2!A1
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(self.wb.get_cell_value('Sheet2', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet2', 'A1').get_type(), CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(self.wb.get_cell_value('Sheet2', 'A3'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet2', 'A3').get_type(), CellErrorType.CIRCULAR_REFERENCE)
        expected_list = [
            ('Sheet1', 'A1'), ('Sheet1', 'A2'), ('Sheet2', 'A3'),
            ('Sheet2', 'A1'), ('Sheet1', 'A1'), ('Sheet2', 'A3')]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list))

        # since A4's contents are updated to not include Sheet2, 
        # A4's reference should not be updated after the rename
        self.wb.set_cell_contents('Sheet3', 'A4', '=Sheet1!A1 + Sheet3!A1')
        # No cycle is created, so only Sheet3!A4 should be notified
        expected_list2 = [
            ('Sheet1', 'A1'), ('Sheet1', 'A2'), ('Sheet2', 'A3'),
            ('Sheet2', 'A1'), ('Sheet1', 'A1'), ('Sheet2', 'A3'),
            ('Sheet3', 'A4')
        ]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list2))
        self.wb.rename_sheet('Sheet2', 'RenamedSheet2')
        # No cell values changed after the rename, no additional cells are notified
        expected_list3 =[
            ('Sheet1', 'A1'), ('Sheet1', 'A2'), ('Sheet2', 'A3'),
            ('Sheet2', 'A1'), ('Sheet1', 'A1'), ('Sheet2', 'A3'),
            ('Sheet3', 'A4')
        ]

        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list3))

        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1').replace(" ", ""), '=RenamedSheet2!A1')
        self.assertEqual(self.wb.get_cell_contents('RenamedSheet2', 'A3').replace(" ", ""), '=RenamedSheet2!A1*Sheet3!A1')
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A2').replace(" ", ""), '=Sheet3!A1')
        self.assertEqual(self.wb.get_cell_contents('RenamedSheet2', 'A1').replace(" ", ""), '=Sheet1!A1+Sheet3!A1')
        self.assertEqual(self.wb.get_cell_contents('Sheet3', 'A4').replace(" ", ""), '=Sheet1!A1+Sheet3!A1') 

    def test_rename_fixes_bad_references_in_multiple_formulas(self):

        formula = "=ASDF!A1"  # Reference to a non-existent sheet

        self.wb.new_sheet('newsheet')

        for i in range(10):
            self.wb.new_sheet(f'Sheet{i}')
            self.wb.set_cell_contents(f'Sheet{i}', 'A1', formula)
        expected_list = [
            ('Sheet0', 'A1'), ('Sheet1', 'A1'), ('Sheet2', 'A1'), 
            ('Sheet3', 'A1'), ('Sheet4', 'A1'), ('Sheet5', 'A1'), 
            ('Sheet6', 'A1'), ('Sheet7', 'A1'), ('Sheet8', 'A1'), 
            ('Sheet9', 'A1')]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list))

        for i in range(10):
            self.assertIsInstance(self.wb.get_cell_value(f'Sheet{i}', 'A1'), CellError)
            self.assertEqual(self.wb.get_cell_value(f'Sheet{i}', 'A1').get_type(), CellErrorType.BAD_REFERENCE)

        self.wb.rename_sheet('newsheet', 'ASDF')  # Renaming a sheet to ASDF

        for i in range(10):
            self.assertEqual(self.wb.get_cell_contents(f'Sheet{i}', 'A1'), formula)
            self.assertEqual(self.wb.get_cell_value(f'Sheet{i}', 'A1'), Decimal(0))

        # All of the cells should be notified because they change from bad reference
        # to 0s.
        expected_list2 = [
            ('Sheet0', 'A1'), ('Sheet1', 'A1'), ('Sheet2', 'A1'), 
            ('Sheet3', 'A1'), ('Sheet4', 'A1'), ('Sheet5', 'A1'), 
            ('Sheet6', 'A1'), ('Sheet7', 'A1'), ('Sheet8', 'A1'), 
            ('Sheet9', 'A1'), 
            ('Sheet0', 'A1'), ('Sheet1', 'A1'), ('Sheet2', 'A1'), 
            ('Sheet3', 'A1'), ('Sheet4', 'A1'), ('Sheet5', 'A1'), 
            ('Sheet6', 'A1'), ('Sheet7', 'A1'), ('Sheet8', 'A1'), 
            ('Sheet9', 'A1')]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list2))

    def test_renaming_creates_circular_reference_error(self):
        # Consider the case where renaming creates a circular reference
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A1', '=SheetBla!A1')
        self.wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1')

        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet2', 'A1')])

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

        # Sheet1!A1 and SheetBla!A1 should be notified because they become part of a cycle
        # Note that Sheet2!A1 is not notified because it gets renamed to SheetBla!A1, which is notified
        self.assertTrue(self.assert_lists_equal(self.called_cells, [('Sheet1', 'A1'), ('Sheet2', 'A1'), ('Sheet1', 'A1'), ('SheetBla', 'A1')]))

        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(self.wb.get_cell_value('SheetBla', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('SheetBla', 'A1').get_type(), CellErrorType.CIRCULAR_REFERENCE)

    def test_renaming_creates_divide_by_zero_error(self):
        # Consider the case where renaming creates a divide by zero error
        self.wb.new_sheet('Sheet1')
        self.wb.new_sheet('Sheet2')
        self.wb.set_cell_contents('Sheet1', 'A1', '=2/SheetBla!A1')

        # Sheet1!A1 should be notified because it's a new cell
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        # Sheet1 references SheetBla, which doesn't exist, so it's a bad reference.
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.BAD_REFERENCE)

        # Renaming Sheet2 should create a divide by zero error
        # because Sheet1!A1 references 2/SheetBla!A1, which references 2/0.
        self.wb.rename_sheet('Sheet2', 'SheetBla')
        self.assertEqual(self.wb.get_cell_contents('Sheet1', 'A1'), '=2/SheetBla!A1')
        self.assertIsInstance(self.wb.get_cell_value('Sheet1', 'A1'), CellError)
        self.assertEqual(self.wb.get_cell_value('Sheet1', 'A1').get_type(), CellErrorType.DIVIDE_BY_ZERO)

        # Since Sheet2 was renamed to SheetBla, then Sheet1!A1 should be notified as it's
        # turns from a bad reference to a divide by zero error
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

    def test_load_json_with_cell_notification(self):
        # Load a new WB with two sheets - 1 with cells, 1 without cells.
        # Then test that cell notifications work properly for this new WB.

        # *************** LOADING THE WORKBOOK ***************
        sheet_name = "Sheet1"
        empty_sheet_name = "EmptySheet"
        self.wb.new_sheet(sheet_name)
        self.wb.new_sheet(empty_sheet_name)
        cell_contents = [
            ["a1", '1'],
            ["b2", '2'],
        ]
        # Only add cells for "Sheet1". Deliberately don't add cells for "EmptySheet".
        for cell in cell_contents:
            self.wb.set_cell_contents(sheet_name, cell[0], cell[1])
        json_file_path = os.path.join(self.script_directory, "wb_test_empty_sheet.json")
        
        # Check that the saved JSON is correct.
        with open(json_file_path, 'w+') as newfile:
            self.wb.save_workbook(newfile)
        with open(json_file_path, 'r+') as newfile:
            d = json.load(newfile)
            self.assertTrue("sheets" in d)
            self.assertTrue("Sheet1" == d["sheets"][0]["name"])
            # Check that "EmptySheet" is still in the saved JSON.
            self.assertTrue("EmptySheet" == d["sheets"][1]["name"])
        
        # Check that we get the correct WB object after reloading that JSON.
        # Specifically, "EmptySheet" should still show up as a WB sheet, even
        # though it doesn't have cells.
        with open(json_file_path, 'r+') as newfile:
            new_wb = Workbook.load_workbook(newfile)
            self.assertEqual(new_wb.list_sheets(), ["Sheet1", "EmptySheet"])

        # *************** TESTING CELL NOTIFICATIONS ***************
        # The new WB object is created inside load_wb(), which internally calls
        # set_cell_contents(), which gathers updated cells and notifies right
        # away at the end of the function. Because there are no notif_functions
        # set up at that point, we'll lose the notifs for the 2 cells created.
        self.called_cells.clear()   # clear because all WBs share this field
        self.assertTrue(self.called_cells == [])
        new_wb.notify_cells_changed(self.on_cells_changed)
        self.assertTrue(new_wb.notify_functions == [self.on_cells_changed])

        # Check cell notifs work properly for newly loaded WB.
        expected_called_cells = []
        new_wb.set_cell_contents(sheet_name, "a1", "2") # change val
        expected_called_cells.append((sheet_name, 'A1'))
        new_wb.set_cell_contents(sheet_name, "b2", "=a1")   # doesn't change val
        self.assertTrue(self.called_cells == expected_called_cells)

        self.called_cells.clear()
        new_wb.set_cell_contents(sheet_name, "a1", "=b2") # circ ref
        # We only only care that these cells are notified. The order should not
        # be enforced (the order is nondeterministic anyways).
        self.assertTrue(set(self.called_cells) == set([(sheet_name, 'A1'), (sheet_name, 'B2')]))

        self.called_cells.clear()
        new_wb.set_cell_contents(sheet_name, "a1", "=emptySheet!c1")
        # C1 shouldn't be called since it was implicitly created and has val None
        # S1!A1 and S1!B2 should be notified since their CIRC_REFs are resolved
        self.assertTrue(set(self.called_cells) == set([(sheet_name, 'A1'), (sheet_name, 'B2')]))

        # Check that notif works for the empty sheet too.
        self.called_cells.clear()
        new_wb.set_cell_contents(empty_sheet_name, "c2", "1")
        self.assertTrue(self.called_cells == [(empty_sheet_name, 'C2')])

    def test_delete_sheet_with_cell_notification(self):
        # Test cell notifs with circular references, implicit references,
        # unchanged values, and deleted sheets.
        self.wb.new_sheet("Sheet1")
        self.wb.new_sheet("Sheet2")

        # Sanity check for notifs. Notice the standardized sheetname and loc.
        self.wb.set_cell_contents("Sheet1", "a1", "1")
        self.wb.set_cell_contents("SHEET1", "B2", "2")
        self.assertTrue(self.called_cells == [("Sheet1", "A1"), ("Sheet1", "B2")])

        # Check no notifs for unchanged values.
        self.called_cells.clear()
        self.wb.set_cell_contents("Sheet1", "b2", "=2*a1")
        self.assertTrue(self.called_cells == [])

        # Create ref going FROM Sheet2, which will be deleted later.
        self.called_cells.clear()
        self.wb.set_cell_contents("Sheet1", "C1", "=Sheet2!C1")
        self.assertTrue(self.wb.get_cell_value("Sheet1", "C1") == Decimal('0'))
        # Notice S2!C1 is implicitly created and has value None, so it's not notified.
        self.assertTrue(self.called_cells == [("Sheet1", "C1")])

        # Create ref going INTO Sheet2.
        self.called_cells.clear()
        self.assertTrue(self.wb.get_cell_value("Sheet1", "A1") == Decimal('1'))
        self.wb.set_cell_contents("Sheet2", "C1", "=Sheet1!A1 + 5")
        self.assertTrue(self.wb.get_cell_value("Sheet2", "C1") == Decimal('6'))
        # S2!C1 was implicitly created, but now it's value has explicitly changed, so notify.
        # S1!C1 references S2!C1 so it's also notified.
        self.assertTrue(self.called_cells == [("Sheet2", "C1"), ("Sheet1", "C1")])

        # Create cell in Sheet2 with no references.
        self.called_cells.clear()
        self.wb.set_cell_contents("Sheet2", "B1", "=3")
        self.assertTrue(self.wb.get_cell_value("Sheet2", "B1") == Decimal('3'))
        self.assertTrue(self.called_cells == [("Sheet2", "B1")])

        # Create a cell whose value doesn't change despite referencing S2.
        self.called_cells.clear()
        self.wb.set_cell_contents("Sheet1", "D1", "=Sheet2!E1")
        self.assertTrue(self.wb.get_cell_value("Sheet1", "D1") == Decimal('0'))
        self.assertTrue(self.called_cells == [("Sheet1", "D1")])

        # ************************ DELETE THE SHEET ************************
        self.assertTrue(self.wb.get_cell_contents("Sheet2", "C1") == "=Sheet1!A1 + 5")
        self.called_cells.clear()
        self.wb.del_sheet("Sheet2")

        # We don't notify S1!A1 even though a deleted cell once referenced it.
        self.assertTrue(("Sheet1", "A1") not in self.called_cells)

        # Recall S2!B1 was explicitly created in S2. However, we don't care
        # that it was deleted since it was never referenced by other sheets.
        self.assertTrue(("Sheet2", "B1") not in self.called_cells)

        # These cells depended on cells that were deleted, so notify.
        self.assertTrue(self.wb.get_cell_contents("Sheet1", "D1") == "=Sheet2!E1")
        self.assertTrue(self.wb.get_cell_value("Sheet1", "D1").get_type() == CellErrorType.BAD_REFERENCE)
        
        self.assertTrue(self.wb.get_cell_contents("Sheet1", "C1") == "=Sheet2!C1")
        self.assertTrue(self.wb.get_cell_value("Sheet1", "C1").get_type() == CellErrorType.BAD_REFERENCE)
        
        self.assertTrue(set(self.called_cells) == set([("Sheet1", "D1"), ("Sheet1", "C1")]))

        # Now create a new sheet with that deleted sheet's name.
        # Check that the bad references resulting from the deleted sheet are fixed.
        self.called_cells.clear()
        self.wb.new_sheet("Sheet2")
        self.assertTrue(self.wb.get_cell_contents("Sheet1", "D1") == "=Sheet2!E1")
        self.assertTrue(self.wb.get_cell_contents("Sheet1", "C1") == "=Sheet2!C1")
        self.assertTrue(set(self.called_cells) == set([("Sheet1", "D1"), ("Sheet1", "C1")]))
    
    def test_copy_sheet_with_cell_notifications(self):
        """
        Test cell notifications with circular references, bad references, and 
        resolved references due to sheet copying.
        """
        self.wb.new_sheet("Sheet1")
        self.wb.new_sheet("Sheet2")

        # Set initial cell contents in Sheet1 and Sheet2.
        self.wb.set_cell_contents("Sheet1", "A1", "1")
        self.wb.set_cell_contents("Sheet1", "B1", "=A1")
        self.wb.set_cell_contents("Sheet1", "C1", "=D1")
        self.wb.set_cell_contents("Sheet2", "A1", "=Sheet1!B1")
        # Notice D1 is not in here because it's an implicit reference.
        self.assertTrue(self.called_cells == [("Sheet1", "A1"), ("Sheet1", "B1"), ("Sheet1", "C1"), ("Sheet2", "A1")])
        self.called_cells.clear()

        # Create bad reference from Sheet2 to Sheet1_1 (sheet DNE).
        self.assertTrue(self.wb.get_cell_value("Sheet2", "A1") == 1)
        self.wb.set_cell_contents("Sheet2", "A1", "=Sheet1_1!A1")
        self.assertTrue(self.wb.get_cell_value("Sheet2", "A1").get_type() == CellErrorType.BAD_REFERENCE)
        self.assertTrue(self.called_cells == [("Sheet2", "A1")])
        self.called_cells.clear()
        
        # Copy Sheet1 into Sheet1_1.
        self.wb.copy_sheet("Sheet1")

        # Bad reference from Sheet2 to Sheet1_1 should be fixed.
        self.assertTrue(self.wb.get_cell_contents("Sheet2", "A1") == "=Sheet1_1!A1")
        self.assertTrue(self.wb.get_cell_contents("Sheet1_1", "A1") == "1")
        self.assertTrue(self.wb.get_cell_value("Sheet2", "A1") == Decimal('1'))
        
        # When copying, we call new_sheet, which calls evaluate, checking for
        # all references with the sheetname. Sheet2!A1 references Sheet1_1!A1,
        # so it's updated and its bad reference goes away.
        self.assertTrue(self.called_cells[0] == ('Sheet2', 'A1'))
        # Then we proceed with the actual process of copying the cells, so we
        # first call set_cell_contents() on Sheet1_1!A1, which is referenced by
        # Sheet2!A1. So we notify Sheet1_1!A1 and Sheet2!A1 (again)
        self.assertTrue(self.called_cells[1:] == [('Sheet1_1', 'A1'), ('Sheet2', 'A1'), ('Sheet1_1', 'B1'), ('Sheet1_1', 'C1')])
        # No notif for cell D1 since it's an implicit reference
        self.assertTrue(("Sheet1_1", "D1") not in self.called_cells)

        # Create a circular reference within Sheet1_1.
        self.called_cells.clear()
        self.wb.set_cell_contents("Sheet1_1", "D1", "=C1")
        self.assertTrue(self.wb.get_cell_value("Sheet1_1", "D1").get_type() == CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(self.wb.get_cell_value("Sheet1_1", "C1").get_type() == CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(set(self.called_cells) == set([("Sheet1_1", "D1"), ("Sheet1_1", "C1")]))
        self.called_cells.clear()

        # Resolve the circular reference within Sheet1_1.
        self.wb.set_cell_contents("Sheet1_1", "C1", "2")
        # We expect notifications for C1 (the changed cell), and D1 (which no longer has a circular reference)
        self.assertTrue(set(self.called_cells) == set([("Sheet1_1", "C1"), ("Sheet1_1", "D1")]))
        # Check that the cells' values are as expected
        self.assertEqual(self.wb.get_cell_contents("Sheet1_1", "C1"), "2")
        self.assertEqual(self.wb.get_cell_contents("Sheet1_1", "D1"), "=C1")
        self.called_cells.clear()
    
    def test_rename_delete_move_copy_with_cell_notifications(self):
        self.wb.new_sheet("Sheet1")
        self.wb.new_sheet("Sheet2")
        self.wb.set_cell_contents("Sheet1", "A1", "=Sheet2!A1")
        self.assertTrue(self.called_cells == [("Sheet1", "A1")])
        self.called_cells.clear()

        self.wb.rename_sheet("Sheet2", "Sheet2_1")
        # sheet doesn't exist now
        self.assertTrue(self.wb.get_cell_contents("Sheet1", "A1") == "=Sheet2_1!A1")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1"), Decimal('0'))
        self.assertTrue(self.called_cells == [])    # value didn't change
        self.called_cells.clear()

        self.wb.del_sheet("Sheet2_1")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1").get_type(), CellErrorType.BAD_REFERENCE)
        self.assertTrue(self.called_cells == [("Sheet1", "A1")])    # value didn't change
        self.called_cells.clear()

        self.wb.new_sheet("Sheet2")
        # Contents still references the renamed sheet that is now deleted, so
        # value remains bad reference.
        self.assertTrue(self.wb.get_cell_contents("Sheet1", "A1") == "=Sheet2_1!A1")
        self.assertTrue(self.called_cells == [])

        self.wb.move_sheet("Sheet2", 0)
        self.assertTrue(self.called_cells == [])    # move sheet doesn't change cell values

        # Resolves Sheet1!A1's bad reference to the deleted copy sheet.
        self.wb.copy_sheet("Sheet2")
        self.assertTrue(self.wb.get_cell_contents("Sheet1", "A1") == "=Sheet2_1!A1")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1"), Decimal('0'))
        self.assertTrue(self.called_cells == [("Sheet1", "A1")])

    def test_cell_notifications_A2B2C_in_copy(self):
        """
        Tests when a cell's value changes to A, then to B, then to C, during 
        the course of loading a workbook.
        """
        self.wb.new_sheet("Sheet1")
        self.wb.set_cell_contents("Sheet1", "A1", "=1 + A2")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1"), Decimal('1'))  # val A
        self.wb.set_cell_contents("Sheet1", "A2", "=3 + A3")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1"), Decimal('4'))  # val B
        self.wb.set_cell_contents("Sheet1", "A3", "5")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1"), Decimal('9'))  # val C
        self.assertEqual(self.called_cells, [('Sheet1', 'A1'), ('Sheet1', 'A2'), ('Sheet1', 'A1'), ('Sheet1', 'A3'), ('Sheet1', 'A2'), ('Sheet1', 'A1')])
        self.called_cells.clear()

        # When copying this sheet, Sheet1_1!A1's value will change many times
        # before it gets to the final value of 9, but our current implementation
        # will still send notifications throughout the process.
        self.wb.copy_sheet("Sheet1")
        self.assertEqual(self.wb.get_cell_value("Sheet1_1", "A1"), Decimal('9'))  # val C
        self.assertEqual(self.called_cells, [('Sheet1_1', 'A1'), ('Sheet1_1', 'A2'), ('Sheet1_1', 'A1'), ('Sheet1_1', 'A3'), ('Sheet1_1', 'A2'), ('Sheet1_1', 'A1')])
    
    def test_cell_notifications_A2B2A_in_copy(self):
        """
        Tests when a cell's value changes to A, then to B, then back to A, during 
        the course of loading a workbook.
        """
        self.wb.new_sheet("Sheet1")
        self.wb.set_cell_contents("Sheet1", "A1", "=1 + A2")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1"), Decimal('1'))  # val A
        self.wb.set_cell_contents("Sheet1", "A2", "=3 + A3")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1"), Decimal('4'))  # val B
        self.wb.set_cell_contents("Sheet1", "A3", "-3")
        self.assertEqual(self.wb.get_cell_value("Sheet1", "A1"), Decimal('1'))  # back to val A
        self.assertEqual(self.called_cells, [('Sheet1', 'A1'), ('Sheet1', 'A2'), ('Sheet1', 'A1'), ('Sheet1', 'A3'), ('Sheet1', 'A2'), ('Sheet1', 'A1')])
        self.called_cells.clear()

        # When copying this sheet, Sheet1_1!A1's value will change many times
        # before it gets to the same place where it started, but our current 
        # implementation will still send notifications throughout the process.
        self.wb.copy_sheet("Sheet1")
        self.assertEqual(self.wb.get_cell_value("Sheet1_1", "A1"), Decimal('1'))  # val A
        self.assertEqual(self.called_cells, [('Sheet1_1', 'A1'), ('Sheet1_1', 'A2'), ('Sheet1_1', 'A1'), ('Sheet1_1', 'A3'), ('Sheet1_1', 'A2'), ('Sheet1_1', 'A1')])

    def assert_lists_equal(self, orig_list, new_list):
        orig_list_counts = {}
        new_list_counts = {}
        for l in orig_list:
            orig_list_counts[l] = orig_list_counts.get(l, 0) + 1
        for l in new_list:
            new_list_counts[l] = new_list_counts.get(l, 0) + 1
        return orig_list_counts == new_list_counts
    
if __name__ == '__main__':
    unittest.main()
