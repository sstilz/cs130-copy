import context
import unittest
from decimal import Decimal
from sheets.cell_error_type import CellErrorType, CellError
from sheets.workbook import Workbook
from sheets.formula_evaluator import FormulaEvaluator

class TestCellNotifications(unittest.TestCase):
    def setUp(self):
         self.called_cells = []
         self.called_cells2 = []
         self.called_cells3 = []
         self.order_ntfy_lst = []
         self.wb = Workbook()
         self.wb.notify_cells_changed(self.on_cells_changed)

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
        # every time this function called, will put A1's value in _retreived_cell_value list
        #weird behavior, this is called twice for each cell change. changed_cells is empty for one of the calls though.
        if len(changed_cells) != 0:
            self.called_cells3.append(workbook.get_cell_value(changed_cells[0][0], changed_cells[0][1]))
            
    def test_cell_notifcation_on_new_non_formula_cell(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "123")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])
    
    def test_cell_notification_simple_ref_with_implicit_cell_creation(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "'123")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        # B1 is implicitly created so should NOT trigger a cell notification
        self.wb.set_cell_contents(name, "C1", "=A1+B1")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'C1')])
    
    def test_cell_notification_implicit_cell_reference_modified(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "'123")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        self.wb.set_cell_contents(name, "C1", "=A1+B1")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'C1')])

        # B1 was implicitly created, but here its value is explicitly changed
        # from None to 5.3, so should trigger a cell notification
        self.wb.set_cell_contents(name, "B1", "5.3")
        expected_list = [('Sheet1', 'A1'), ('Sheet1', 'C1'), ('Sheet1', 'B1'), ('Sheet1', 'C1')]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list))
    
    def test_something_to_nothing(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "6")
        self.wb.set_cell_contents(name, "A1", None)
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

    def test_nothing_to_something(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", None)
        self.wb.set_cell_contents(name, "A1", "6")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

    def test_cell_update_for_value_change(self):
        #Check that if content changes, but values does not change, then there is no cell notification
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "= 6")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        # Since the value remains 6, no new notification should be sent.
        self.wb.set_cell_contents(name, "A1", "= 2 * 3")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        self.wb.set_cell_contents(name, "A1", "6")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

    def test_parse_error(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "6")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        self.wb.set_cell_contents(name, "A1", "=++6")  

        cell_error_object = self.wb.get_cell_value(name, "A1")
        self.assertTrue(isinstance(cell_error_object, CellError))
        self.assertTrue(cell_error_object.get_error_type_string() == "#ERROR!")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

    def test_cycle_error(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "6")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        self.wb.set_cell_contents(name, "B1", "=A1")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'B1')])

        self.wb.set_cell_contents(name, "C1", "=B1")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1')])

        self.wb.set_cell_contents(name, "A1", "=C1")

        cell_error_object = self.wb.get_cell_value(name, "A1")
        self.assertTrue(isinstance(cell_error_object, CellError))
        self.assertTrue(cell_error_object.get_error_type_string() == "#CIRCREF!")

        expected_list = [
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1'), 
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1')
            ]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list))

        self.wb.set_cell_contents(name, "A1", "6")
        expected_list_2 = [
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1'), 
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1'),
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1')
            ]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list_2))

    def test_add_sheet(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        # Referring to a sheet that doesn't exist.
        self.wb.set_cell_contents("Sheet1", "A1", "=Sheet2!A1")

        # Check cell has #REF! error and gets notified.
        cell_error_object = self.wb.get_cell_value(name, "A1")
        self.assertTrue(isinstance(cell_error_object, CellError))
        self.assertTrue(cell_error_object.get_error_type_string() == "#REF!")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])

        (index2, name2) = self.wb.new_sheet("Sheet2")
        self.assertTrue(self.wb.get_cell_value(name, "A1") == Decimal('0'))
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

    def test_delete_sheet(self):
        (index, name) = self.wb.new_sheet("Sheet1")
        (index2, name2) = self.wb.new_sheet("Sheet2")

        self.wb.set_cell_contents("Sheet1", "A1", "=Sheet2!A1")
        self.assertTrue(self.wb.get_cell_value(name, "A1") == Decimal('0'))
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])
    
        self.wb.del_sheet("Sheet2")
        cell_error_object = self.wb.get_cell_value(name, "A1")
        self.assertTrue(cell_error_object.get_error_type_string() == "#REF!")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

    def test_multiple_functions(self):
        self.wb.notify_cells_changed(self.on_cells_changed2)
        self.assertTrue(self.wb._get_notify_functions() == [self.on_cells_changed, self.on_cells_changed2])
        (index, name) = self.wb.new_sheet("Sheet1")

        # Note that on_cells_changed2 will NOT raise an excpetion as A1 is in called_cells2
        self.wb.set_cell_contents("Sheet1", "A1", "6")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])
        self.assertTrue(self.called_cells2 == [('Sheet1', 'A1')])

    def test_handles_exceptions(self):
        self.wb.notify_cells_changed(self.on_cells_changed2)
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A2", "=6")
        # As Sheet1!A1 is not in called_cells, then on_cells_changed2n 
        # will raise an exception, but on_cells_changed will still be called 
        
        self.assertTrue(self.called_cells == [('Sheet1', 'A2')])
        self.assertTrue(self.called_cells2 == []) # Should've thrown an error

    def test_something_to_something_else(self):
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "6")
        self.wb.set_cell_contents(name, "A1", "7")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

    def test_same_notification_function(self):
        # Add test for same function multiple times
        # i.e. if you register A, then B, then A.
        self.wb.notify_cells_changed(self.on_cells_changed2)
        self.wb.notify_cells_changed(self.on_cells_changed)
        self.assertTrue(self.wb._get_notify_functions() == [self.on_cells_changed, self.on_cells_changed2, self.on_cells_changed])
        (index, name) = self.wb.new_sheet("Sheet1")

        #Also testing here that if a notify function throws an exception, we handle it gracefully (by ignoring it)
        self.wb.set_cell_contents("Sheet1", "A1", "6")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])
        self.assertTrue(self.called_cells2 == [('Sheet1', 'A1')])


        self.wb.notify_cells_changed(self.on_cells_changed2)
        self.wb.set_cell_contents("Sheet1", "A1", "4")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1'), ('Sheet1', 'A1'), ('Sheet1', 'A1')])
        self.assertTrue(self.called_cells2 == [('Sheet1', 'A1'), ('Sheet1', 'A1'), ('Sheet1', 'A1')])

    def test_raise_notifications_for_change_to_formula(self):
        # Check that if a cell that is not a formula and then is made a formula then cell notification 
        # is raised both for the references and then the cell itself
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "6")
        self.wb.set_cell_contents(name, "B1", "=A1+2")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'B1')])

        self.wb.set_cell_contents(name, "C1", "5")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1')])

        self.wb.set_cell_contents(name, "A1", "=C1+3")
        expected_list = [('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1'), ('Sheet1', 'A1'), ('Sheet1', 'B1')]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list))

    def test_apply_and_unapply_formula_cell(self):
        #Check that if you update a cell to not a formula -> cell notification
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "A1", "=1+B1")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])
        self.wb.set_cell_contents(name, "A1", "3")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])
    
    def test_notification_called_after_change(self):
        (index, name) = self.wb.new_sheet("Sheet1")
        self.wb.notify_cells_changed(self.on_cells_changed3)
        self.wb.set_cell_contents(name, "A1", "1")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])
        self.assertTrue(self.called_cells3 == [1])
        self.wb.set_cell_contents(name, "A1", "3")
        self.assertTrue(self.called_cells3 == [1, 3])
        self.wb.set_cell_contents(name, "A1", "=5+B1")
        self.assertTrue(self.called_cells3 == [1, 3, 5])

    def test_notification_function_many(self):
        (index, name) = self.wb.new_sheet("Sheet1")
        self.wb.notify_cells_changed(self.on_cells_changed)
        self.wb.notify_cells_changed(self.on_cells_changed)
        self.wb.notify_cells_changed(self.on_cells_changed)
        self.wb.notify_cells_changed(self.on_cells_changed)
        self.wb.set_cell_contents(name, "A1", "1")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1'), ('Sheet1', 'A1'), ('Sheet1', 'A1'), ('Sheet1', 'A1')])

    #Helper fxns for test_notification_in_order
    def order1_ntfy_fxn(self, workbook, changed_cells):
        if len(changed_cells) != 0:
            self.order_ntfy_lst.append(1)
    def order2_ntfy_fxn(self, workbook, changed_cells):
        if len(changed_cells) != 0:
            self.order_ntfy_lst.append(2)
    def order3_ntfy_fxn(self, workbook, changed_cells):
        if len(changed_cells) != 0:
            self.order_ntfy_lst.append(3)
                
    def test_notification_in_order(self):
        #Check ntfy fxns called in order of registration
        (index, name) = self.wb.new_sheet("Sheet1")
        self.wb.notify_cells_changed(self.order1_ntfy_fxn)
        self.wb.notify_cells_changed(self.order2_ntfy_fxn)
        self.wb.notify_cells_changed(self.order3_ntfy_fxn)
        self.wb.set_cell_contents(name, "A1", "1")
        self.assertEqual(self.order_ntfy_lst, [1, 2, 3])
    
    def test_notification_error_prop(self):
        # Circular reference and parse errors test

        (index, name) = self.wb.new_sheet("Sheet1")
        self.wb.set_cell_contents(name, "A1", "1")
        self.wb.set_cell_contents(name, "B1", "=A1+1")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'B1')])

        #Create circ ref
        self.wb.set_cell_contents(name, "A1", "=B1+4")
        self.assertIsInstance(self.wb.get_cell_value(name, "A1"), CellError)
        self.assertEqual(self.wb.get_cell_value(name,"A1").get_type(), CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(self.wb.get_cell_value(name, "B1"), CellError)
        self.assertEqual(self.wb.get_cell_value(name,"B1").get_type(), CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(self.assert_lists_equal(self.called_cells, [
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), 
            ('Sheet1', 'A1'), ('Sheet1', 'B1')])
        )

        # Breaking cycle, should notify both A1 and B1
        self.wb.set_cell_contents(name, "A1", "=5")
        self.assertTrue(self.assert_lists_equal(self.called_cells, [
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), 
            ('Sheet1', 'A1'), ('Sheet1', 'B1'),
            ('Sheet1', 'A1'), ('Sheet1', 'B1')])
        )

        #Make sure no regular behavior broke (B1 update doesn't notify A1)
        self.wb.set_cell_contents(name, "B1", "=A1+2")
        self.assertTrue(self.assert_lists_equal(self.called_cells, [
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), 
            ('Sheet1', 'A1'), ('Sheet1', 'B1'),
            ('Sheet1', 'A1'), ('Sheet1', 'B1'),
            ('Sheet1', 'B1')])
        )

        #Create parse error
        self.wb.set_cell_contents(name, "A1", "=++2")
        self.assertTrue(self.assert_lists_equal(self.called_cells, [
            ('Sheet1', 'A1'), ('Sheet1', 'B1'), 
            ('Sheet1', 'A1'), ('Sheet1', 'B1'),
            ('Sheet1', 'A1'), ('Sheet1', 'B1'),
            ('Sheet1', 'B1'),
            ('Sheet1', 'A1'), ('Sheet1', 'B1')])
        )
    
    def test_implicit_conversion_chain(self):
        #Check implicit conversion (contents 0) for root and internal nodes
        (index, name) = self.wb.new_sheet("Sheet1")

        self.wb.set_cell_contents(name, "B1", "=A1")
        self.assertTrue(self.called_cells == [('Sheet1', 'B1')])
        self.wb.set_cell_contents(name, "C1", "=B1")
        self.assertTrue(self.called_cells == [('Sheet1', 'B1'), ('Sheet1', 'C1')])
        self.wb.set_cell_contents(name, "A1", "=7")
        expected_list = [('Sheet1', 'B1'), ('Sheet1', 'C1'), ('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1')]
        self.assertTrue(self.assert_lists_equal(self.called_cells, expected_list))

    def test_notify_copy(self):
        (index, name) = self.wb.new_sheet("Sheet1")
        self.wb.set_cell_contents(name, "A1", "=Sheet1_1!A2")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])
        _, name1 = self.wb.copy_sheet(name)
        self.assertTrue(self.assert_lists_equal(self.called_cells, [('Sheet1', 'A1'), ('Sheet1_1', 'A1'), ('Sheet1', 'A1')]))

    def test_notify_renaming(self):
        (index, name) = self.wb.new_sheet("Sheet1")
        self.wb.set_cell_contents(name, "A1", "=Sheet3!A2")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])
        (index, name2) = self.wb.new_sheet("Sheet2")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1')])
        self.wb.rename_sheet('Sheet2', 'Sheet3')
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])
    
    def test_notifications_for_deleted_and_cleared_cells(self):
        self.wb.new_sheet("Sheet1")
        
        # Test with literal
        self.wb.set_cell_contents("Sheet1", "A1", "5")
        self.wb.set_cell_contents("Sheet1", "A1", " ")
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])
        self.called_cells.clear()

        # Test with formula
        self.wb.set_cell_contents("Sheet1", "A1", "=A2")
        self.wb.set_cell_contents("Sheet1", "A1", None)
        self.assertTrue(self.called_cells == [('Sheet1', 'A1'), ('Sheet1', 'A1')])

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
