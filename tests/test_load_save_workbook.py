from collections import OrderedDict
import context
import tempfile
import unittest
from decimal import Decimal
import json

from sheets.cell_error_type import CellErrorType, CellError
from sheets.workbook import Workbook
from sheets.formula_evaluator import FormulaEvaluator
from sheets.workbook_utility import index_to_cell_location
import os


class TestWorkbookLoadSave(unittest.TestCase):
    def setUp(self):
        self.workbook = Workbook()
        self.script_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "example_workbooks") 
    
    def test_load_only_simple(self):
        # Manually create a JSON file that represents a workbook.
        json_file_path = os.path.join(self.script_directory, "wb_load_only_simple.json")
        with open(json_file_path, 'w+') as json_file:
            json.dump({
                "sheets": [
                    {
                        "name": "Sheet1",
                        "cell-contents": {
                            "A1": "1"
                        }
                    }
                ]
            }, json_file)

        # Load the workbook from the JSON file
        with open(json_file_path, 'r') as json_file:
            loaded_workbook = self.workbook.load_workbook(json_file)

        # Check that the loaded workbook matches the expected workbook
        assert loaded_workbook.get_cell_contents("Sheet1", 'A1') == "1"
    
    def test_load_only_complex(self):
        """
         Checks that sheets appear in same order as JSON, sheet name's 
         capitalization is preserved, and contents of every non-empty cell is 
         loaded.
         """
        # Path to the JSON file that represents a workbook
        json_file_path = os.path.join(self.script_directory, "wb_load_only_complex.json")

        # Manually create a JSON file that represents a workbook
        with open(json_file_path, 'w+') as json_file:
            json.dump({
                "sheets": [
                    {
                        "name": "Sheet1",
                        "cell-contents": {
                            "A1": "1",
                            "B2": "2",
                            "C3": "3"
                        }
                    },
                    {
                        "name": "ShEet2",
                        "cell-contents": {
                            "D4": "4",
                            "E5": "",
                            "F6": "6"
                        }
                    },
                    {
                        "name": "Sheet3",
                        "cell-contents": {
                            "G7": "=A2*A3"
                        }
                    }
                ]
            }, json_file)

        # Load the workbook from the JSON file
        with open(json_file_path, 'r') as json_file:
            loaded_workbook = self.workbook.load_workbook(json_file)

        # Check sheet order and sheet capitialization is preserved
        loaded_sheets = []
        for ws_obj in loaded_workbook.worksheet_order:
            loaded_sheets.append(ws_obj.sheet_name)
        assert loaded_sheets == ['Sheet1', 'ShEet2', 'Sheet3']

        # Define expected cell contents for each sheet
        expected_contents = {
            'Sheet1': { (0, 0): "1", (1, 1): "2", (2, 2): "3" },
            'ShEet2': { (3, 3): "4", (4, 4): None, (5, 5): "6" },
            'Sheet3': { (6, 6): "=A2*A3" }
        }

        # Check cell CONTENTS are properly loaded (order doesn't matter)
        for ws_obj in loaded_workbook.worksheet_order:
            cell_map = ws_obj.cell_map
            expected = expected_contents[ws_obj.sheet_name]
            for key in expected:
                assert cell_map[key].content == expected[key]

    def test_save_only_simple(self):
        # Create a workbook
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", "1")

        # Path to the JSON file to save the workbook to
        json_file_path = os.path.join(self.script_directory, "wb_save_only_simple.json")

        # Save the workbook to a JSON file
        with open(json_file_path, 'w+') as json_file:
            self.workbook.save_workbook(json_file)

        # Load the JSON file and check its contents
        with open(json_file_path, 'r') as json_file:
            saved_workbook_json = json.load(json_file)

        expected_workbook_json = {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "A1": "1"
                    }
                }
            ]
        }

        assert saved_workbook_json == expected_workbook_json
        
    def test_only_nonempty_cells_are_saved(self):
        # When given a WB with some empty and some non-empty cells, only the 
        # non-empty cells are present in the saved JSON.

        # Create the WB and convert it to JSON.
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", "1")
        self.workbook.set_cell_contents("Sheet1", "B2", "") # empty
        self.workbook.set_cell_contents("Sheet1", "B3") # empty

        json_file_path = os.path.join(self.script_directory, "wb_save_only_nonempty_cells.json")
        with open(json_file_path, 'w+') as json_file:
            self.workbook.save_workbook(json_file)

        with open(json_file_path, 'r') as json_file:
            saved_workbook = json.load(json_file)

        # The saved workbook's JSON should not include empty cells:
        expected_workbook = {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "A1": "1"
                    }
                }
            ]
        }
        assert saved_workbook == expected_workbook

    def test_sheet_with_no_cells_still_has_explicitly_empty_cell_contents(self):
        """
        If a sheet has no cells set, then there should be an empty cell-contents map.
        """
        self.workbook.new_sheet("EmptySheet")

        json_file_path = os.path.join(self.script_directory, "wb_save_sheet_with_no_cells.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)

        # Load the JSON data from the file
        with open(json_file_path, "r") as file:
            actual_json = json.load(file)
        
        expected_json = {
            "sheets": [
                {
                    "name": "EmptySheet",
                    "cell-contents": {}
                }
            ]
        }

        # Check that the cell-contents field is an empty map for the sheet
        self.assertEqual(expected_json, actual_json)

    def test_save_load_simple(self):
        # Create a file and save
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", "1")
        json_file_path = os.path.join(self.script_directory, "wb_load_save_simple.json")

        # This dumps the workbook to a JSON file into that json_file_path
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)
        
        # This converts the JSON back into a workbook.
        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)
        
        # Theoretically, if we're able to get back exactly what we started with,
        # then the save and load functions are working correctly (at least when
        # called together).
        self.assertEqual(loaded_workbook.get_cell_contents("Sheet1", 'A1'), "1")

    def test_workbook_order(self):
        # Create a file and save
        sheet_name_order = [f"sheet{i}" for i in range(100)]
        for sn in sheet_name_order:
            self.workbook.new_sheet(sn)
        json_file_path = os.path.join(self.script_directory, "wb_test_workbook_order.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)
        
        with open(json_file_path, 'r+') as newfile:
            self.assertEqual(Workbook.load_workbook(newfile).list_sheets(), sheet_name_order)

    def test_preserve_capitalization(self):
        sheet_name = "sHeETfHeQ"
        self.workbook.new_sheet(sheet_name)

        json_file_path = os.path.join(self.script_directory, "wb_preserve_capitalization.json")

        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)

        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)
            self.assertIn(sheet_name, loaded_workbook.list_sheets())

    def test_nonempty_cells(self):
        # Create a workbook object to compare against later.
        sheet_name = "Sheet1"
        self.workbook.new_sheet(sheet_name)
        cell_contents = [
            ["A1", "Value1"],
            ["B2", "Value2"],
            ["C3", "Value3"],
        ]
        for cell in cell_contents:
            self.workbook.set_cell_contents(sheet_name, cell[0], cell[1])

        # Convert that workbook to a JSON file.
        json_file_path = os.path.join(self.script_directory, "wb_test_nonempty_cells.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)

        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)
            loaded_contents = [[cell[0], loaded_workbook.get_cell_contents(sheet_name, cell[0])] for cell in cell_contents]
            self.assertEqual(loaded_contents, cell_contents)
    
    def test_load_workbook_works_when_cell_contents_field_is_before_name(self):
        # Cell-contents field before name
        json_file_path = os.path.join(self.script_directory, "wb_cell_contents_after_name.json")
        with open(json_file_path, 'w') as switched_field_order_json:
            json.dump({
                "sheets":[
                    {
                        "cell-contents":{
                            "A1":"'123",
                            "B1":"5.3",
                            "C1":"=A1*B1"
                        },
                        "name":"Sheet1"
                    },
                    {
                        "cell-contents":{
                            "D1":"4"
                        },
                        "name":"Sheet2"
                    }
                ]
            }, switched_field_order_json)

        with open(json_file_path, 'r+') as newfile:
            loaded_wb = Workbook.load_workbook(newfile)
            
            # Check sheet order is preserved
            loaded_sheets = []
            for ws_obj in loaded_wb.worksheet_order:
                loaded_sheets.append(ws_obj.sheet_name)
            assert loaded_sheets == ['Sheet1', 'Sheet2']

            # Check cells are properly loaded in for each sheet (order doesn't 
            # matter), despite the "cell-contents" field being before the "name"
            # in the JSON
            ws_obj1 = loaded_wb.worksheet_order[0]
            ws_obj2 = loaded_wb.worksheet_order[1]
            cell_map1 = ws_obj1.cell_map
            cell_map2 = ws_obj2.cell_map

            # Check cell_map for Sheet1
            assert cell_map1[(0, 0)].content == "'123"  # A1
            assert cell_map1[(0, 1)].content == "5.3"   # B1
            assert cell_map1[(0, 2)].content == "=A1*B1" # C1

            # Check cell_map for Sheet2
            assert cell_map2[(0, 3)].content == "4" # D1
            
    def test_empty_cells(self):
        # Create a file and save
        sheet_name = "Sheet1"
        self.workbook.new_sheet(sheet_name)
        cell_contents = [
            ["A1", "Value1"],
            ["B2", ""],
            ["C3", "Value3"],
            ["D4", " "],
            ["E5", None],
        ]
        for cell in cell_contents:
            self.workbook.set_cell_contents(sheet_name, cell[0], cell[1])
        json_file_path = os.path.join(self.script_directory, "wb_test_empty_cells.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)

        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)
            loaded_contents = []
            # Changed to collect all cells in sheet, because not all cells are necessarily saved
            for loc, cell in loaded_workbook._get_sheet(sheet_name).cell_map.items():
                loaded_contents.append([index_to_cell_location(loc[0], loc[1]), cell.content])
            self.assertEqual(loaded_contents, [["A1", "Value1"], ["C3", "Value3"]])
   
    def test_load_double_quotes_in_formulas(self):
        # Create WB with double quotes in formula and save to JSON.
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", '="e" & "f"')
        self.workbook.set_cell_contents("Sheet1", "A2", '"abc"')

        json_file_path = os.path.join(self.script_directory, "wb_load_double_quotes_in_formulas.json")
        with open(json_file_path, 'w+') as json_file:
            self.workbook.save_workbook(json_file)

        # Check the JSON created from above looks like the expected JSON below.
        with open(json_file_path, 'r') as json_file:
            saved_workbook_json = json.load(json_file)

        expected_workbook_json = {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "A1": "=\"e\" & \"f\"",  # double quotes are escaped
                        "A2": "\"abc\""
                    }
                }
            ]
        }

        assert saved_workbook_json == expected_workbook_json
        
    def test_double_quotes_without_formulas(self):
        # Test double quotes without formulas using the load and save functions
        # together.

        # Create a file and save with double-quoted cell content
        sheet_name = "Sheet1"
        self.workbook.new_sheet(sheet_name)
        cell_contents = [
            ["A1", 'Value with "quotes"'],
            ["B2", 'Another "quoted" value'],
        ]
        for cell in cell_contents:
            self.workbook.set_cell_contents(sheet_name, cell[0], cell[1])
        json_file_path = os.path.join(self.script_directory, "wb_test_double_quotes_without_formulas.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)

        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)
            loaded_contents = [[cell[0], loaded_workbook.get_cell_contents(sheet_name, cell[0])] for cell in cell_contents]
            self.assertEqual(loaded_contents, cell_contents)

    def test_input_cell_format(self):
        # Reading cell contents should support upper- and lowercase cell names.
        # Create a file and save with upper and lower case cell names
        sheet_name = "Sheet1"
        self.workbook.new_sheet(sheet_name)
        cell_contents = [
            ["A1", "Value1"],
            ["b2", "Value2"],
            ["C3", "Value3"],
        ]
        for cell in cell_contents:
            self.workbook.set_cell_contents(sheet_name, cell[0], cell[1])
        json_file_path = os.path.join(self.script_directory, "wb_test_input_cell_format.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)

        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)
            loaded_contents = [[cell[0], loaded_workbook.get_cell_contents(sheet_name, cell[0])] for cell in cell_contents]
            self.assertEqual(loaded_contents, cell_contents)

    def test_output_cell_format_is_all_uppercase(self):
        sheet_name = "Sheet1"
        self.workbook.new_sheet(sheet_name)
        cell_contents = [
            ["a1", '1'],
            ["b2", '2'],
        ]
        for cell in cell_contents:
            self.workbook.set_cell_contents(sheet_name, cell[0], cell[1])
        json_file_path = os.path.join(self.script_directory, "wb_test_output_cell_format.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)

        with open(json_file_path, 'r+') as newfile:
            loaded_contents = json.load(newfile)
            self.assertEqual(
                loaded_contents,
                {
                    'sheets': [{
                        'name': 'Sheet1',
                        'cell-contents': {"A1": '1', "B2": '2'}
                    }]
                }
            )

    def test_empty_workbook_still_outputs_empty_sheet(self):
        # Even though self.workbook has no sheets, the saved JSON should still
        # have `sheets: []`.
        json_file_path = os.path.join(self.script_directory, "wb_test_empty_workbook.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)

        # Check the resulting JSON file still has "[]"
        with open(json_file_path, 'r+') as newfile:
            actual_json = json.load(newfile)
            expected_json = {"sheets": []}
            assert actual_json == expected_json
        
        # Check that when converting back to a WS object, this doesn't
        # accidentally add more sheets.
        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)
            self.assertEqual(loaded_workbook.list_sheets(), [])

    def test_sheets_with_no_cells_still_saved(self):
        # Tests that sheets without cells are still included in the saved JSON.
        sheet_name = "Sheet1"
        empty_sheet_name = "EmptySheet"
        self.workbook.new_sheet(sheet_name)
        self.workbook.new_sheet(empty_sheet_name)
        cell_contents = [
            ["a1", '1'],
            ["b2", '2'],
        ]
        for cell in cell_contents:
            self.workbook.set_cell_contents(sheet_name, cell[0], cell[1])
        json_file_path = os.path.join(self.script_directory, "wb_test_empty_sheet.json")
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)
        
        with open(json_file_path, 'r+') as newfile:
            d = json.load(newfile)
            self.assertTrue("sheets" in d)
            self.assertTrue("Sheet1" == d["sheets"][0]["name"])
            # Check that "EmptySheet" is still in the saved JSON.
            self.assertTrue("EmptySheet" == d["sheets"][1]["name"])
        
        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)
            self.assertEqual(loaded_workbook.list_sheets(), ["Sheet1", "EmptySheet"])

    def test_malformed_json_no_content(self):
        # This JSON is malformed because the required "cell-contents" field is
        # missing (and replaced with "data").
        malformed_json_content = {
            "sheets": [
                {
                    "name": "Sheet1", 
                    "data": [
                        {"A1": "Value1"}, 
                        {"B2": "Value2"}
                    ]
                }
            ]
        }
        json_file_path = os.path.join(self.script_directory, "test_malformed_json_no_content.json")
        with open(json_file_path, 'w+') as newfile:
            json.dump(malformed_json_content, newfile)

        # Attempt to load the malformed JSON and expect an exception
        with self.assertRaises(Exception) as context:
            with open(json_file_path, 'r+') as newfile:
                Workbook.load_workbook(newfile)
        self.assertIn("Malformed JSON", str(context.exception))

    def test_malformed_missing_comma(self):
        # Test case for missing comma in the "data" list
        malformed_json_content = '''
        {
            "sheets": [
                {
                    "name": "Sheet1", 
                    "data": [
                        {"A1": "Value1"}
                        {"B2": "Value2"}
                    ]
                }
            ]
        }
        '''
        json_file_path = os.path.join(self.script_directory, "test_malformed_missing_comma.json")
        
        with open(json_file_path, 'w+') as newfile:
            newfile.write(malformed_json_content)

        # Attempt to load the malformed JSON and expect an exception
        with self.assertRaises(json.JSONDecodeError) as context:
            with open(json_file_path, 'r+') as newfile:
                Workbook.load_workbook(newfile)

    def test_malformed_missing_double_quotes_around_key(self):
        # Test case for missing double quotes around the "name" field/key
        malformed_json_content = '''
        {
            "sheets": [
                {
                    name: "Sheet1",
                    "cell-contents": {
                        "A1": "1"
                    }
                }
            ]
        }
        '''
        json_file_path = os.path.join(self.script_directory, "test_malformed_missing_double_quotes_around_key.json")
        
        with open(json_file_path, 'w+') as newfile:
            newfile.write(malformed_json_content)

        # Attempt to load the malformed JSON and expect an exception
        with self.assertRaises(json.JSONDecodeError) as context:
            with open(json_file_path, 'r+') as newfile:
                Workbook.load_workbook(newfile)

    def test_malformed_mismatched_opening_and_closing_brackets(self):
        # Test case for missing closing brackets for the "sheets" list
        malformed_json_content = '''
        {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "A1": "1"
                    }
                }
            
        }
        '''

        json_file_path = os.path.join(self.script_directory, "test_malformed_mismatched_brackets.json")
        
        with open(json_file_path, 'w+') as newfile:
            newfile.write(malformed_json_content)

        # Attempt to load the malformed JSON and expect an exception
        with self.assertRaises(json.JSONDecodeError):
            with open(json_file_path, 'r+') as newfile:
                Workbook.load_workbook(newfile)

    def test_malformed_extra_comma_in_the_list(self):
        # Test case for an extra comma in the "cell-contents" field
        malformed_json_content = '''
        {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "A1": "1",
                        "B2": "2",
                    }
                }
            ]
        }
        '''

        json_file_path = os.path.join(self.script_directory, "test_malformed_extra_comma.json")
        
        with open(json_file_path, 'w+') as newfile:
            newfile.write(malformed_json_content)

        # Attempt to load the malformed JSON and expect an exception
        with self.assertRaises(json.JSONDecodeError):
            with open(json_file_path, 'r+') as newfile:
                Workbook.load_workbook(newfile)

    def test_malformed_incorrect_colon_placement(self):
        # Test case for missing colon in between "name" and "Sheet1"
        malformed_json_content = '''
        {
            "sheets": [
                {
                    "name" "Sheet1",
                    "cell-contents": {
                        "A1": "1",
                        "B2": "2"
                    }
                }
            ]
        }
        '''

        json_file_path = os.path.join(self.script_directory, "test_malformed_incorrect_colon_placement.json")
        
        with open(json_file_path, 'w+') as newfile:
            newfile.write(malformed_json_content)

        # Attempt to load the malformed JSON and expect an exception
        with self.assertRaises(json.JSONDecodeError):
            with open(json_file_path, 'r+') as newfile:
                Workbook.load_workbook(newfile)

    def test_missing_required_field_cell_contents(self):
        # Prepare and write a JSON string with missing required fields to a file
        # Missing "cell-contents" field
        missing_field_json = '''
        {
            "sheets": [
                {
                    "name": "Sheet1"
                }
            ]
        }  
        '''
        json_file_path = os.path.join(self.script_directory, "wb_missing_field.json")
        
        with open(json_file_path, 'w') as newfile:
            newfile.write(missing_field_json)

        # Attempt to load the workbook from the file
        with open(json_file_path, 'r') as newfile:
            with self.assertRaises(KeyError) as context:
                Workbook.load_workbook(newfile)
        
        self.assertIn("Malformed JSON", str(context.exception), "Loading a workbook with missing required fields should raise KeyError with appropriate message.")
    
    def test_missing_required_field_name(self):
        # Prepare and write a JSON string with missing required "name" field
        # (it is instead swapped for "name2")
        missing_field_json = '''
        {
            "sheets": [
                {
                    "name2": "Sheet1",
                    "cell-contents": {
                        "A1": "1",
                        "B2": "2"
                    }
                }
            ]
        }
        '''
        json_file_path = os.path.join(self.script_directory, "wb_missing_field_name.json")
        
        with open(json_file_path, 'w') as newfile:
            newfile.write(missing_field_json)

        # Attempt to load the workbook from the file
        with open(json_file_path, 'r') as newfile:
            with self.assertRaises(KeyError) as context:
                Workbook.load_workbook(newfile)
        self.assertIn("Malformed JSON", str(context.exception))

    def test_missing_required_field_sheets(self):
        # Prepare and write a JSON string with missing required "sheets" field
        # (it is instead swapped for "sheets2")
        missing_field_json = '''
        {
            "sheets2": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "A1": "1",
                        "B2": "2"
                    }
                }
            ]
        }'''
        json_file_path = os.path.join(self.script_directory, "wb_missing_field_sheets.json")
        
        with open(json_file_path, 'w') as newfile:
            newfile.write(missing_field_json)

        # Attempt to load the workbook from the file
        with open(json_file_path, 'r') as newfile:
            with self.assertRaises(KeyError) as context:
                Workbook.load_workbook(newfile)
        self.assertIn("Malformed JSON", str(context.exception))
    
    def test_wrong_type(self):
        # Prepare and write a JSON string with the wrong type for cell contents
        wrong_type_json = {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "A1": 123 # Cell content should be a string, not a number
                    }
                }
            ]
        }  
        json_file_path = os.path.join(self.script_directory, "wb_wrong_type.json")
        
        with open(json_file_path, 'w') as newfile:
            json.dump(wrong_type_json, newfile)

        # Attempt to load the workbook from the file
        with open(json_file_path, 'r') as newfile:
            with self.assertRaises(TypeError) as context:
                Workbook.load_workbook(newfile)
        
        self.assertIn("Incorrect data type", str(context.exception), "Loading a workbook with wrong data type for cell contents should raise TypeError.")

    def test_invalid_cell_reference_beyond_max_extent(self):
        # Tests that load_workbook refuses to load a workbook with a cell 
        # reference beyond the max extent. 

        # Note: `load_workbook` logic accounts for this because it will look at 
        # every JSON cell and call set_cell_contents with it, which will call 
        # add_cell, which will raise an exception if the cell's location is 
        # invalid.
        invalid_cell_reference_json = {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "ZZZZ99999": "=2" # invalid cell reference: beyond max extent
                    }
                }
            ]
        }
        json_file_path = os.path.join(self.script_directory, "wb_invalid_cell_reference_beyond_max_extent.json")
        
        with open(json_file_path, 'w') as newfile:
            json.dump(invalid_cell_reference_json, newfile)

        with open(json_file_path, 'r') as newfile:
            with self.assertRaises(ValueError) as context:
                Workbook.load_workbook(newfile)
            self.assertIn("Invalid cell location", str(context.exception), "Setting cell contents with an invalid reference should raise ValueError.")
    
    def test_invalid_cell_reference(self):
        invalid_cell_reference_json = {
            "sheets": [
                {
                    "name": "Sheet1",
                    "cell-contents": {
                        "1A": "=2" 
                    }
                }
            ]
        }
        json_file_path = os.path.join(self.script_directory, "wb_invalid_cell_reference.json")
        
        with open(json_file_path, 'w') as newfile:
            json.dump(invalid_cell_reference_json, newfile)

        with open(json_file_path, 'r') as newfile:
            with self.assertRaises(ValueError) as context:
                Workbook.load_workbook(newfile)
            self.assertIn("Invalid cell location", str(context.exception), "Setting cell contents with an invalid reference should raise ValueError.")

    def test_bad_formula(self):
        # load and save_workbook() shouldn't reject cells representing bad 
        # formulas (i.e., bad ref) since that's the formula evaluator's job.
        self.workbook.new_sheet("Sheet1")
        bad_formula = "=1/+"

        self.workbook.set_cell_contents("Sheet1", "A1", bad_formula)
        cell_value = self.workbook.get_cell_value("Sheet1", "A1")

        self.assertIsInstance(cell_value, CellError, "A bad formula should result in a CellError.")
        self.assertEqual(cell_value.get_type(), CellErrorType.PARSE_ERROR, "The CellError should indicate the type of error caused by the bad formula.")
        json_file_path = os.path.join(self.script_directory, "wb_test_bad_formula.json")
        with open(json_file_path, 'w+') as newfile:
            # We explicitly check that save_workbook doesn't raise an exception.
            self.workbook.save_workbook(newfile)
        
        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)

            cell_value = loaded_workbook.get_cell_value("Sheet1", "A1") 
            self.assertIsInstance(cell_value, CellError, "A bad formula should result in a CellError.")
            self.assertEqual(cell_value.get_type(), CellErrorType.PARSE_ERROR, "The CellError should indicate the type of error caused by the bad formula.")
    
    def test_bad_reference(self):
        self.workbook.new_sheet("Sheet1")
        bad_reference = "=Sheet2!A3"

        self.workbook.set_cell_contents("Sheet1", "A1", bad_reference)
        cell_value = self.workbook.get_cell_value("Sheet1", "A1")

        self.assertIsInstance(cell_value, CellError, "A bad reference should result in a CellError.")
        self.assertEqual(cell_value.get_type(), CellErrorType.BAD_REFERENCE, "The CellError should indicate the type of error caused by the bad reference.")
        json_file_path = os.path.join(self.script_directory, "wb_test_bad_reference.json")    
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)
        
        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)

            cell_value = loaded_workbook.get_cell_value("Sheet1", "A1") 
            self.assertIsInstance(cell_value, CellError, "A bad reference should result in a CellError.")
            self.assertEqual(cell_value.get_type(), CellErrorType.BAD_REFERENCE, "The CellError should indicate the type of error caused by the bad reference.")

    def test_bad_reference_beyond_extent(self):
        self.workbook.new_sheet("Sheet1")
        bad_reference = "=Sheet1!ZZZZZ99999"

        self.workbook.set_cell_contents("Sheet1", "A1", bad_reference)
        cell_value = self.workbook.get_cell_value("Sheet1", "A1")

        self.assertIsInstance(cell_value, CellError, "A bad reference should result in a CellError.")
        self.assertEqual(cell_value.get_type(), CellErrorType.BAD_REFERENCE, "The CellError should indicate the type of error caused by the bad reference.")
        json_file_path = os.path.join(self.script_directory, "wb_test_bad_reference_beyond_extent.json") 
        with open(json_file_path, 'w+') as newfile:
            self.workbook.save_workbook(newfile)
        
        with open(json_file_path, 'r+') as newfile:
            loaded_workbook = Workbook.load_workbook(newfile)

            cell_value = loaded_workbook.get_cell_value("Sheet1", "A1") 
            self.assertIsInstance(cell_value, CellError, "A bad reference should result in a CellError.")
            self.assertEqual(cell_value.get_type(), CellErrorType.BAD_REFERENCE, "The CellError should indicate the type of error caused by the bad reference.")

    def test_load_single_quote_sheet_name_raises_exception(self):
        # Loading a WB with a single quote in sheetname SHOULD raise an error
        # because this is considered valid JSON with an invalid input (similar
        # to the invalid cell reference example in the spec).
        single_quote_sheet_json = {
            "sheets": [
                {
                    "name": "'Sheet1", # Sheet name cannot have single quotes
                    "cell-contents": {
                        "A1": "=2" 
                    }
                }
            ]
        }
        json_file_path = os.path.join(self.script_directory, "test_load_single_quote_sheet_name.json")  

        with open(json_file_path, 'w+') as newfile:
            json.dump(single_quote_sheet_json, newfile)
        
        with open(json_file_path, 'r') as newfile:
            with self.assertRaises(Exception) as context:
                self.workbook.load_workbook(newfile)

    def test_load_single_invalid_sheet_name_raises_exception(self):
        # This should cause an error because we create invalid input.
        single_quote_sheet_json = {
            "sheets": [
                {
                    "name": " Sheet1", # Sheet name cannot begin with whitespace
                    "cell-contents": {
                        "A1": "=2" 
                    }
                }
            ]
        }
        json_file_path = os.path.join(self.script_directory, "test_load_single_quote_sheet_name.json")  

        with open(json_file_path, 'w+') as newfile:
            json.dump(single_quote_sheet_json, newfile)
        
        with open(json_file_path, 'r') as newfile:
            with self.assertRaises(Exception) as context:
                self.workbook.load_workbook(newfile)

if __name__ == '__main__':
    unittest.main()
