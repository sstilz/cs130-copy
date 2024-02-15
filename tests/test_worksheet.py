import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import unittest
from sheets.worksheet import Worksheet
from sheets.cell import Cell

class TestWorksheet(unittest.TestCase):
    def setUp(self):
        self.worksheet = Worksheet("TestSheet")

    def test_add_cell_in_already_occupied_location(self):
        self.worksheet.add_cell((0, 0), None)
        self.assertEqual(self.worksheet.get_extent(), (1, 1))

        # Adding a cell in a location that already exists should raise 
        # a ValueError.
        with self.assertRaises(ValueError):
            self.worksheet.add_cell((0, 0), None)

    def test_remove_cell(self):
        self.worksheet.add_cell((0, 0), None)
        self.assertEqual(self.worksheet.get_extent(), (1, 1))

        self.worksheet.remove_cell((0, 0))
        self.assertEqual(self.worksheet.get_extent(), (0, 0))

        # Removing a non-existent cell should raise ValueError
        with self.assertRaises(ValueError):
            self.worksheet.remove_cell((1, 1))

    def test_get_extent(self):
        # Initially, the extent should be (0, 0)
        self.assertEqual(self.worksheet.get_extent(), (0, 0))

        self.worksheet.add_cell((2, 3), None)
        self.assertEqual(self.worksheet.get_extent(), (4, 3))

        self.worksheet.add_cell((1, 5), None)
        self.assertEqual(self.worksheet.get_extent(), (6, 3))

    def test_remove_get_extent(self):
        self.worksheet.add_cell((5, 5), None)
        self.worksheet.add_cell((10, 10), None)
        self.worksheet.remove_cell((10, 10))
        self.assertEqual(self.worksheet.get_extent(), (6, 6))

    def tearDown(self):
        del self.worksheet  # Clean up the worksheet object

    def test_worksheet_initialization(self):
        self.assertEqual(self.worksheet.sheet_name, "TestSheet")

    def test_get_cell_with_nonexistent_cell_raises_error(self):
        # removed raise ValueError -> unset cell should just be None, as long as it's a valid location
        self.assertIsNone(self.worksheet.get_cell((1, 1)))

    def test_cell_existence(self):
        self.worksheet.add_cell((2, 2), None)
        self.assertTrue(self.worksheet.get_cell_exist((2, 2)))
        self.assertFalse(self.worksheet.get_cell_exist((3, 3)))

    def test_invalid_cell_locations(self):
        with self.assertRaises(ValueError):
            self.worksheet.add_cell((-1, -1), None)

        with self.assertRaises(ValueError):
            self.worksheet.add_cell((10000, 10000), None)  # Beyond max limit

    def test_invalid_sheet_names(self):
        with self.assertRaises(ValueError):
            Worksheet("")

        with self.assertRaises(ValueError):
            Worksheet("    ")
        

if __name__ == '__main__':
    unittest.main()