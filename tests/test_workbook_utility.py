import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import unittest
from sheets.workbook_utility import index_to_cell_location, parse_cell_location_string

class TestParseCellLocation(unittest.TestCase):
    def test_valid_cell_locations(self):
        test_cases = {
            "a1": (0, 0),   # check case sensitivity
            "A1": (0, 0),
            "B2": (1, 1),
            "AA27": (26, 26),
            "ZZZ1000": (999, 18277),
            "ABC123": (122, 730),
            "XFD1048576": (1048575, 16383)
        }

        for cell_location, expected_result in test_cases.items():
            with self.subTest(cell_location=cell_location):
                result = parse_cell_location_string(cell_location)
                self.assertEqual(result, expected_result)

    def test_invalid_cell_locations(self):
        invalid_locations = [
            "A"
            "1A",  # Invalid format
            "ABCD",  # Missing row
            "XYZ0",  # Invalid row number
            "A0",  # Zero row not allowed
            "AA-1",  # Negative row not allowed
            "AAA1.5",  # Decimal row not allowed
            "A1A",  # Extra characters
            "AA27A",  # Extra characters
            "AA27A2A",  # Extra characters
        ]

        for cell_location in invalid_locations:
            with self.subTest(cell_location=cell_location):
                with self.assertRaises(ValueError):
                    parse_cell_location_string(cell_location)

    def test_invalid_formats_with_special_characters(self):
        invalid_locations = [
            "A#1",       # Special character in column part
            "B$2",   
            "C%3",       
            "D4^",       # Special character in row part
            "E5&",       
        ]
        for cell_location in invalid_locations:
            with self.subTest(cell_location=cell_location):
                with self.assertRaises(ValueError):
                    parse_cell_location_string(cell_location)

    def test_excessively_large_numbers(self):
        test_cases = {
            "99999": ValueError,  # ZZZZ9999 is our max
            "ZZZZZ": ValueError,  
        }
        for cell_location, expected_exception in test_cases.items():
            with self.subTest(cell_location=cell_location):
                with self.assertRaises(expected_exception):
                    parse_cell_location_string(cell_location)

    def test_lowercase_cell_locations(self):
        test_cases = {
            "a1": (0, 0),
            "b2": (1, 1),
            "aa27": (26, 26),
        }
        for cell_location, expected_result in test_cases.items():
            with self.subTest(cell_location=cell_location):
                result = parse_cell_location_string(cell_location)
                self.assertEqual(result, expected_result)
    
class TestGenerateCellLocation(unittest.TestCase):
    def test_valid_cell_locations(self):
        test_cases = {
            (0, 0): "A1",
            (1, 1): "B2",
            (26, 26): "AA27",
            (999, 18277): "ZZZ1000",
            (122, 730): "ABC123",
            (1048575, 16383): "XFD1048576"
        }

        for coordinates, expected_result in test_cases.items():
            with self.subTest(coordinates=coordinates):
                result = index_to_cell_location(*coordinates)
                self.assertEqual(result, expected_result)

    def test_invalid_coordinates(self):
        invalid_coordinates = [
            (-1, 0),  # Invalid row
            (0, -1),  # Invalid column
            (-1, -1),  # Invalid row and column
        ]

        for coordinates in invalid_coordinates:
            with self.subTest(coordinates=coordinates):
                with self.assertRaises(ValueError):
                    index_to_cell_location(*coordinates)

    def test_index_to_cell_location_first_column(self):
        result = index_to_cell_location(0, 0)
        self.assertEqual(result, 'A1')

    def test_index_to_cell_location_single_letter_column(self):
        result = index_to_cell_location(1, 0)
        self.assertEqual(result, 'A2')

    def test_index_to_cell_location_double_letter_column(self):
        result = index_to_cell_location(4, 26)
        self.assertEqual(result, 'AA5')

    def test_index_to_cell_location_large_column_index(self):
        result = index_to_cell_location(9, 702)
        self.assertEqual(result, 'AAA10')

    def test_index_to_cell_location_negative_column_index(self):
        with self.assertRaises(ValueError):
            index_to_cell_location(2, -1)
                    
if __name__ == '__main__':
    unittest.main()