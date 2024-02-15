import context
import unittest
from sheets.worksheet import Worksheet
from sheets.cell import Cell


class TestWorksheet(unittest.TestCase):
    def setUp(self):
        self.worksheet = Worksheet("TestSheet")


    def test_serialize_sheet(self):
        self.worksheet.add_cell((0, 0), Cell("1"))
        self.worksheet.add_cell((1, 1), Cell("2"))
        self.worksheet.add_cell((2, 2), Cell("3"))

        serialized_data = self.worksheet.serialize()

        # Expected format of the serialized data
        expected_data = {
            'A1': '1',
            'B2': '2',
            'C3': '3'
        }

        self.assertEqual(serialized_data, expected_data)

    def test_serialize_empty_worksheet(self):
        serialized_data = self.worksheet.serialize()
        self.assertEqual(serialized_data, {}, "Serialized data of an empty worksheet should be an empty dictionary.")

    def test_serialize_format_and_type(self):
        self.worksheet.add_cell((0, 0), Cell("Test"))
        self.worksheet.add_cell((0, 1), Cell("1"))
        self.worksheet.add_cell((0, 2), Cell("=A1 + B2"))
        serialized_data = self.worksheet.serialize()
        self.assertIsInstance(serialized_data, dict, "Serialized data should be a dictionary.")

        expected_data = {
            'A1': 'Test',
            'B1': '1',
            'C1': '=A1 + B2'
        }

        self.assertEqual(serialized_data, expected_data)

if __name__ == '__main__':
    unittest.main()