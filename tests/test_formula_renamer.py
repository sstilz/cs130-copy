import context
import unittest
from lark import Transformer
import re

from sheets.formula_renamer import FormulaRenamer
from sheets.lark_parser import LarkParser
from sheets.formula_constructer import FormulaReconstructor

class FormulaRenamerTests(unittest.TestCase):
    def setUp(self):
        self.p = LarkParser()
        self.r = FormulaReconstructor()

    def test_needs_quotes(self):
        renamer = FormulaRenamer('old', 'new')
        
        # Test sheetname that does not need quotes
        sheetname1 = 'valid_sheetname'
        self.assertFalse(renamer.needs_quotes(sheetname1))
        
        # Test sheetname that needs quotes
        sheetname2 = 'invalid sheetname'
        self.assertTrue(renamer.needs_quotes(sheetname2))
        
    def test__sheetname(self):
        renamer = FormulaRenamer('old', 'new')
        
        # Test sheetname that does not need quotes and does not match old_name
        children1 = 'valid_sheetname'
        self.assertEqual(renamer._sheetname(children1), 'valid_sheetname')
        
        # Test sheetname that needs quotes and does not match old_name
        children2 = 'invalid sheetname'
        self.assertEqual(renamer._sheetname(children2), '\'invalid sheetname\'')
        
        # Test sheetname that matches old_name
        children3 = 'old'
        self.assertEqual(renamer._sheetname(children3), 'new')
        
        # Test sheetname that needs quotes and matches old_name
        children4 = '\'invalid sheetname\''
        self.assertEqual(renamer._sheetname(children4), '\'invalid sheetname\'')
        
    def test_parse_formula(self):
        renamer = FormulaRenamer('old', 'new')
        
        formula = "= old!A1 + A2"
        _, tree, _ = self.p.parse_formula(formula)

        self.assertIsNotNone(tree)
        tree = renamer.transform(tree)
    
        self.assertEqual(self.r.reconstruct_formula(tree), "=new!A1+A2")
        
if __name__ == '__main__':
    unittest.main()
    