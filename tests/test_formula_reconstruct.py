import context
import unittest

from sheets.lark_parser import LarkParser
from sheets.formula_constructer import FormulaReconstructor

class TestFormulaReconstructor(unittest.TestCase):
    def test_reconstruct_formula(self):
        reconstructor = FormulaReconstructor()
        _, tree, _ = LarkParser().parse_formula("=A1 + B2 * C3") 
        formula = reconstructor.reconstruct_formula(tree)
        self.assertEqual(formula.replace(" ", ""), "=A1+B2*C3")

    def test_reconstruct_formula_with_parentheses(self):
        reconstructor = FormulaReconstructor()
        _, tree, _ = LarkParser().parse_formula("=(A1 + B2) * C3") 
        formula = reconstructor.reconstruct_formula(tree)
        self.assertEqual(formula.replace(" ", ""), "=(A1+B2)*C3")

    def test_reconstruct_formula_with_subtraction(self):
        reconstructor = FormulaReconstructor()
        _, tree, _ = LarkParser().parse_formula("=A1 - B2")
        formula = reconstructor.reconstruct_formula(tree)
        self.assertEqual(formula.replace(" ", ""), "=A1-B2")

    def test_reconstruct_formula_with_multiplication(self):
        reconstructor = FormulaReconstructor()
        _, tree, _ = LarkParser().parse_formula("=A1 * B2")
        formula = reconstructor.reconstruct_formula(tree)
        self.assertEqual(formula.replace(" ", ""), "=A1*B2")

    def test_reconstruct_formula_with_division(self):
        reconstructor = FormulaReconstructor()
        _, tree, _ = LarkParser().parse_formula("=A1 / B2")
        formula = reconstructor.reconstruct_formula(tree)
        self.assertEqual(formula.replace(" ", ""), "=A1/B2")

    def test_reconstruct_formula_with_sheet_refs(self):
        reconstructor = FormulaReconstructor()
        formula = "=Sheet1!A1 + Sheet2!B2 * Sheet3!C3"
        _, tree, _ = LarkParser().parse_formula(formula)
        formula = reconstructor.reconstruct_formula(tree)
        self.assertEqual(formula.replace(" ", ""), formula.replace(" ", ""))

    def test_reconstruct_formula_with_nested_parentheses(self):
        reconstructor = FormulaReconstructor()
        _, tree, _ = LarkParser().parse_formula("=((A1 + B2) * C3) / D4")
        formula = reconstructor.reconstruct_formula(tree)
        self.assertEqual(formula, "=((A1+B2)*C3)/D4")

    def test_reconstruct_formula_with_sheet_refs_quotes(self):
        reconstructor = FormulaReconstructor()
        formula = "='Sheet1'!A1 + 'Sheet-2'!B2 * 'Sheet3'!C3"
        _, tree, _ = LarkParser().parse_formula(formula)
        formula = reconstructor.reconstruct_formula(tree)
        self.assertEqual(formula.replace(" ", ""), formula.replace(" ", ""))

if __name__ == '__main__':
    unittest.main()
