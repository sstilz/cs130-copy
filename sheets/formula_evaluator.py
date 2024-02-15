import lark
import decimal
import sheets
from sheets.cell_error_type import CellErrorType, CellError
import lark
import decimal

class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self, workbook, sheet):
        self.workbook = workbook
        self.sheet = sheet

    def evaluate(self, parse_tree):
        values = self.visit(parse_tree)
        if self._is_error(values):
            return self._convert_error(values)
        return values
        
    def find_first_error(self, values):
        if self._is_circ_ref(values):
            return "#CIRCREF!"
        for value in values:
            if CellError.is_error_string(value):
                return value
        return None
    
    def translate_cell_value_to_decimal(self, value):
        if CellError.is_error_string(value):
            return value
        if self._is_error(value):
            return value[0]
        if value is None:
            return 0
        try: 
            stripped_decimal = str(float(value)).rstrip('0').rstrip('.')
            return decimal.Decimal(stripped_decimal)
        except:
            return "#VALUE!"
        
    def translate_cell_value_to_string(self, value):
        if CellError.is_error_string(value):
            return value 
        if self._is_error(value):
            return value[0]   
        if value is None:
            return ""
        try: 
            return str(value)
        except:
            return "#VALUE!"
        
    def handle_empty_cell_decimal(self, values):
        if(len(values) == 1):
            # Could return as "left" instead, but just choosing right to be consistent
            return 0, self.translate_cell_value_to_decimal(values[0])
        if (values[0] == lark.Token('ADD_OP', '-') or values[0] == lark.Token('ADD_OP', '+')):
            return 0, self.translate_cell_value_to_decimal(values[1])
        left = self.translate_cell_value_to_decimal(values[0])
        right = self.translate_cell_value_to_decimal(values[2])
        return left, right

    def handle_empty_cell_string(self, values):
        left = self.translate_cell_value_to_string(values[0])
        right = self.translate_cell_value_to_string(values[1])
        return left, right

    def add_expr(self, tree):
        values = self.visit_children(tree)
        error = self.find_first_error(values)
        if error:
            return error
        left, right = self.handle_empty_cell_decimal(values)      
        if CellError.is_error_string(left):
            return left
        if CellError.is_error_string(right):
            return right
        return left + right if values[1] == '+' else left - right

    def mul_expr(self, tree):
        values = self.visit_children(tree)
        error = self.find_first_error(values)
        if error:
            return error
        left, right = self.handle_empty_cell_decimal(values)
        if CellError.is_error_string(left):
            return left
        if CellError.is_error_string(right):
            return right
        return left * right if tree.children[1] == '*' else left / right if right != 0 else "#DIV/0!"

    def unary_op(self, tree):
        values = self.visit_children(tree)
        error = self.find_first_error(values)
        if error:
            return error
        left, right = self.handle_empty_cell_decimal(values)
        if CellError.is_error_string(left):
            return left
        if CellError.is_error_string(right):
            return right
        return +right if values[0] == '+' else -right

    def concat_expr(self, tree):
        values = self.visit_children(tree)
        error = self.find_first_error(values)
        if error:
            return error
        left, right = self.handle_empty_cell_string(values)
        left, right = '' if left is None else left, '' if right is None else right
        return str(left) + str(right)
    
    def cell(self, tree):
        values = self.visit_children(tree)
        sheet = self.sheet if len(values) == 1 else values[0]
        
        sheet = self._strip_outer_single_quotes(sheet)

        try: 
            # If cell references another cell, then try to get the value of that cell
            value = self.workbook.get_cell_value(sheet, values[-1])
            
            # If the cell is an error, we want the string representation of the erro
            if isinstance(value, CellError):
               value = CellError.get_string_from_error_type(value.get_type())
        except (KeyError, ValueError):
            value = "#REF!"   
        return value 

    def parens(self, tree):
        values = self.visit_children(tree)
        left, right = self.handle_empty_cell_decimal(values)
        if CellError.is_error_string(right):
            return right
        return right
    
    def number(self, tree):
        stripped_decimal = str(float(tree.children[0])).rstrip('0').rstrip('.')
        return decimal.Decimal(stripped_decimal)
    
    def string(self, tree):
        return tree.children[0][1:-1]
    
    def base(self, tree):
       return self.visit_children(tree)[0]
    
    def _convert_error(self, values):
        if (isinstance(values, list) and CellError.is_error_string(values[0])):
            error_str = values[0]
        else:
            error_str = values
        error_type = CellError.get_error_type_from_string(error_str)
        return CellError(error_type, CellError.get_detail_from_error_type(error_type))
    
    def _is_error(self, values):
        return (isinstance(values, list) and CellError.is_error_string(values[0])) or CellError.is_error_string(values)
    
    def _strip_outer_single_quotes(self, string):
        if string.startswith("'") and string.endswith("'"):
            return string[1:-1]
        return string
    
    def _is_circ_ref(self, values):
        if ("#CIRCREF!" in values):
            return True
        return False
