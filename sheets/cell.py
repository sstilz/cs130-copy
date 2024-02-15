from decimal import Decimal
from .lark_parser import LarkParser
from .value_type import ValueType
from .cell_error_type import CellErrorType, CellError

class Cell:
    def __init__(self, content, loc_tup=None):
        self.content = None
        self.type = None
        self.value = None
        self.tree = None
        self.refs = []

        self.loc = loc_tup

        self.update(content)

    def update(self, content):
        self.content = content
        self.type = None  # Reset type to None when updating
        self.tree = None
        self.refs = []

        self._format_content()
        self._detect_type()
        self._evaluate_content()

    def _format_content(self):
        if self.content:
            self.content = self.content.strip()
            
            if self.content == "'" or len(self.content) == 0:
                self.content = None
                self.value = None
        else:
            self.content = None
            self.value = None
        

    def _strip_trailing_zeros(self, decimal_str):
        stripped_decimal = str(float(decimal_str)).rstrip('0').rstrip('.')
        return stripped_decimal

    def _detect_type(self):
        if self.content is not None:
            if self.content.startswith("'"):
                self.type = ValueType.STRING
            elif self.content.startswith('='):
                self.type = ValueType.FORMULA
            else:
                self.type = self._get_literal_type()

    def _get_literal_type(self):
        try:
            dec = Decimal(self._strip_trailing_zeros(self.content))
            if (dec == Decimal('Infinity') or dec == Decimal('-Infinity') or
                    dec.is_nan()):
                return ValueType.STRING
            return ValueType.NUMBER
        except ValueError:
            errors = ["#ERROR!", "#CIRCREF!", "#REF!", "#NAME?", "#VALUE!", "#DIV/0!"]
            if self.content.upper() in errors:
                return ValueType.ERROR
            return ValueType.STRING

    def _evaluate_content(self):
        if self.type == ValueType.STRING:
            if self.content.startswith("'"):
                self.value = self.content[1:]
            else:
                self.value = self.content
        elif self.type == ValueType.FORMULA:
            parser = LarkParser()
            ref, tree, error = parser.parse_formula(self.content)
            if (error):
                self.value = CellError(CellErrorType.PARSE_ERROR, error)
                self.type = ValueType.ERROR
            else:
                self.tree = tree
                self.refs = ref
        elif self.type == ValueType.NUMBER:
            self.value = Decimal(self._strip_trailing_zeros(self.content))
        elif self.type == ValueType.ERROR:
            self.content = self.content.upper()
            self.value = CellError(self._get_error_type(), self.content)

    def _get_error_type(self):
        if self.content == "#ERROR!":
            return CellErrorType.PARSE_ERROR
        elif self.content == "#CIRCREF!":
            return CellErrorType.CIRCULAR_REFERENCE
        elif self.content == "#REF!":
            return CellErrorType.BAD_REFERENCE
        elif self.content == "#NAME?":
            return CellErrorType.BAD_NAME
        elif self.content == "#VALUE!":
            return CellErrorType.TYPE_ERROR
        elif self.content == "#DIV/0!":
            return CellErrorType.DIVIDE_BY_ZERO

    def is_formula(self):
        return self.type == ValueType.FORMULA
    
    def is_number(self):
        return self.type == ValueType.NUMBER
    
    def is_string(self):
        return self.type == ValueType.STRING
    
    def get_refs(self):
        return self.refs
    
    def get_type(self):
        return self.type

    def get_value(self):
        return self.value
    
    def get_content(self):
        return self.content
    
    def get_tree(self):
        return self.tree
