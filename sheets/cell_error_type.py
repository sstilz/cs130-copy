from typing import Optional
import enum

cell_error_strings = [
    "#ERROR!", "#CIRCREF!", "#REF!", "#NAME?", "#VALUE!", "#DIV/0!"
]

cell_error_details = [
    "Parsing error", "Circular reference detected", "Invalid reference",
     "Invalid function name", "Type mismatch", "Divide by zero"
]


class CellErrorType(enum.Enum):
    '''
    This enum specifies the kinds of errors that spreadsheet cells can hold.
    '''

    # A formula doesn't parse successfully ("#ERROR!")
    PARSE_ERROR = 1

    # A cell is part of a circular reference ("#CIRCREF!")
    CIRCULAR_REFERENCE = 2

    # A cell-reference is invalid in some way ("#REF!")
    BAD_REFERENCE = 3

    # Unrecognized function name ("#NAME?")
    BAD_NAME = 4

    # A value of the wrong type was encountered during evaluation ("#VALUE!")
    TYPE_ERROR = 5

    # A divide-by-zero was encountered during evaluation ("#DIV/0!")
    DIVIDE_BY_ZERO = 6



class CellError:
    '''
    This class represents an error value from user input, cell parsing, or
    evaluation.
    '''

    def __init__(self, error_type: CellErrorType, detail: str,
                 exception: Optional[Exception] = None):
        self._error_type = error_type
        self._detail = detail
        self._exception = exception

    def get_type(self) -> CellErrorType:
        ''' The category of the cell error. '''
        return self._error_type

    def get_detail(self) -> str:
        ''' More detail about the cell error. '''
        return self._detail

    def get_exception(self) -> Optional[Exception]:
        '''
        If the cell error was generated from a raised exception, this is the
        exception that was raised.  Otherwise this will be None.
        '''
        return self._exception

    def __str__(self) -> str:
        return f'ERROR[{self._error_type}, "{self._detail}"]'

    def __repr__(self) -> str:
        return self.__str__()
    
    def get_error_type_string(self) -> str:
        error_type = self.get_type()
        return CellError.get_string_from_error_type(error_type)
    
    def get_string_from_error_type(error_type) -> str:
        return cell_error_strings[error_type.value - 1]
    
    def is_error_string(value) -> bool:
        return value in cell_error_strings

    def get_error_type_from_string(value) -> str:
        if value in cell_error_strings:
            return CellErrorType(cell_error_strings.index(value) + 1)
        return None
    
    def get_detail_from_error_type(error_type) -> str:
        return cell_error_details[error_type.value - 1]
