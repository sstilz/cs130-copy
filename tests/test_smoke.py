import context
import decimal
import sheets

# Make a new empty workbook
wb = sheets.Workbook()
(index, name) = wb.new_sheet()

wb.set_cell_contents(name, 'a1', '12')
wb.set_cell_contents(name, 'b1', '34')
wb.set_cell_contents(name, 'c1', '=a1+b1')

# value should be a decimal.Decimal('46')
value = wb.get_cell_value(name, 'c1')
assert value == decimal.Decimal('46')

wb.set_cell_contents(name, 'd3', '=nonexistent!b4')

# value should be a CellError with type BAD_REFERENCE
value = wb.get_cell_value(name, 'd3')

assert isinstance(value, sheets.CellError)
assert value.get_type() == sheets.CellErrorType.BAD_REFERENCE

# Cells can be set to error values as well
wb.set_cell_contents(name, 'e1', '#div/0!')
wb.set_cell_contents(name, 'e2', '=e1+5')
value = wb.get_cell_value(name, 'e2')
assert isinstance(value, sheets.CellError)
assert value.get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO