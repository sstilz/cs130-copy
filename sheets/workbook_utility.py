import re

def parse_cell_location_string(cell_location):
    # Use regular expressions to extract letters and numbers from the cell location
    match = re.match(r'^([A-Z]+)(\d+)$', cell_location, re.IGNORECASE)

    if match:
        # Extract groups from the match
        column_letters, row_number = match.groups()
        if int(row_number) < 1:
            raise ValueError("Invalid row number")
        
        # Convert column letters to a numeric value
        column = 0
        for char in column_letters:
            column = column * 26 + (ord(char.upper()) - ord('A')) + 1

        # Zero index the column and row
        column -= 1
        row = int(row_number) - 1

        return row, column
    else:
        raise ValueError("Invalid cell location format")
    
def index_to_cell_location(row, col):
    if row < 0 or col < 0:
        raise ValueError("Row and column indices must be non-negative")

    result = ""
    while col >= 0:
        result = chr(col % 26 + ord('A')) + result
        col //= 26
        col -= 1

    return result + str(row + 1)

def is_valid_location(cell_loc):
    """ Returns if a cell's location is within A1 and ZZZZ9999, inclusive. """
    row, col = cell_loc
    # Assuming ZZZZ maps to 475253
    return 0 <= row < 9999 and 0 <= col <= 475253 