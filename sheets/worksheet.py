import heapq
from sheets.workbook_utility import index_to_cell_location, is_valid_location

class Worksheet:
    def __init__(self, sheet_name):
        # max heap to track extent
        self.row_heap = list()
        self.col_heap = list()

        if not sheet_name.strip():
            raise ValueError("Sheet name cannot be empty or whitespace")
        self.sheet_name = sheet_name

        self.cell_map = dict()  # {location : Cell object}
       
    def get_extent(self):
        if len(self.row_heap) == 0 or len(self.col_heap) == 0:
            return (0, 0)
        
        # spec wants 1-indexed and col,row
        return ( -self.col_heap[0] + 1, -self.row_heap[0] + 1)
    
    def add_cell(self, cell_loc, cell_obj, is_implicit=False):
        """ Uses `cell_loc` for extent and adds `cell_obj` to `cell_map`. """
        if cell_loc in self.cell_map:
            raise ValueError("Adding cell that already exists")
        
        row, col = cell_loc
        if not is_valid_location(cell_loc):
            raise ValueError("Invalid cell location")
        
        self.cell_map[cell_loc] = cell_obj
        
        # If a cell is implicitly created it's value is None, so don't count
        # it towards the extent.
        if is_implicit:
            return
        heapq.heappush(self.row_heap, -row)
        heapq.heappush(self.col_heap, -col)
    
    def update_cell(self, cell_loc, cell_obj, contents):
        if cell_loc not in self.cell_map:
            raise ValueError("Cell does not exist")
        
        old_value = cell_obj.value
        row, col = cell_loc
        cell_obj.update(contents)   # this evaluates the cell's value

        if not old_value and cell_obj.value:
            heapq.heappush(self.row_heap, -row)
            heapq.heappush(self.col_heap, -col)
        
        elif old_value and not cell_obj.value:
            # Keep the cell in `cell_map` but remove it from the heaps since
            # empty cells are not considered part of the extent.
            self._remove_from_heap_and_heapify(cell_loc)

    def remove_cell(self, cell_loc):
        if cell_loc not in self.cell_map:
            raise ValueError("Cell does not exist")

        self._remove_from_heap_and_heapify(cell_loc)
        del self.cell_map[cell_loc]
    

    def get_cell(self, cell_loc):
        if cell_loc not in self.cell_map:
            # If the cell is unset, then we check if it's a valid cell location
            # If it isn't we raise a ValueError
            if not is_valid_location(cell_loc):
                raise ValueError("Invalid cell location")
            # If it is a valid cell location but unset, then we return None
            return None
        c = self.cell_map.get(cell_loc)
        return c
    
    def get_cell_exist(self, cell_loc):
        return cell_loc in self.cell_map
    
    def serialize(self):
        # converts info about sheet into a dict
        cells = {}
        for location in self.cell_map: # {location : Cell object}
            cell_obj = self.cell_map[location]
            cell_contents = cell_obj.content
            if cell_contents:
                fmt_location = index_to_cell_location(location[0], location[1])
                cells[fmt_location] = cell_contents
        return cells

    def _remove_from_heap_and_heapify(self, cell_loc):
        """ Removes cell from row and col heaps and heapifies them """
        row, col = cell_loc
        
        self.row_heap.remove(-row)
        self.col_heap.remove(-col)

        # Re-heapify after removal
        heapq.heapify(self.row_heap)  
        heapq.heapify(self.col_heap)

    def _pretty_print_cell_map(self):
        """ Returns a pretty printed string of the cell_map. """
        output = f"\nSheet name: {self.sheet_name}\nLocation\tContents\tValue\n"
        for location, cell_obj in self.cell_map.items():
            cell_location = index_to_cell_location(location[0], location[1])
            cell_content = cell_obj.content
            cell_value = cell_obj.value
            output += f"{cell_location}\t\t{cell_content}\t\t{cell_value}\n"
        return output