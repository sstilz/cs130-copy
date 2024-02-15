import re
import string
from typing import Callable, Iterable, List, Optional, TextIO, Tuple, Any

from lark import Token
from sheets.cell_error_type import CellErrorType, CellError
from sheets.worksheet import Worksheet
from sheets.graph import Graph
from sheets.cell import Cell
from sheets.formula_evaluator import FormulaEvaluator
from sheets.formula_renamer import FormulaRenamer
from sheets.formula_constructer import FormulaReconstructor
from sheets.graph import Graph
from sheets.workbook_utility import parse_cell_location_string, index_to_cell_location, is_valid_location
import decimal
import json

class Workbook:
    # A workbook containing zero or more named spreadsheets.
    #
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    def __init__(self):
        # Initialize a new empty workbook.
        self.worksheet_order = list()   # ordered list of WS objects
        self.sheet_to_tab = dict() # {uppercase sheet name : index of tab order}

        self.graph = Graph()
        self.notify_functions = []  # all registered functions (order matters)

    def num_sheets(self) -> int:
        return len(self.worksheet_order)

    def list_sheets(self) -> List[str]:
        return [s.sheet_name for s in self.worksheet_order]

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        if sheet_name is None:
            sheet_name = self._gen_new_sheetname()
        
        self._validate_sheet_name(sheet_name)
        
        self.sheet_to_tab[sheet_name.upper()] = len(self.worksheet_order)
        self.worksheet_order.append(Worksheet(sheet_name))
        self._evaluate()
        return (self.sheet_to_tab[sheet_name.upper()], sheet_name)

    def del_sheet(self, sheet_name: str) -> None:
        if sheet_name.upper() not in self.sheet_to_tab:
            raise KeyError(f"Sheet '{sheet_name}' not found")
        
        # delete all nodes in sheet from graph
        sheet_object = self._get_sheet(sheet_name)
        self.graph.clear_refs_criterion(lambda node_tuple: node_tuple[0].upper() == sheet_object.sheet_name.upper())

        # update sheet_to_tab to represent post-deletion indexes
        sheet_ind = self.sheet_to_tab[sheet_name.upper()]
        for i in range(sheet_ind + 1, len(self.worksheet_order)):
            self.sheet_to_tab[
                self.worksheet_order[i].sheet_name.upper()
            ] -= 1

        del self.worksheet_order[sheet_ind]
        del self.sheet_to_tab[sheet_name.upper()]

        self._evaluate()

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        sheet_object = self._get_sheet(sheet_name)
        return sheet_object.get_extent()

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str] = None) -> None:
        sheet_object = self._get_sheet(sheet_name)
        curr_sheet_name = sheet_object.sheet_name
        curr_loc = parse_cell_location_string(location) # curr_loc = (row, col)
        curr_cell_node = (curr_sheet_name.upper(), curr_loc)
        cell_exists = sheet_object.get_cell_exist(curr_loc)
        all_cells_changed = []

        if cell_exists:
            # Remove the existent cell's edges since they'll be recreated later.
            new_cell = sheet_object.get_cell(curr_loc)
            old_value = new_cell.value
            sheet_object.update_cell(curr_loc, new_cell, contents)
            self.graph.clear_refs(curr_cell_node)
            
            if not new_cell.is_formula():
                if new_cell.value != old_value:
                    all_cells_changed.append((sheet_name, location))         
        else:
            new_cell = Cell(contents, curr_loc)
            sheet_object.add_cell(curr_loc, new_cell)
            self.graph.add_node(curr_cell_node)

            if not new_cell.is_formula():
                all_cells_changed.append((sheet_name, location))

        self._update_cell_dependencies(new_cell, sheet_name, curr_cell_node)
        self._detect_cycle_and_propagate(curr_cell_node)
        self._notify(all_cells_changed)
        self._evaluate()
        
    def get_cell_contents(self, sheet_name: str, location: str) -> Optional[str]:
        cell = self._get_cell(sheet_name, location)
        if cell == None:
            return None
        return cell.content

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        cell = self._get_cell(sheet_name, location)
        if cell == None:
            return None
        return cell.value
    
    
    @staticmethod
    def load_workbook(fp: TextIO):
        # This is a static method (not an instance method) to load a workbook
        # from a text file or file-like object in JSON format, and return the
        # new Workbook instance.  Note that the _caller_ of this function is
        # expected to have opened the file; this function merely reads the file.
        #
        # If the contents of the input cannot be parsed by the Python json
        # module then a json.JSONDecodeError should be raised by the method.
        # (Just let the json module's exceptions propagate through.)  Similarly,
        # if an IO read error occurs (unlikely but possible), let any raised
        # exception propagate through.
        #
        # If any expected value in the input JSON is missing (e.g. a sheet
        # object doesn't have the "cell-contents" key), raise a KeyError with
        # a suitably descriptive message.
        #
        # If any expected value in the input JSON is not of the proper type
        # (e.g. an object instead of a list, or a number instead of a string),
        # raise a TypeError with a suitably descriptive message.

        #Reconstruct the SavedWorkbook/SavedSheet classes
        json_dct = json.load(fp) #Gets the JSON string as a dict

        if 'sheets' not in json_dct:
            raise KeyError("Malformed JSON: No sheets")
        
        #Reconstruct the actual Workbook/Worksheet classes
        wb = Workbook()
        try:
            sheet_names = [sheet_data["name"] for sheet_data in json_dct['sheets']]
        except KeyError as context:
            raise KeyError("Malformed JSON: Missing required fields name")
        for sn in sheet_names: 
            wb.new_sheet(sn)

        for sheet_data in json_dct['sheets']:
            if "cell-contents" not in sheet_data:
                raise KeyError("Malformed JSON: Missing required fields - cell-contents")
            sn = sheet_data["name"]
            cells = sheet_data["cell-contents"]
            
            for cell_location, cell_contents in cells.items():
                if not isinstance(cell_contents, str):
                    raise TypeError("Malformed JSON: Incorrect data type, cell content must be string")
                
                wb.set_cell_contents(sn, cell_location, cell_contents)
        return wb

    def save_workbook(self, fp: TextIO) -> None:
        # Instance method (not a static/class method) to save a workbook to a
        # text file or file-like object in JSON format.  Note that the _caller_
        # of this function is expected to have opened the file; this function
        # merely writes the file.
        #
        # If an IO write error occurs (unlikely but possible), let any raised
        # exception propagate through.
        
        # Serialize the workbook to a JSON string.
        sheets = list()
        for ws in self.worksheet_order:
            ws_dict = {
                "name": ws.sheet_name,
                "cell-contents": ws.serialize()
            }

            sheets.append(ws_dict)

        json.dump({"sheets": sheets}, fp, indent=4)
        return  

    def _gen_new_sheetname_based_on_original(self, original_name: str) -> str:
        i = 1
        new_name = f"{original_name}_{i}"
        while new_name.upper() in self.sheet_to_tab:
            i += 1
            new_name = f"{original_name}_{i}"
        return new_name

    def _standardize_changed_cells_for_notifs(self, changed_cells):
        for i, (sheet_name, location) in enumerate(changed_cells):
            original_sheetname = self._get_original_sheetname(sheet_name)
            if type(location) == tuple:
                row, col = location
                location = index_to_cell_location(row, col)
            changed_cells[i] = (original_sheetname, location.upper())
        return changed_cells
    
    def notify_cells_changed(self,
            notify_function: Callable[["Workbook", Iterable[Tuple[str, str]]], None]) -> None:
        # Request that all changes to cell values in the workbook are reported
        # to the specified notify_function.  The values passed to the notify
        # function are the workbook, and an iterable of 2-tuples of strings,
        # of the form ([sheet name], [cell location]).  The notify_function is
        # expected not to return any value; any return-value will be ignored.
        #
        # Multiple notification functions may be registered on the workbook;
        # functions will be called in the order that they are registered.
        #
        # A given notification function may be registered more than once; it
        # will receive each notification as many times as it was registered.
        #
        # If the notify_function raises an exception while handling a
        # notification, this will not affect workbook calculation updates or
        # calls to other notification functions.
        #
        # A notification function is expected to not mutate the workbook or
        # iterable that it is passed to it.  If a notification function violates
        # this requirement, the behavior is undefined.
        self.notify_functions.append(notify_function)
    
    def move_sheet(self, sheet_name: str, index: int) -> None:
        # Move the specified sheet to the specified index in the workbook's
        # ordered sequence of sheets. The index can range from 0 to
        # workbook.num_sheets() - 1. The index is interpreted as if the
        # specified sheet were removed from the list of sheets, and then
        # re-inserted at the specified index.
        #
        # The sheet name match is case-insensitive; the text must match, but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        #
        # If the index is outside the valid range, an IndexError is raised.
        if not self._get_sheet_exists(sheet_name):
            raise KeyError(f"Sheet name '{sheet_name}' doesn't exist in workbook")
        if index < 0 or index >= self.num_sheets():
            raise IndexError(f"Index '{index}' not in bounds of workbook")

        original_index = self.sheet_to_tab[sheet_name.upper()]
        sheet = self.worksheet_order.pop(original_index)
        self.worksheet_order.insert(index, sheet)

        for i, sheet in enumerate(self.worksheet_order):
            self.sheet_to_tab[sheet.sheet_name.upper()] = i

    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        # Make a copy of the specified sheet, storing the copy at the end of the
        # workbook's sequence of sheets.  The copy's name is generated by
        # appending "_1", "_2", ... to the original sheet's name (preserving the
        # original sheet name's case), incrementing the number until a unique
        # name is found.  As usual, "uniqueness" is determined in a
        # case-insensitive manner.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # The copy should be added to the end of the sequence of sheets in the
        # workbook.  Like new_sheet(), this function returns a tuple with two
        # elements:  (0-based index of copy in workbook, copy sheet name).  This
        # allows the function to report the new sheet's name and index in the
        # sequence of sheets.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        if not self._get_sheet_exists(sheet_name):
            raise KeyError(f"Sheet '{sheet_name}' not found")
        original_sheet_obj = self._get_sheet(sheet_name)
        original_name = original_sheet_obj.sheet_name

        # 1. Create a new sheet with a unique name
        copied_sheet_name = self._gen_new_sheetname_based_on_original(original_name)
        idx, _ = self.new_sheet(copied_sheet_name)

        # 2. Iterate through old sheet and set new cells and their contents        
        for location_tup in original_sheet_obj.cell_map: #not using k,v iteration because weird behavior with the k already being a tuple, which unpacks weirdly during iteration.
            cell_obj = original_sheet_obj.cell_map[location_tup]
            self.set_cell_contents(copied_sheet_name, index_to_cell_location(location_tup[0], location_tup[1]), cell_obj.content)
        return idx, copied_sheet_name
    
    def rename_sheet(self, sheet_name: str, new_sheet_name: str) -> None:
        # Rename the specified sheet to the new sheet name.  Additionally, all
        # cell formulas that referenced the original sheet name are updated to
        # reference the new sheet name (using the same case as the new sheet
        # name, and single-quotes iff [if and only if] necessary).
        #
        # The sheet_name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # As with new_sheet(), the case of the new_sheet_name is preserved by
        # the workbook.
        #
        # If the sheet_name is not found, a KeyError is raised.
        #
        # If the new_sheet_name is an empty string or is otherwise invalid, a
        # ValueError is raised.
        renamed_sheet_obj = self._rename_sheet(sheet_name, new_sheet_name)
        
        # update cell loc tups
        for loc, cell in renamed_sheet_obj.cell_map.items():
            # rename cell
            cell.loc = (new_sheet_name, loc)

        # Change graph
        for cell_sheet_name, cell_loc in self.graph.get_all_nodes():
            if cell_sheet_name.upper() == sheet_name.upper():
                curr_cell_node = (cell_sheet_name.upper(), cell_loc)
                self.graph.rename_cell(curr_cell_node, (new_sheet_name.upper(), cell_loc))
        
        # Change formulas
        renamer = FormulaRenamer(sheet_name, new_sheet_name)
        for sheet_obj in self.worksheet_order:
            for loc, cell in sheet_obj.cell_map.items():
                if cell.is_formula() and self._sheet_name_in_cell_ref(sheet_name, cell.get_refs()):
                    new_tree = renamer.transform(cell.tree)
                    new_content = FormulaReconstructor().reconstruct_formula(new_tree)
                    cell.update(new_content)
        
        # Update cell dependencies if needed and detect cycles and propagate
        for cell_sheet_name, cell_loc in self.graph.get_all_nodes():
            # if cell belongs to the renamed sheet
            if cell_sheet_name.upper() == new_sheet_name.upper():   
                curr_cell_node = (cell_sheet_name.upper(), cell_loc)
                cell = self._get_cell_loc_tup(cell_sheet_name, cell_loc)
                # If the cell is not part of the renamed sheet, then it was implicitly referenced by another cell
                # And now that it's part of the renamed sheet, it needs to be updated in the sheet object.
                if not cell:
                    sheet_object = self._get_sheet(new_sheet_name.upper())
                    sheet_object.add_cell(cell_loc, Cell(None, cell_loc), is_implicit=True)
                    cell = self._get_cell_loc_tup(cell_sheet_name, cell_loc)

                # Update cell dependencies
                self._update_cell_dependencies(cell, cell_sheet_name, curr_cell_node)
                self._detect_cycle_and_propagate(curr_cell_node)
        self._evaluate()


    def _update_cell_dependencies(self, new_cell, sheet_name, curr_cell_node):
        children = new_cell.refs    # empty if not a formula
        for child_loc in children:
            (child_sheet, child_loc_str) = child_loc if len(child_loc) == 2 else (sheet_name, child_loc[0])
            child_loc = parse_cell_location_string(child_loc_str)
            self.graph.add_edge((child_sheet.upper(), child_loc), curr_cell_node)

    def _sheet_name_in_cell_ref(self, sheet_name, cell_ref):
        for ref in cell_ref:
            (cell_sheet_name, loc_str) = ref if len(ref) == 2 else (sheet_name, ref[0])
            if sheet_name.upper() == cell_sheet_name.upper() or f"'{sheet_name.upper()}'" == cell_sheet_name.upper():
                return True
        return False
            
    def _evaluate(self):
        topo_sort = self.graph.get_topo_sort()
        all_cells_changed = []
        for (cell_sheet_name, cell_loc) in topo_sort:
            # If referring to a cell in a sheet that DNE, skip evaluation
            # because that invalid cell object wouldn't have even been created.
            if not self._get_sheet_exists(cell_sheet_name) or not is_valid_location(cell_loc):
                continue
            try:
                c = self._get_cell_loc_tup(cell_sheet_name, cell_loc)
                if c == None:
                    c = Cell(None, cell_loc)
                    self._get_sheet(cell_sheet_name).add_cell(cell_loc, c, is_implicit=True)
                    # Should cell notify as you had to create a cell for an implicitly referenced cell that is unset
            except ValueError:
                return
            old_value = c.value
            e = FormulaEvaluator(self, cell_sheet_name)
            if c.tree:
                if self._refers_to_self(c, cell_sheet_name, cell_loc):
                    c.value = CellError(CellErrorType.CIRCULAR_REFERENCE, "Self circular reference")
                elif self._refers_to_single_none_cell(c, cell_sheet_name):
                    c.value = decimal.Decimal('0')
                else:
                    v = e.evaluate(c.tree)
                    if CellError.is_error_string(v):
                        error_type = CellError.get_error_type_from_string(v)
                        c.value = CellError(error_type, CellError.get_detail_from_error_type(error_type))
                    else:
                        c.value = v
            if c.value != old_value and not self._is_same_error(c.value, old_value):
                all_cells_changed.append((cell_sheet_name, cell_loc))

        self._notify(all_cells_changed)
    
    def _is_same_error(self, e1, e2):
        if isinstance(e1, CellError) and isinstance(e2, CellError):
            return e1.get_type() == e2.get_type()
        return False
    
    def _refers_to_single_cell(self, c):   
        # Ex, this cell's reference is 'A1'
        if c.content == None:
            return False
        if (len(c.tree.children) == 1 and isinstance(c.tree.children[0], Token) and len(c.get_refs()) == 1):
            return True
    
        # Ex, this cell's reference is `Sheet1!A1`
        if (len(c.tree.children) == 2 and isinstance(c.tree.children[0], Token) and isinstance(c.tree.children[1], Token) and len(c.get_refs()) == 1):
            return True
        return False

    def _refers_to_self(self, c, cell_sheet_name, cell_loc):
        if not self._refers_to_single_cell(c):
            return False
        ref_sheet_name = c.get_refs()[0][0] if len(c.get_refs()[0]) == 2 else cell_sheet_name
        ref_loc_str = c.get_refs()[0][1] if len(c.get_refs()[0]) == 2 else c.get_refs()[0][0]
        return (cell_sheet_name, cell_loc) == (ref_sheet_name, parse_cell_location_string(ref_loc_str))
    
    def _refers_to_single_none_cell(self, c, sheetname):
        if not self._refers_to_single_cell(c):
            return False
        ref_sheetname = c.get_refs()[0][0] if len(c.get_refs()[0]) == 2 else sheetname
        location = c.get_refs()[0][1] if len(c.get_refs()[0]) == 2 else c.get_refs()[0][0]
        location = parse_cell_location_string(location)
        
        try:
            sheet_obj = self._get_sheet(ref_sheetname)
            cell_obj = sheet_obj.get_cell(location)
            if cell_obj.content == None:
                return True
            return False
        except:
            # Cell ref doesn't even exist yet, so can't possibly be a None cell.
            return False
        
    def _is_immediately_evaluatable(self, c):
        # either single number, string, or cell
        is_single_node = len(c.tree.children) == 1
        return len(c.tree.children) == 1 and len(c.refs) == 1

    def _detect_cycle_and_propagate(self, curr_cell_node):
        # Cycle detection with Tarjan's algorithm
        component = self.graph.get_component(curr_cell_node)

        assert len(component) > 0
        if len(component) != 1:
            # Keep track of cells that have been changed so that we can notify them.
            all_cells_changed = []
            # Cycle detected, so set every cell in the cycle to a CIRCULAR_REFERENCE error.
            for cell_sheet_name, cell_loc in component:
                c = self._get_cell_loc_tup(cell_sheet_name, cell_loc)
                if (c.value == None or not isinstance(c.value, CellError) or c.value.get_type() != CellErrorType.CIRCULAR_REFERENCE):
                    all_cells_changed.append((cell_sheet_name, cell_loc))
                c.value = CellError(CellErrorType.CIRCULAR_REFERENCE, "Circular reference detected")
            self._notify(all_cells_changed)
            return True
        return False

    def _notify(self, changed_cells: Iterable[Tuple[str, str]]) -> None:
        # ("SHEET1", "A1") -> ("Sheet1", (0, 0))
        changed_cells = self._standardize_changed_cells_for_notifs(changed_cells)
    
        for notify_function in self.notify_functions:
            try:
                notify_function(self, changed_cells)
            except Exception as e:
                pass

    # Given a cell's sheetname (case insensitive), retrieve the actual sheetname
    # with the case preserved.
    def _get_original_sheetname(self, name):
        sheet_object = self._get_sheet(name)
        return sheet_object.sheet_name

    def _get_sheet(self, sheet_name):
        if not self._get_sheet_exists(sheet_name):
            raise KeyError(f"Sheet '{sheet_name}' not found.")
        
        sheet_idx = self.sheet_to_tab[sheet_name.upper()]
        return self.worksheet_order[sheet_idx]
    
    def _validate_sheet_name(self, sheet_name):
        # Technically shouldn't get to here because of the above check.
        if len(sheet_name) == 0:
            raise ValueError("Sheet name cannot be empty")
        
        if sheet_name.strip() != sheet_name:
            raise ValueError("Sheet name cannot start or end with whitespace")
        
        # Quote marks (single or double) are not allowed.
        if '"' in sheet_name or "'" in sheet_name:
            raise ValueError("Sheet name cannot contain quote marks")
        
        # The only allowed characters are alphanumeric, space, and the following
        # punctuation: .?!,:;!@#$%^&*()-_
        allowed_chars = string.ascii_letters + string.digits + ' .?!,:;!@#$%^&*()-_'
        if any(char not in allowed_chars for char in sheet_name):
            raise ValueError("Sheet name contains invalid characters")
        
        if sheet_name.upper() in self.sheet_to_tab:
            raise ValueError("Sheet name must be unique")
        

    def _rename_sheet(self, sheet_name, new_sheet_name):
        self._validate_sheet_name(new_sheet_name)
        
        sheet_idx = self.sheet_to_tab[sheet_name.upper()]
        sheet_obj = self._get_sheet(sheet_name)

        sheet_obj.sheet_name = new_sheet_name
        del self.sheet_to_tab[sheet_name.upper()]
        self.sheet_to_tab[new_sheet_name.upper()] = sheet_idx 

        return sheet_obj 
    
    def _get_sheet_exists(self, sheet_name):
        return sheet_name.upper() in self.sheet_to_tab

    def _get_cell(self, sheet_name, location_str):
        sheet_object = self._get_sheet(sheet_name)
        location_tuple = parse_cell_location_string(location_str)
        return sheet_object.get_cell(location_tuple)
    
    def _get_cell_loc_tup(self, sheet_name, location_tuple):
        sheet_object = self._get_sheet(sheet_name)
        c = sheet_object.get_cell(location_tuple)
        return c

    def _get_notify_functions(self):
        return self.notify_functions
        
    def _gen_new_sheetname(self):
        i = 1
        sheet_name = f"Sheet{i}"
        while sheet_name.upper() in self.sheet_to_tab:
            i += 1
            sheet_name = f"Sheet{i}"
        return sheet_name

    def _pretty_print_all_sheets_in_order(self):
        """ 
        Returns a string representation of all sheets in the workbook,
        including their cell map. This will be helpful for debugging.
        """
        output = ""
        for i in range(len(self.worksheet_order)):
            sheet_obj = self.worksheet_order[i]
            output += sheet_obj._pretty_print_cell_map()
        return output
