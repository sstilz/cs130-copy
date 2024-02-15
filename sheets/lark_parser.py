import lark
import os

class LarkParser():
    def __init__(self): 
        self.parser = lark.Lark.open("formulas.lark", start='formula', rel_to=__file__)
        self.cell_ref_finder = CellRefFinder()

    def parse_formula(self, formula):
        try:
            tree = self.parser.parse(formula)
            self.cell_ref_finder.visit(tree) # Gather cell references
            
            return self.cell_ref_finder.refs, tree, None
        except Exception as e:
            # Handle parsing errors here, set the cell's contents to the error value
            return None, None, "#ERROR!"  # Return no cell references and the parsing error

class CellRefFinder(lark.Visitor):
    def __init__(self):
        self.refs = []

    def cell(self, tree):
        if len(tree.children) == 1:
            self.refs.append((str(tree.children[0]).upper(), ))
        else:
            assert(len(tree.children) == 2)
            self.refs.append((tree.children[0].upper(), tree.children[1].upper()))
