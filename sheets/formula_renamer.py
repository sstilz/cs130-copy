from lark import Transformer, Token, Tree
import re

class FormulaRenamer(Transformer):
    def __init__(self, old_name, new_name):
        super().__init__()
        self.old_name = old_name
        self.new_name = new_name

    def needs_quotes(self, sheetname):
        sheet_name_pattern = r'^[A-Za-z_][A-Za-z0-9_]*$'

        return not re.match(sheet_name_pattern, sheetname)

    def _sheetname(self, sheetname):

        if sheetname.startswith('\'') and sheetname.endswith('\''):
            sheetname = sheetname[1:-1]

        if sheetname == self.old_name:
            sheetname = self.new_name

        if self.needs_quotes(sheetname):
            sheetname = f'\'{sheetname}\''

        return Token('SHEET_NAME', sheetname)

    def cell(self, items):
        # Renaming the sheet in the cell reference
        if len(items) == 2:
            return Tree(Token('RULE', 'cell'), [Token('SHEET_NAME', self._sheetname(items[0])), Token('CELLREF', items[1])])
        else:
            return Tree(Token('RULE', 'cell'), [Token('CELLREF', items[0])])