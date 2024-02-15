import lark
from lark.reconstruct import Reconstructor

class FormulaReconstructor():
    def __init__(self): 
        parser = lark.Lark.open("formulas.lark", start='formula', rel_to=__file__)
        parser.options.maybe_placeholders = False

        self._recon = Reconstructor(parser)

    def reconstruct_formula(self, tree):
        return "=" + self._recon.reconstruct(tree)
