import sys
from functools import lru_cache

from xlcalculator.xlfunctions import xl, func_xltypes

from . import ast_nodes, xltypes
import traceback

# optimise = False # 0.04087, 0.060144, 0.04104, 0.039606, 0.0392739, 0.04424
# optimise = True # 0.034895, 0.03976, 0.02205, 0.02350, 0.0282478, 0.26637

class EvaluatorContext(ast_nodes.EvalContext):

    def __init__(self, evaluator, ref):
        super().__init__(evaluator.namespace, ref)
        self.evaluator = evaluator

    @property
    def cells(self):
        return self.evaluator.model.cells

    @property
    def ranges(self):
        return self.evaluator.model.ranges

    @lru_cache(maxsize=None)
    def eval_cell(self, addr):
        # Check for a cycle.
        if addr in self.seen:
            raise RuntimeError(
                f'Cycle detected for {addr}:\n- ' + '\n- '.join(self.seen))
        self.seen.append(addr)

        return self.evaluator.evaluate(addr, None)


class Evaluator:
    """Traverses and evaluates a given model."""

    def __init__(self, model, namespace=None, freeze_cells_mode:bool=False):
        self.model = model
        self.namespace = namespace \
            if namespace is not None else xl.FUNCTIONS.copy()
        self.cache_count = 0
        self.model.freeze_cell_values = freeze_cells_mode

    def _get_context(self, ref):
        return EvaluatorContext(self, ref)

    # let's just set this through the initialisation for now (limit user error)    
    # def set_freeze_cells_mode(self, toggle: bool):
    #     self.model.freeze_cell_values = toggle

    # def get_freeze_cells_mode(self) -> bool:
    #     return self.model.freeze_cell_values

    def resolve_names(self, addr):
        # Although defined names have been resolved in Model.create_node()
        # we need to attempt to resolve defined names as we might have been
        # given one in argument addr.
        if addr not in self.model.defined_names:
            return addr

        defn = self.model.defined_names[addr]

        if isinstance(defn, xltypes.XLCell):
            return defn.address

        if isinstance(defn, xltypes.XLRange):
            raise ValueError(
                f"I can't resolve {addr} to a cell. It's a "
                f"range and they aren't supported yet.")

        if isinstance(defn, xltypes.XLFormula):
            raise ValueError(
                f"I can't resolve {addr} to a cell. It's a "
                f"formula and they aren't supported as a cell "
                f"reference.")

    def evaluate(self, addr, context=None):
        # 1. Resolve the address to a cell.
        addr = self.resolve_names(addr)

        if addr not in self.model.cells:
            # Blank cell that has no stored value in the model.
            return func_xltypes.BLANK
        cell = self.model.cells[addr]

        # 2. If there is no formula, we simply return the cell value.
        if cell.formula is None:
            return func_xltypes.ExcelType.cast_from_native(
                self.model.cells[addr].value)
        
        # if cell.formula.evaluate is False:
        if self.model.freeze_cell_values and addr in self.model.frozen_cells:
            # print(f"Won't evaluate: {addr}. Cell already has frozen value: {self.model.cells[addr].value}") # TEMP
            return func_xltypes.ExcelType.cast_from_native(self.model.cells[addr].value)            


        # 3. Prepare the execution environment and evaluate the formula.
        #    (Note: Range nodes will automatically evaluate all their
        #           dependencies.)
        context = context if context is not None else self._get_context(addr)
        try:

            value = cell.formula.ast.eval(context)

            if self.model.freeze_cell_values:
                # cell.formula.evaluate = False
                self.model.frozen_cells[addr] = 1
                # print(f"Value is frozen for: {addr} <= {value}") # TEMP
            # else:
            #     print(f">>> {addr} <= {value}")
        except Exception as err:
            # Joel 2024-06-03
            # raise RuntimeError(
            #     f"Problem evaluating cell {addr} formula "
            #     f"{cell.formula.formula}: {repr(err)}"
            # ).with_traceback(sys.exc_info()[2])
            print(f"Problem evaluating cell {addr} formula "
                f"{cell.formula.formula}: {repr(err)}")
            traceback.print_exc()
            raise RuntimeError(
                f"ERROR: Problem evaluating cell {addr} formula {cell.formula.formula}")
            

        # 4. Update the cell value.
        #    Note for later: If an array is returned, we should distribute the
        #    values to the respective cell (known as spilling). 
        # TODO - Add spilling.
        cell.value = value
        print(f"{addr} => {cell.value}") #TEMP

        # TODO: Joel - the property below seems to be dynamically added and with no further use. Need to check if this is needed at all...
        cell.need_update = False

        return value

    def set_cell_value(self, address, value):
        """Sets the value of a cell in the model."""
        self.model.set_cell_value(address, value)

    def get_cell_value(self, address):
        """Gets the value of a cell in the model."""
        return self.model.get_cell_value(address)
