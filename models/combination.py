"""
FILE: models/combination.py

PURPOSE: Defines the Combination dataclass — represents a group of
         NumberItems whose values sum to (or near) the target.

CONTAINS:
- Combination — dataclass with items list, target, and computed properties
                for sum_value, difference, and size

DEPENDS ON:
- models/number_item.py → NumberItem

USED BY:
- core/solver_python.py → creates Combination objects when a valid subset is found
- core/solver_c.py → creates Combination objects from C solver results
- gui/results_panel.py → displays combinations in exact/approximate lists
- gui/combo_info_panel.py → shows selected combination details
- core/finalization_manager.py → receives Combination for finalization

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — combination result data model   | Project skeleton setup (Phase 1A)|
"""

# Group 1: Python standard library
from dataclasses import dataclass, field
from typing import List

# Group 3: This project's modules
from models.number_item import NumberItem


@dataclass
class Combination:
    """
    WHAT:
        A group of NumberItems found by the solver whose sum falls within
        the target ± buffer range. Carries the items, the target that was
        searched for, and computed properties for sum, difference, and size.

    WHY ADDED:
        The solver yields raw item groups; the UI and finalization manager
        need structured objects with pre-computed display values. This
        dataclass provides that consistent interface.

    ASSUMPTIONS:
        - items list is never empty (solver only yields valid combinations).
        - target is the original search target (not adjusted for seeds —
          seed adjustment happens before the solver runs, so the target
          passed here is already the "remaining target").
    """

    # The NumberItems in this combination.
    items: List[NumberItem] = field(default_factory=list)

    # The target value this combination was searched against.
    target: float = 0.0

    @property
    def sum_value(self) -> float:
        """
        WHAT: Sum of all item values in this combination.
        CALLED BY: gui/combo_info_panel.py, gui/results_panel.py
        RETURNS: float — the total sum, rounded to 2 decimal places.
        """
        return round(sum(item.value for item in self.items), 2)

    @property
    def difference(self) -> float:
        """
        WHAT: How far this combination's sum is from the target.
              Positive means over target, negative means under.
        CALLED BY: gui/results_panel.py (for exact/approx classification),
                   gui/combo_info_panel.py (for display)
        RETURNS: float — (sum - target), rounded to 2 decimal places.
        """
        return round(self.sum_value - self.target, 2)

    @property
    def size(self) -> int:
        """
        WHAT: Number of items in this combination.
        CALLED BY: gui/results_panel.py (for size-group headers)
        RETURNS: int — count of items.
        """
        return len(self.items)

    @property
    def item_indices(self) -> set:
        """
        WHAT: Set of indices of all items in this combination.
              Used for fast overlap checking during finalization removal.
        CALLED BY: core/finalization_manager.py → to check if a combo
                   contains any finalized items
        RETURNS: set of int — the index values of contained items.
        """
        return {item.index for item in self.items}
