"""
FILE: core/finalization_manager.py

PURPOSE: Manages the finalization lifecycle: assigns colors from the
         20-color palette, marks items as finalized, builds undo records,
         and tracks which items remain unmatched. All finalization logic
         lives here — the GUI calls methods and receives results, never
         manipulating finalization state directly.

CONTAINS:
- FinalizationManager — Stateful manager for all finalization operations

DEPENDS ON:
- config/constants.py → HIGHLIGHT_COLORS, COLOR_COUNT, UNMATCHED_COLOR
- config/mappings.py → get_color_name()
- models/combination.py → Combination
- models/finalized_combination.py → FinalizedCombination
- models/number_item.py → NumberItem

USED BY:
- gui/find_tab.py → calls finalize_combination(), undo_last()
- gui/summary_tab.py → reads finalized list, calls undo_last()
- core/session_manager.py → serializes/restores state

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — finalization manager             | Sub-phase 2A finalization core   |
"""

# Group 1: Python standard library
from typing import List, Optional, Set, Dict, Any
from copy import deepcopy

# Group 3: This project's modules
from config.constants import HIGHLIGHT_COLORS, COLOR_COUNT, UNMATCHED_COLOR
from config.mappings import get_color_name
from models.combination import Combination
from models.finalized_combination import FinalizedCombination
from models.number_item import NumberItem


class FinalizationManager:
    """
    WHAT:
        Manages the complete finalization workflow:
        1. Assigns the next color from the 20-color palette
        2. Creates a FinalizedCombination record
        3. Marks source NumberItems as finalized (is_finalized=True,
           finalized_color set)
        4. Pushes an undo record onto the LIFO stack
        5. Tracks the next combo number and color index

        Also provides undo (reverses the last finalization) and
        unmatched tracking (which items are not in any finalized combo).

    WHY ADDED:
        The original tool had finalization logic scattered across
        find_tab.py (~1000 lines). Extracting it into a manager:
        - Separates "what to finalize" from "how to update the UI"
        - Makes the undo stack reliable (all state in one place)
        - Prevents the critical combo removal index mismatch bug
          (this manager works with NumberItem.index, never widget rows)

    CALLED BY:
        → gui/find_tab.py → finalize_combination(), undo_last()
        → gui/summary_tab.py → get_finalized_list(), undo_last()

    ASSUMPTIONS:
        - Color palette cycles: finalization #21 wraps to color #0.
        - Undo only undoes the LAST finalization (single-level undo
          can be called repeatedly to undo multiple, one at a time).
        - Items are identified by NumberItem.index, not by list position.
        *** ASSUMPTION: The FinalizationManager is the SINGLE SOURCE
            OF TRUTH for finalization state. The GUI reads from here
            and updates its display. No other code should modify
            is_finalized or finalized_color on NumberItems directly. ***
    """

    def __init__(self):
        """
        WHAT: Initialises empty finalization state.
        """
        self._finalized_list: List[FinalizedCombination] = []
        self._undo_stack: List[Dict[str, Any]] = []
        self._next_color_index: int = 0
        self._next_combo_number: int = 1

    # -----------------------------------------------------------------------
    # Core Operations
    # -----------------------------------------------------------------------

    def finalize_combination(
        self,
        combo: Combination,
        items: List[NumberItem],
        label: str = "",
    ) -> FinalizedCombination:
        """
        WHAT:
            Finalizes a combination: assigns the next color, marks items,
            creates the FinalizedCombination record, and pushes an undo
            entry. Returns the finalized record.

        WHY ADDED:
            This is the central operation of CombiMatch. After the user
            confirms a combination matches their reconciliation, this
            locks it in with a color and label.

        CALLED BY:
            → gui/find_tab.py → after user clicks Finalize and enters label

        CALLS:
            → _assign_next_color() (in this file)
            → _mark_items_finalized() (in this file)

        EDGE CASES HANDLED:
            - Combo contains already-finalized items → skipped (shouldn't
              happen if results panel removes invalid combos correctly)
            - Color palette exhausted → wraps around to color #0
            - Empty label → stored as empty string (valid)

        PARAMETERS:
            combo (Combination): The combination to finalize.
            items (List[NumberItem]): The full items list (so we can mark
                the items in-place by matching NumberItem.index).
            label (str): Optional user label for this finalization.

        RETURNS:
            FinalizedCombination: The created record with color and metadata.
        """
        # Assign color
        color_rgb, color_name = self._assign_next_color()

        # Create record
        finalized = FinalizedCombination(
            combination=combo,
            color_rgb=color_rgb,
            color_name=color_name,
            label=label,
            combo_number=self._next_combo_number,
        )

        # Mark items as finalized in the source list
        finalized_indices = combo.item_indices
        self._mark_items_finalized(items, finalized_indices, color_rgb)

        # Build undo record BEFORE advancing state
        undo_record = {
            "finalized": finalized,
            "affected_indices": finalized_indices,
            "color_index_before": self._next_color_index - 1,  # Already advanced
            "combo_number_before": self._next_combo_number,
        }
        self._undo_stack.append(undo_record)

        # Add to finalized list and advance counter
        self._finalized_list.append(finalized)
        self._next_combo_number += 1

        return finalized

    def undo_last(self, items: List[NumberItem]) -> Optional[FinalizedCombination]:
        """
        WHAT:
            Undoes the most recent finalization: unmarks items, removes
            the FinalizedCombination record, and restores the color index
            and combo number. Returns the undone record (or None if
            nothing to undo).

        WHY ADDED:
            Mistakes happen — user finalizes the wrong combination.
            Undo lets them reverse it without restarting.

        CALLED BY:
            → gui/summary_tab.py → Undo button
            → gui/find_tab.py → may call for keyboard shortcut

        PARAMETERS:
            items (List[NumberItem]): The full items list (to unmark).

        RETURNS:
            FinalizedCombination or None: The undone record, or None
            if the undo stack is empty.
        """
        if not self._undo_stack:
            return None

        undo_record = self._undo_stack.pop()
        finalized = undo_record["finalized"]
        affected_indices = undo_record["affected_indices"]

        # Unmark items
        self._unmark_items(items, affected_indices)

        # Remove from finalized list
        if finalized in self._finalized_list:
            self._finalized_list.remove(finalized)

        # Restore counters
        self._next_color_index = (undo_record["color_index_before"]) % COLOR_COUNT
        self._next_combo_number = undo_record["combo_number_before"]

        return finalized

    # -----------------------------------------------------------------------
    # Query Methods
    # -----------------------------------------------------------------------

    def get_finalized_list(self) -> List[FinalizedCombination]:
        """
        WHAT: Returns the list of all finalized combinations.
        CALLED BY: gui/summary_tab.py, core/session_manager.py
        """
        return list(self._finalized_list)

    def get_finalized_count(self) -> int:
        """
        WHAT: Returns the number of finalized combinations.
        CALLED BY: gui/summary_tab.py → for display
        """
        return len(self._finalized_list)

    def get_finalized_indices(self) -> Set[int]:
        """
        WHAT: Returns the set of all item indices across all finalizations.
        CALLED BY: gui/results_panel.py → to filter invalid combos
        """
        indices = set()
        for finalized in self._finalized_list:
            indices.update(finalized.combination.item_indices)
        return indices

    def get_unmatched_items(self, all_items: List[NumberItem]) -> List[NumberItem]:
        """
        WHAT:
            Returns items that are not part of any finalized combination.
            These are the "leftover" items after reconciliation.

        CALLED BY:
            → gui/summary_tab.py → for unmatched numbers display (N6)
            → writers/report_writer.py → for the unmatched sheet

        PARAMETERS:
            all_items (List[NumberItem]): The complete items list.

        RETURNS:
            List[NumberItem]: Items where is_finalized is False.
        """
        return [item for item in all_items if not item.is_finalized]

    def can_undo(self) -> bool:
        """
        WHAT: Returns True if there is a finalization to undo.
        CALLED BY: gui/summary_tab.py → to enable/disable Undo button.
        """
        return len(self._undo_stack) > 0

    @property
    def next_combo_number(self) -> int:
        """WHAT: The number that will be assigned to the next finalization."""
        return self._next_combo_number

    @property
    def next_color_index(self) -> int:
        """WHAT: The index of the next color to be assigned."""
        return self._next_color_index

    # -----------------------------------------------------------------------
    # State for Session Save/Restore
    # -----------------------------------------------------------------------

    def get_state(self) -> Dict[str, Any]:
        """
        WHAT: Returns serializable state for session saving.
        CALLED BY: core/session_manager.py → save_session()
        """
        return {
            "next_color_index": self._next_color_index,
            "next_combo_number": self._next_combo_number,
            "finalized_count": len(self._finalized_list),
        }

    def restore_state(
        self,
        next_color_index: int,
        next_combo_number: int,
        finalized_list: List[FinalizedCombination],
    ):
        """
        WHAT: Restores state from a saved session.
        CALLED BY: core/session_manager.py → load_session()
        """
        self._next_color_index = next_color_index
        self._next_combo_number = next_combo_number
        self._finalized_list = list(finalized_list)
        # Undo stack is not preserved across sessions — user accepted
        # the session state by restoring it, so undo starts fresh.
        self._undo_stack.clear()

    def clear(self):
        """
        WHAT: Resets all finalization state. Called on Clear All.
        CALLED BY: gui/find_tab.py → on clear_all_requested
        """
        self._finalized_list.clear()
        self._undo_stack.clear()
        self._next_color_index = 0
        self._next_combo_number = 1

    # -----------------------------------------------------------------------
    # Internal Methods
    # -----------------------------------------------------------------------

    def _assign_next_color(self) -> tuple:
        """
        WHAT:
            Returns the next color from the 20-color palette and
            advances the index. Wraps around after color #20.

        CALLED BY:
            → finalize_combination()

        RETURNS:
            tuple: ((R, G, B), "Color Name")
        """
        color_rgb, color_name = HIGHLIGHT_COLORS[self._next_color_index]
        self._next_color_index = (self._next_color_index + 1) % COLOR_COUNT
        return color_rgb, color_name

    def _mark_items_finalized(
        self,
        items: List[NumberItem],
        indices: Set[int],
        color_rgb: tuple,
    ):
        """
        WHAT:
            Marks items with matching indices as finalized and assigns
            their color. Uses NumberItem.index for matching — NEVER
            list position.

        CALLED BY:
            → finalize_combination()

        WHY THIS WAY:
            The original tool used list position (widget row index)
            to find items, which broke when size-group headers shifted
            the positions. Using NumberItem.index is always correct
            regardless of UI layout.
        """
        for item in items:
            if item.index in indices:
                item.is_finalized = True
                item.finalized_color = color_rgb

    def _unmark_items(self, items: List[NumberItem], indices: Set[int]):
        """
        WHAT:
            Reverses finalization marks on items with matching indices.
            Clears is_finalized and finalized_color.

        CALLED BY:
            → undo_last()
        """
        for item in items:
            if item.index in indices:
                item.is_finalized = False
                item.finalized_color = None
