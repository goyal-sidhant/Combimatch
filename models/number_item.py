"""
FILE: models/number_item.py

PURPOSE: Defines the NumberItem dataclass — the fundamental data unit
         representing a single number loaded into CombiMatch, whether
         from manual input or Excel.

CONTAINS:
- NumberItem — dataclass with value, index, source info, finalization state,
               and seed flag

DEPENDS ON:
- config/constants.py → ROUNDING_PRECISION
- models/source_tag.py → SourceTag

USED BY:
- core/number_parser.py → creates NumberItem for each parsed number
- readers/excel_reader.py → creates NumberItem for each grabbed Excel cell
- core/solver_python.py → uses NumberItem.value for combination sums
- core/finalization_manager.py → reads/writes is_finalized, finalized_color
- gui/source_panel.py → displays NumberItem in the source list
- Nearly every module touches NumberItem — it is the central data unit.

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — central number data structure   | Project skeleton setup (Phase 1A)|
"""

# Group 1: Python standard library
from dataclasses import dataclass, field
from typing import Optional

# Group 3: This project's modules
from config.constants import ROUNDING_PRECISION
from models.source_tag import SourceTag


@dataclass
class NumberItem:
    """
    WHAT:
        Represents a single number loaded into CombiMatch. Every number
        the user enters (manually or from Excel) becomes a NumberItem.
        Carries the numeric value, a sequential index, optional source
        tracking for Excel, finalization state, and seed flag.

    WHY ADDED:
        The solver, UI, and Excel highlighter all need to work with the
        same number objects. A shared dataclass ensures consistency —
        when the finalization manager marks an item as finalized, the
        source list, solver, and highlighter all see the same state.

    ASSUMPTIONS:
        - value is rounded to ROUNDING_PRECISION (2) decimal places in
          __post_init__. This matches real-world invoice data and prevents
          floating-point noise from creating false mismatches.
        - index is sequential among valid items (not the original line
          number or Excel row). It serves as the unique identifier within
          a single CombiMatch session.
    """

    # The numeric value, rounded to 2 decimal places on creation.
    value: float = 0.0

    # Sequential index among all loaded items (0-based).
    # Used as the unique identifier for this item within a session.
    index: int = 0

    # Where this number came from (Excel workbook/sheet/cell).
    # None for manually entered numbers.
    source: Optional[SourceTag] = None

    # Whether this item has been locked into a finalized combination.
    is_finalized: bool = False

    # The RGB color tuple assigned when finalized (e.g., (173, 216, 230)).
    # None if not finalized.
    finalized_color: Optional[tuple[int, int, int]] = None

    # Whether this item is pinned as a seed number (must-include).
    is_seed: bool = False

    # The original text representation (for display/debugging).
    # Set by the parser to show exactly what the user entered.
    original_text: str = ""

    def __post_init__(self):
        """
        WHAT: Rounds the value to ROUNDING_PRECISION decimal places.
        WHY: All real-world invoice data is 2 decimal places. Rounding
             on creation prevents floating-point noise from causing
             false mismatches during combination sum checks.
        """
        self.value = round(self.value, ROUNDING_PRECISION)
