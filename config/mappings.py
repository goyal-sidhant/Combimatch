"""
FILE: config/mappings.py

PURPOSE: Lookup tables and mapping dictionaries used across the application.
         Currently holds color name lookups and search order option labels.

CONTAINS:
- COLOR_NAME_MAP            — Maps RGB tuples to human-readable color names
- SEARCH_ORDER_OPTIONS      — Maps internal search order keys to display labels
- get_color_name()          — Looks up a color name from an RGB tuple

DEPENDS ON:
- config/constants.py → HIGHLIGHT_COLORS, UNMATCHED_COLOR,
                         SEARCH_ORDER_SMALLEST, SEARCH_ORDER_LARGEST

USED BY:
- core/finalization_manager.py → get_color_name() for display
- gui/input_panel.py → SEARCH_ORDER_OPTIONS for dropdown labels

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — color and search order mappings | Project skeleton setup (Phase 1A)|
"""

# Group 3: This project's modules
from config.constants import (
    HIGHLIGHT_COLORS,
    UNMATCHED_COLOR,
    SEARCH_ORDER_SMALLEST,
    SEARCH_ORDER_LARGEST,
)


# ---------------------------------------------------------------------------
# Color Name Lookup
# ---------------------------------------------------------------------------

# Maps RGB tuples to their human-readable names.
# Built from HIGHLIGHT_COLORS and UNMATCHED_COLOR so there's a single
# source of truth in constants.py.
COLOR_NAME_MAP: dict[tuple[int, int, int], str] = {
    rgb: name for rgb, name in HIGHLIGHT_COLORS
}
COLOR_NAME_MAP[UNMATCHED_COLOR[0]] = UNMATCHED_COLOR[1]


def get_color_name(rgb: tuple[int, int, int]) -> str:
    """
    WHAT:
        Returns the human-readable name for an RGB color tuple.
        Falls back to "RGB(r, g, b)" if the color is not in the lookup.

    CALLED BY:
        → core/finalization_manager.py → when assigning colors
        → gui/summary_tab.py → when displaying finalized combo cards

    CALLS:
        → COLOR_NAME_MAP (dict lookup)

    RETURNS:
        str: Color name like "Light Blue" or fallback "RGB(173, 216, 230)".
    """
    return COLOR_NAME_MAP.get(rgb, f"RGB{rgb}")


# ---------------------------------------------------------------------------
# Search Order Options (for UI dropdown)
# ---------------------------------------------------------------------------

# Maps internal constant values to user-facing labels.
# Used by the search order dropdown in the input panel.
SEARCH_ORDER_OPTIONS: dict[str, str] = {
    SEARCH_ORDER_SMALLEST: "Smallest first",
    SEARCH_ORDER_LARGEST: "Largest first",
}
