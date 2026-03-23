"""
FILE: core/parameter_validator.py

PURPOSE: Validates all search parameters together before the solver runs.
         Checks target, buffer, min/max size, max results, and cross-validates
         that min <= max. Returns all errors at once so the user can fix
         everything in one pass.

CONTAINS:
- validate_search_parameters() — Full validation, returns SearchParameters or errors

DEPENDS ON:
- core/target_parser.py → parse_target(), parse_buffer()
- models/search_parameters.py → SearchParameters
- config/constants.py → DEFAULT_MIN_SIZE

USED BY:
- gui/find_tab.py → calls before starting a search
- gui/input_panel.py → may call for real-time validation

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — search parameter validation     | Sub-phase 1A foundation          |
"""

# Group 1: Python standard library
from typing import List

# Group 3: This project's modules
from config.constants import DEFAULT_MIN_SIZE, SEARCH_ORDER_SMALLEST
from core.target_parser import parse_target, parse_buffer
from models.search_parameters import SearchParameters


def validate_search_parameters(
    target_text: str,
    buffer_text: str,
    min_size: int,
    max_size: int,
    max_results: int,
    item_count: int,
    search_order: str = SEARCH_ORDER_SMALLEST,
    seed_indices: list = None,
) -> dict:
    """
    WHAT:
        Validates all search parameters and returns either a validated
        SearchParameters object or a list of error messages. All errors
        are collected and returned together so the user can fix everything
        in one pass rather than playing whack-a-mole.

    WHY ADDED:
        The solver assumes valid parameters. This function ensures that
        assumption holds before any search begins.

    CALLED BY:
        → gui/find_tab.py → _on_find_combinations()

    CALLS:
        → core/target_parser.py → parse_target(), parse_buffer()

    EDGE CASES HANDLED:
        - Target empty → error "Target sum is required"
        - Buffer empty → defaults to 0 (exact match mode)
        - min_size < 1 → clamped to 1
        - max_size > item_count → clamped to item_count
        - min_size > max_size → error
        - max_results < 1 → error
        - item_count = 0 → error "No numbers loaded"
        - Comma-formatted target/buffer → handled by parse_target/parse_buffer

    ASSUMPTIONS:
        - min_size and max_size come from spinboxes (always integers).
        - max_results comes from a spinbox (always integer).
        - item_count is the total loaded items minus finalized items.
        *** ASSUMPTION: search_order defaults to "smallest_first" because
            most users want to see small combinations first. ***

    PARAMETERS:
        target_text (str): Raw text from the target input field.
        buffer_text (str): Raw text from the buffer input field.
        min_size (int): Minimum combination size from spinbox.
        max_size (int): Maximum combination size from spinbox.
        max_results (int): Maximum results from spinbox.
        item_count (int): Number of available (non-finalized) items.
        search_order (str): "smallest_first" or "largest_first".
        seed_indices (list): Indices of pinned seed items.

    RETURNS:
        dict: {
            "success": bool,
            "error": str or None,       — Summary error message
            "data": {
                "params": SearchParameters or None,
                "errors": List[str]      — Individual error messages
            }
        }
    """
    errors: List[str] = []

    if seed_indices is None:
        seed_indices = []

    # --- Check item count first ---
    if item_count <= 0:
        return {
            "success": False,
            "error": "No numbers available for searching.",
            "data": {"params": None, "errors": ["Load numbers before searching."]}
        }

    # --- Parse target ---
    target_result = parse_target(target_text)
    if not target_result["success"]:
        errors.append(target_result["error"])
    target_value = target_result["data"]

    # --- Parse buffer ---
    buffer_result = parse_buffer(buffer_text)
    if not buffer_result["success"]:
        errors.append(buffer_result["error"])
    buffer_value = buffer_result["data"] if buffer_result["data"] is not None else 0.0

    # --- Validate min_size ---
    if min_size < DEFAULT_MIN_SIZE:
        min_size = DEFAULT_MIN_SIZE

    # --- Validate max_size ---
    if max_size > item_count:
        max_size = item_count
    if max_size < DEFAULT_MIN_SIZE:
        max_size = DEFAULT_MIN_SIZE

    # --- Cross-validate min/max ---
    if min_size > max_size:
        errors.append(
            f"Minimum size ({min_size}) cannot be greater than maximum size ({max_size})."
        )

    # --- Validate max_results ---
    if max_results < 1:
        errors.append("Maximum results must be at least 1.")

    # --- Return errors if any ---
    if errors:
        return {
            "success": False,
            "error": "Please fix the following issues:",
            "data": {"params": None, "errors": errors}
        }

    # --- All valid: build SearchParameters ---
    params = SearchParameters(
        target=target_value,
        buffer=buffer_value,
        min_size=min_size,
        max_size=max_size,
        max_results=max_results,
        search_order=search_order,
        seed_indices=seed_indices,
    )

    return {
        "success": True,
        "error": None,
        "data": {"params": params, "errors": []}
    }
