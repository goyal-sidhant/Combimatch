"""
FILE: models/search_parameters.py

PURPOSE: Defines the SearchParameters dataclass — holds all validated
         search settings passed from the UI to the solver.

CONTAINS:
- SearchParameters — dataclass with target, buffer, size range,
                     max results, search order, and seed indices

DEPENDS ON:
- config/constants.py → SEARCH_ORDER_SMALLEST (for default value)

USED BY:
- core/parameter_validator.py → creates SearchParameters after validation
- core/solver_python.py → reads all fields to configure the search
- core/solver_c.py → reads all fields for C solver invocation
- core/solver_manager.py → passes SearchParameters to the chosen solver
- core/session_manager.py → serializes for session save

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — search parameter container      | Project skeleton setup (Phase 1A)|
"""

# Group 1: Python standard library
from dataclasses import dataclass, field
from typing import List

# Group 3: This project's modules
from config.constants import SEARCH_ORDER_SMALLEST


@dataclass
class SearchParameters:
    """
    WHAT:
        All the settings needed to run a combination search, validated
        and packaged by parameter_validator.py. The solver reads these
        fields directly — no further parsing needed.

    WHY ADDED:
        Passing individual arguments (target, buffer, min_size, max_size,
        max_results, search_order, seed_indices) to the solver is fragile
        and hard to extend. A single dataclass keeps it clean.

    ASSUMPTIONS:
        - All values are pre-validated. The solver trusts these values
          and does not re-validate them.
        - seed_indices contains the index values of pinned seed items
          (from NumberItem.index), not list positions.
    """

    # The target sum to match.
    target: float = 0.0

    # The tolerance range. Combinations within target ± buffer are valid.
    buffer: float = 0.0

    # Minimum number of items in a combination.
    min_size: int = 1

    # Maximum number of items in a combination.
    max_size: int = 10

    # Maximum number of combinations to find before stopping.
    max_results: int = 25

    # Search order: "smallest_first" or "largest_first".
    search_order: str = SEARCH_ORDER_SMALLEST

    # Indices of items pinned as seeds (must-include).
    # Empty list means no seeds — normal search.
    seed_indices: List[int] = field(default_factory=list)
