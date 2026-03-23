"""
FILE: core/solver_python.py

PURPOSE: Python-based subset-sum solver using itertools.combinations().
         This is the fallback solver when the C DLL is not available.
         Correct for all datasets including those with negative numbers.

CONTAINS:
- find_combinations()  — Generator that yields Combination objects

DEPENDS ON:
- config/constants.py → EXACT_MATCH_THRESHOLD, PROGRESS_CHECK_INTERVAL,
                         SEARCH_ORDER_SMALLEST, SEARCH_ORDER_LARGEST
- models/number_item.py → NumberItem
- models/combination.py → Combination

USED BY:
- core/solver_manager.py → calls find_combinations() inside SolverThread

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — Python itertools solver         | Sub-phase 1B solver engine       |
| 22-03-2026 | max_results caps approx only, not exact   | Exact matches must never be lost |
"""

# Group 1: Python standard library
import itertools
from typing import List, Optional, Callable, Generator

# Group 3: This project's modules
from config.constants import (
    EXACT_MATCH_THRESHOLD,
    PROGRESS_CHECK_INTERVAL,
    SEARCH_ORDER_SMALLEST,
    SEARCH_ORDER_LARGEST,
)
from models.number_item import NumberItem
from models.combination import Combination


def find_combinations(
    items: List[NumberItem],
    target: float,
    buffer: float,
    min_size: int,
    max_size: int,
    max_results: int,
    search_order: str = SEARCH_ORDER_SMALLEST,
    seed_indices: Optional[List[int]] = None,
    stop_flag: Optional[Callable[[], bool]] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> Generator[Combination, None, None]:
    """
    WHAT:
        Generator that yields Combination objects whose item sums fall
        within target ± buffer. Uses itertools.combinations() to
        enumerate all subsets of each size from min_size to max_size.

        Supports seed numbers (must-include items), configurable search
        order (smallest-first or largest-first), and responsive stop
        via a callable flag checked every PROGRESS_CHECK_INTERVAL
        iterations.

    WHY ADDED:
        The Python solver is the fallback when the C DLL is unavailable.
        It uses itertools.combinations() because negative numbers make
        pruned backtracking and dynamic programming approaches unreliable.
        itertools is simple, correct, and handles all edge cases.

    CALLED BY:
        → core/solver_manager.py → SolverThread._run_python_solver()

    CALLS:
        → itertools.combinations() — Python built-in
        → models/combination.py → Combination constructor

    EDGE CASES HANDLED:
        - Negative values in items → handled naturally by itertools
        - Seed items → subtracted from target, excluded from search pool
        - Buffer = 0 → exact match mode (within EXACT_MATCH_THRESHOLD)
        - max_results reached for approx → stops collecting approx,
          but continues searching all sizes for exact matches
        - stop_flag returns True → stops immediately
        - Empty items list → yields nothing
        - Single item matching target → yields size-1 combination
        - All items are seeds → searches size-0 combos (just the seeds)

    ASSUMPTIONS:
        - Items have already been filtered to exclude finalized items.
        - min_size and max_size have already been validated and clamped
          (possibly by smart_bounds).
        - stop_flag is a callable returning True when the user clicks
          Stop. If None, the solver runs until completion. Checked both
          every PROGRESS_CHECK_INTERVAL iterations AND after every
          yielded result, so the solver stops promptly.
        - progress_callback receives (current_iteration_count, current_size)
          and is called every PROGRESS_CHECK_INTERVAL iterations.
        *** ASSUMPTION: max_results caps ONLY approximate matches.
            Exact matches are NEVER capped — every exact match is
            valuable and should always be reported. This prevents the
            scenario where approximate matches at smaller sizes fill
            the limit before exact matches at larger sizes are found.
            Decided in Session 2. ***
        *** ASSUMPTION: When seeds are present, min_size and max_size
            refer to the TOTAL combination size including seeds, not
            just the non-seed portion. The solver adjusts internally. ***

    PARAMETERS:
        items (List[NumberItem]): Available items (non-finalized only).
        target (float): Target sum to match.
        buffer (float): Tolerance range (target ± buffer).
        min_size (int): Minimum total combination size.
        max_size (int): Maximum total combination size.
        max_results (int): Cap for APPROXIMATE matches only. Exact
                           matches are never capped.
        search_order (str): "smallest_first" or "largest_first".
        seed_indices (Optional[List[int]]): Indices of pinned seed items.
        stop_flag (Optional[Callable]): Returns True to stop searching.
        progress_callback (Optional[Callable]): Called with (iterations, size).

    YIELDS:
        Combination: Each valid combination found (includes seed items).
    """
    if seed_indices is None:
        seed_indices = []

    # --- Separate seed items from searchable items ---
    seed_index_set = set(seed_indices)
    seed_items = [item for item in items if item.index in seed_index_set]
    search_items = [item for item in items if item.index not in seed_index_set]

    # Calculate adjusted target: subtract seed values
    seed_sum = sum(item.value for item in seed_items)
    adjusted_target = round(target - seed_sum, 2)
    seed_count = len(seed_items)

    # Matching bounds (adjusted target ± buffer)
    lower_bound = adjusted_target - buffer
    upper_bound = adjusted_target + buffer

    # --- Determine size range for the non-seed portion ---
    # Total combo size = seed_count + search_portion_size
    # So search_portion_size ranges from (min_size - seed_count) to (max_size - seed_count)
    search_min = max(min_size - seed_count, 0)
    search_max = min(max_size - seed_count, len(search_items))

    if search_min > search_max:
        return  # No valid sizes to search

    # --- Build size range based on search order ---
    if search_order == SEARCH_ORDER_LARGEST:
        size_range = range(search_max, search_min - 1, -1)
    else:
        # Default: smallest first
        size_range = range(search_min, search_max + 1)

    approx_found = 0
    iteration_count = 0

    for size in size_range:
        # Check stop flag at start of each size
        if stop_flag is not None and stop_flag():
            return

        if size == 0:
            # Special case: only seeds, no additional items needed
            combo_sum = seed_sum
            difference = round(combo_sum - target, 2)

            if abs(difference) <= buffer + EXACT_MATCH_THRESHOLD:
                combo = Combination(
                    items=list(seed_items),
                    target=target,
                )
                is_exact = abs(combo.difference) < EXACT_MATCH_THRESHOLD
                # Exact matches always yielded; approx only if under cap
                if is_exact or approx_found < max_results:
                    yield combo
                    if not is_exact:
                        approx_found += 1
            continue

        # Generate all combinations of this size from search items
        for combo_tuple in itertools.combinations(search_items, size):
            iteration_count += 1

            # Check stop flag periodically
            if iteration_count % PROGRESS_CHECK_INTERVAL == 0:
                if stop_flag is not None and stop_flag():
                    return
                if progress_callback is not None:
                    progress_callback(iteration_count, size + seed_count)

            # Check if this combination's sum is within bounds
            combo_sum = sum(item.value for item in combo_tuple)
            combo_sum_rounded = round(combo_sum, 2)

            if lower_bound - EXACT_MATCH_THRESHOLD <= combo_sum_rounded <= upper_bound + EXACT_MATCH_THRESHOLD:
                # Valid combination found — combine with seeds
                all_items = list(seed_items) + list(combo_tuple)
                combo = Combination(
                    items=all_items,
                    target=target,
                )

                is_exact = abs(combo.difference) < EXACT_MATCH_THRESHOLD

                # Exact matches are NEVER capped.
                # Approximate matches are capped at max_results.
                if is_exact:
                    yield combo
                elif approx_found < max_results:
                    yield combo
                    approx_found += 1
                # else: approx cap reached, skip this approx match
                # but keep searching for more exact matches

                # Check stop flag after every result
                if stop_flag is not None and stop_flag():
                    return

    # Final progress callback at end
    if progress_callback is not None:
        progress_callback(iteration_count, 0)
