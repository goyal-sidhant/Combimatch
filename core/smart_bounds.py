"""
FILE: core/smart_bounds.py

PURPOSE: Computes mathematically viable min/max combination sizes BEFORE
         the solver runs. Skips impossible sizes to dramatically reduce
         search space. Also estimates total combinations and detects
         no-solution scenarios.

CONTAINS:
- compute_smart_bounds()      — Main function: viable min/max, search space estimate
- estimate_search_space()     — Calculates total combinations across a size range

DEPENDS ON:
- config/constants.py → SEARCH_SPACE_WARNING_LIMIT

USED BY:
- core/solver_manager.py → calls before starting solver to narrow size range
- gui/find_tab.py → calls to show viable bounds hints in the UI
- gui/input_panel.py → calls to update bounds hint labels in real time
- gui/dialogs.py → uses search space estimate for warning dialog

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — smart bounds + search space est | Sub-phase 1B solver engine       |
"""

# Group 1: Python standard library
import math
from typing import List

# Group 3: This project's modules
from config.constants import SEARCH_SPACE_WARNING_LIMIT


def compute_smart_bounds(
    values: List[float],
    target: float,
    buffer: float,
    user_min_size: int,
    user_max_size: int,
) -> dict:
    """
    WHAT:
        Analyses the loaded numbers to determine the smallest and largest
        combination sizes that could POSSIBLY produce a valid result.
        Sizes outside this range are mathematically impossible and can
        be skipped, saving enormous amounts of search time.

        Also estimates the total number of combinations that will be
        checked, and flags when no solution is possible.

    WHY ADDED:
        The original tool displayed bounds hints but didn't actually use
        them — the solver still searched all sizes from min to max.
        This function makes bounds functional: the solver only searches
        sizes within the viable range, and the UI warns when the search
        space is too large.

    CALLED BY:
        → core/solver_manager.py → to narrow the search range
        → gui/find_tab.py → to show bounds hints in the UI
        → gui/input_panel.py → to update hints when parameters change

    CALLS:
        → estimate_search_space() (in this file)
        → math.comb() for combination counting

    EDGE CASES HANDLED:
        - Empty values list → no_solution = True
        - All values negative → smart_min may be large (many items needed)
        - All values positive → standard pruning, very effective
        - Mixed positive/negative → bounds are wider (less pruning)
        - Buffer of 0 → exact match mode, bounds still work
        - user_min > smart_max → no_solution = True
        - smart_min > user_max → no_solution = True
        - Single item matching target → smart_min = smart_max = 1

    ASSUMPTIONS:
        - values list contains only non-finalized items (finalized items
          have already been filtered out by the caller).
        - target and buffer are already validated (non-None, buffer >= 0).
        *** ASSUMPTION: We compute bounds using sorted values, which gives
            correct results for all-positive datasets. For mixed datasets
            with negatives, bounds are conservative (wider than strictly
            necessary) to avoid missing valid combinations. This is
            acceptable because correctness > speed. ***

    PARAMETERS:
        values (List[float]): The available number values (non-finalized).
        target (float): The target sum to match.
        buffer (float): The tolerance (target ± buffer).
        user_min_size (int): User's requested minimum combination size.
        user_max_size (int): User's requested maximum combination size.

    RETURNS:
        dict: {
            "smart_min": int,           — Viable minimum size (>= user_min)
            "smart_max": int,           — Viable maximum size (<= user_max)
            "no_solution": bool,        — True if no combo can possibly match
            "estimated_combinations": int, — Total combos to check in viable range
            "exceeds_warning_limit": bool, — True if estimate > SEARCH_SPACE_WARNING_LIMIT
            "user_min": int,            — Original user min (for UI display)
            "user_max": int,            — Original user max (for UI display)
            "item_count": int,          — Number of available items
        }
    """
    item_count = len(values)

    # --- Edge case: no items ---
    if item_count == 0:
        return {
            "smart_min": 0,
            "smart_max": 0,
            "no_solution": True,
            "estimated_combinations": 0,
            "exceeds_warning_limit": False,
            "user_min": user_min_size,
            "user_max": user_max_size,
            "item_count": 0,
        }

    # Clamp user sizes to valid range
    effective_min = max(user_min_size, 1)
    effective_max = min(user_max_size, item_count)

    if effective_min > effective_max:
        return {
            "smart_min": effective_min,
            "smart_max": effective_max,
            "no_solution": True,
            "estimated_combinations": 0,
            "exceeds_warning_limit": False,
            "user_min": user_min_size,
            "user_max": user_max_size,
            "item_count": item_count,
        }

    # Target bounds: combination sum must be in [lower_bound, upper_bound]
    lower_bound = target - buffer
    upper_bound = target + buffer

    # --- Compute smart minimum ---
    # Sort descending: take the LARGEST items first.
    # Cumulative sum tells us the maximum possible sum for k items.
    # Find smallest k where cumsum >= lower_bound.
    sorted_desc = sorted(values, reverse=True)
    smart_min = effective_min
    cumsum = 0.0
    for k in range(1, item_count + 1):
        cumsum += sorted_desc[k - 1]
        if cumsum >= lower_bound:
            smart_min = k
            break
    else:
        # Even using ALL items, we can't reach lower_bound
        # This means no solution if all values are positive.
        # With negatives, we might still find solutions at larger sizes,
        # so we only flag no_solution if sum of all items < lower_bound.
        if cumsum < lower_bound:
            return {
                "smart_min": effective_min,
                "smart_max": effective_max,
                "no_solution": True,
                "estimated_combinations": 0,
                "exceeds_warning_limit": False,
                "user_min": user_min_size,
                "user_max": user_max_size,
                "item_count": item_count,
            }

    # --- Compute smart maximum ---
    # Sort ascending: take the SMALLEST items first.
    # Cumulative sum tells us the minimum possible sum for k items.
    # Find first k where cumsum > upper_bound. Smart max = k - 1.
    sorted_asc = sorted(values)
    smart_max = effective_max
    cumsum = 0.0
    for k in range(1, item_count + 1):
        cumsum += sorted_asc[k - 1]
        if cumsum > upper_bound:
            smart_max = k - 1
            break
    # If we never exceeded upper_bound, smart_max stays at effective_max
    # (all items together are still within range — valid for negatives)

    # Clamp smart bounds within user bounds
    smart_min = max(smart_min, effective_min)
    smart_max = min(smart_max, effective_max)

    # --- Check for no-solution ---
    no_solution = smart_min > smart_max

    # --- Estimate search space ---
    if no_solution:
        estimated_combinations = 0
    else:
        estimated_combinations = estimate_search_space(
            item_count, smart_min, smart_max
        )

    exceeds_warning = estimated_combinations > SEARCH_SPACE_WARNING_LIMIT

    return {
        "smart_min": smart_min,
        "smart_max": smart_max,
        "no_solution": no_solution,
        "estimated_combinations": estimated_combinations,
        "exceeds_warning_limit": exceeds_warning,
        "user_min": user_min_size,
        "user_max": user_max_size,
        "item_count": item_count,
    }


def estimate_search_space(item_count: int, min_size: int, max_size: int) -> int:
    """
    WHAT:
        Calculates the total number of combinations that will be checked
        across all sizes from min_size to max_size. Uses math.comb()
        which is exact and fast.

    WHY ADDED:
        Needed to warn the user when the search space is dangerously
        large (e.g., C(100, 50) ≈ 10^29) before the solver starts.

    CALLED BY:
        → compute_smart_bounds() (in this file)

    CALLS:
        → math.comb() — Python built-in for binomial coefficient

    PARAMETERS:
        item_count (int): Total number of available items.
        min_size (int): Minimum combination size.
        max_size (int): Maximum combination size.

    RETURNS:
        int: Total number of combinations across all sizes.
             Capped at 10^15 to avoid extremely slow computation
             for absurdly large values (if it's that big, we already
             know it exceeds the warning limit).
    """
    total = 0
    cap = 10**15  # Cap to avoid spending time computing absurd values

    for size in range(min_size, max_size + 1):
        total += math.comb(item_count, size)
        if total > cap:
            return total  # Already way beyond any reasonable limit

    return total
