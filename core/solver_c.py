"""
FILE: core/solver_c.py

PURPOSE: Python wrapper for the C solver DLL (solver.dll). Translates Python
         data structures (NumberItem, SearchParameters) into C arrays, calls
         the DLL via ctypes, and converts results back to Combination objects.
         Must produce IDENTICAL results to core/solver_python.py.

CONTAINS:
- load_solver_dll()        — Loads the DLL and configures function signatures
- find_combinations_c()    — Wrapper that mirrors solver_python.find_combinations()

DEPENDS ON:
- config/settings.py → get_dll_path()
- config/constants.py → EXACT_MATCH_THRESHOLD, PROGRESS_CHECK_INTERVAL,
                         SEARCH_ORDER_SMALLEST, SEARCH_ORDER_LARGEST
- models/number_item.py → NumberItem
- models/combination.py → Combination

USED BY:
- core/solver_manager.py → calls find_combinations_c() when DLL is available

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 23-03-2026 | Created — ctypes wrapper for C solver DLL | Phase 6 performance solver       |
| 23-03-2026 | Added streaming via RESULT_CALLBACK       | Results must stream live to UI   |
"""

# Group 1: Python standard library
import ctypes
import os
from typing import List, Optional, Callable, Generator

# Group 3: This project's modules
from config.settings import get_dll_path
from config.constants import (
    EXACT_MATCH_THRESHOLD,
    PROGRESS_CHECK_INTERVAL,
    SEARCH_ORDER_SMALLEST,
    SEARCH_ORDER_LARGEST,
)
from models.number_item import NumberItem
from models.combination import Combination


# --- C struct mirror ---
class CombinationResult(ctypes.Structure):
    """
    WHAT: ctypes mirror of the CombinationResult struct in solver.h.
    """
    _fields_ = [
        ("indices", ctypes.c_int32 * 512),
        ("count", ctypes.c_int32),
        ("sum_value", ctypes.c_double),
        ("difference", ctypes.c_double),
    ]


# Progress callback type: int callback(int64_t iterations, int32_t current_size)
PROGRESS_CALLBACK = ctypes.CFUNCTYPE(
    ctypes.c_int,          # return type (0 = continue, non-zero = stop)
    ctypes.c_int64,        # iterations
    ctypes.c_int32,        # current_size
)

# Result callback type: int callback(const CombinationResult *result)
# Called for each valid combination found — enables live streaming.
RESULT_CALLBACK = ctypes.CFUNCTYPE(
    ctypes.c_int,                          # return type (0 = continue, non-zero = stop)
    ctypes.POINTER(CombinationResult),     # result
)


def load_solver_dll():
    """
    WHAT:
        Loads the solver DLL and configures function argument/return types.
        Returns the loaded DLL object, or None if loading fails.

    CALLED BY:
        → core/solver_manager.py → on startup to check DLL availability

    RETURNS:
        ctypes.CDLL or None: The loaded DLL, or None if not found/loadable.
    """
    dll_path = get_dll_path()
    if not os.path.isfile(dll_path):
        return None

    try:
        dll = ctypes.CDLL(dll_path)
    except OSError:
        return None

    # Configure find_combinations_c signature
    dll.find_combinations_c.argtypes = [
        ctypes.POINTER(ctypes.c_double),   # values
        ctypes.POINTER(ctypes.c_int32),    # indices
        ctypes.c_int32,                    # item_count
        ctypes.c_double,                   # target
        ctypes.c_double,                   # buffer
        ctypes.c_int32,                    # min_size
        ctypes.c_int32,                    # max_size
        ctypes.c_int32,                    # max_results
        ctypes.c_int32,                    # search_order
        ctypes.POINTER(ctypes.c_int32),    # seed_flags
        ctypes.POINTER(CombinationResult), # results (ignored in streaming mode)
        ctypes.c_int32,                    # max_result_buf (ignored in streaming mode)
        PROGRESS_CALLBACK,                 # progress_cb
        ctypes.c_double,                   # exact_match_threshold
        RESULT_CALLBACK,                   # result_cb (NULL for batch, non-NULL for streaming)
    ]
    dll.find_combinations_c.restype = ctypes.c_int64

    return dll


def _prepare_c_args(
    items: List[NumberItem],
    search_order: str,
    seed_indices: Optional[List[int]],
    stop_flag: Optional[Callable[[], bool]],
    progress_callback: Optional[Callable[[int, int], None]],
):
    """
    WHAT: Converts Python data structures to C arrays. Shared by both
          batch and streaming modes.
    CALLED BY: find_combinations_c(), find_combinations_c_streaming()
    RETURNS: Tuple of (c_values, c_indices, c_seed_flags, c_search_order,
             c_progress_cb, items_by_index, item_count)
    """
    if seed_indices is None:
        seed_indices = []

    item_count = len(items)

    c_values = (ctypes.c_double * item_count)()
    c_indices = (ctypes.c_int32 * item_count)()
    c_seed_flags = (ctypes.c_int32 * item_count)()

    seed_index_set = set(seed_indices)
    items_by_index = {}

    for i, item in enumerate(items):
        c_values[i] = item.value
        c_indices[i] = item.index
        c_seed_flags[i] = 1 if item.index in seed_index_set else 0
        items_by_index[item.index] = item

    c_search_order = 1 if search_order == SEARCH_ORDER_LARGEST else 0

    def c_progress(iterations, current_size):
        if stop_flag is not None and stop_flag():
            return 1
        if progress_callback is not None:
            progress_callback(int(iterations), int(current_size))
        return 0

    c_progress_cb = PROGRESS_CALLBACK(c_progress)

    return (c_values, c_indices, c_seed_flags, c_search_order,
            c_progress_cb, items_by_index, item_count)


def find_combinations_c_streaming(
    dll,
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
    on_result: Optional[Callable] = None,
) -> None:
    """
    WHAT:
        Streaming wrapper for the C solver. Each result is delivered
        live via the on_result callback as it's found — no buffering.
        This is NOT a generator — it's a regular function that calls
        on_result(Combination) for each result found.

    WHY ADDED:
        Generator functions in Python don't execute until iterated.
        The streaming path needs to run immediately (called by
        SolverThread._run_c_solver), so it must be a separate
        non-generator function.

    CALLED BY:
        → core/solver_manager.py → SolverThread._run_c_solver()

    PARAMETERS:
        on_result: Callback(Combination) called for each result found.
        (All other params match find_combinations_c exactly)
    """
    if len(items) == 0:
        return

    (c_values, c_indices, c_seed_flags, c_search_order,
     c_progress_cb, items_by_index, item_count) = _prepare_c_args(
        items, search_order, seed_indices, stop_flag, progress_callback,
    )

    def c_result_handler(result_ptr):
        """
        WHAT: C result callback bridge. Converts CombinationResult from C
              into a Combination object and delivers it via on_result.
        """
        r = result_ptr.contents
        combo_items = []
        for j in range(r.count):
            item_index = r.indices[j]
            if item_index in items_by_index:
                combo_items.append(items_by_index[item_index])

        if combo_items:
            combo = Combination(items=combo_items, target=target)
            on_result(combo)
        return 0  # Continue

    c_result_cb = RESULT_CALLBACK(c_result_handler)

    # Call C solver — results streamed via callback, buffer unused
    dll.find_combinations_c(
        c_values,
        c_indices,
        ctypes.c_int32(item_count),
        ctypes.c_double(target),
        ctypes.c_double(buffer),
        ctypes.c_int32(min_size),
        ctypes.c_int32(max_size),
        ctypes.c_int32(max_results),
        ctypes.c_int32(c_search_order),
        c_seed_flags,
        None,                              # results — not used in streaming
        ctypes.c_int32(0),                 # max_result_buf — not used
        c_progress_cb,
        ctypes.c_double(EXACT_MATCH_THRESHOLD),
        c_result_cb,
    )


def find_combinations_c(
    dll,
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
        Batch wrapper that calls the C solver DLL and yields Combination
        objects. Mirrors the interface of core/solver_python.find_combinations()
        exactly, so tests and the solver_manager can swap transparently.

    WHY:
        This is a generator (uses yield), so it collects all C results
        in a buffer then yields them one at a time. Used by the test
        suite and as fallback. For live UI streaming, use
        find_combinations_c_streaming() instead.

    CALLED BY:
        → tests/test_c_solver_verification.py → via wrapper
        → core/solver_manager.py → (streaming uses find_combinations_c_streaming)

    PARAMETERS:
        dll: The loaded ctypes.CDLL from load_solver_dll().
        (All other params match solver_python.find_combinations exactly)

    YIELDS:
        Combination: Each valid combination found.

    ASSUMPTIONS:
        - The DLL has been loaded successfully (dll is not None).
        - Items are already filtered (no finalized items).
        *** ASSUMPTION: The C solver uses the same PROGRESS_CHECK_INTERVAL
            (1000) as the Python solver. The progress callback serves
            double duty: reporting progress AND checking the stop flag.
            Returns 0 to continue, non-zero to stop. ***
    """
    if len(items) == 0:
        return

    (c_values, c_indices, c_seed_flags, c_search_order,
     c_progress_cb, items_by_index, item_count) = _prepare_c_args(
        items, search_order, seed_indices, stop_flag, progress_callback,
    )

    # Allocate results buffer
    max_result_buf = max(max_results * 2, 10000)
    results_array = (CombinationResult * max_result_buf)()

    # Null result callback — use buffer mode
    c_null_result_cb = RESULT_CALLBACK(0)

    result_count = dll.find_combinations_c(
        c_values,
        c_indices,
        ctypes.c_int32(item_count),
        ctypes.c_double(target),
        ctypes.c_double(buffer),
        ctypes.c_int32(min_size),
        ctypes.c_int32(max_size),
        ctypes.c_int32(max_results),
        ctypes.c_int32(c_search_order),
        c_seed_flags,
        results_array,
        ctypes.c_int32(max_result_buf),
        c_progress_cb,
        ctypes.c_double(EXACT_MATCH_THRESHOLD),
        c_null_result_cb,
    )

    # Convert C results to Combination objects
    for i in range(result_count):
        r = results_array[i]
        combo_items = []
        for j in range(r.count):
            item_index = r.indices[j]
            if item_index in items_by_index:
                combo_items.append(items_by_index[item_index])

        if combo_items:
            combo = Combination(items=combo_items, target=target)
            yield combo
