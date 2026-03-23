"""
FILE: core/solver_manager.py

PURPOSE: Orchestrates the search process. Selects the right solver
         (C DLL or Python fallback), runs it in a background QThread,
         and emits results in batches to keep the UI responsive.

CONTAINS:
- SolverThread    — QThread that runs the solver and emits batched signals
- SolverManager   — Checks DLL availability, creates SolverThread

DEPENDS ON:
- config/constants.py → BATCH_SIZE, BATCH_INTERVAL, EXACT_MATCH_THRESHOLD,
                         SEARCH_ORDER_SMALLEST
- config/settings.py → get_dll_path()
- core/solver_python.py → find_combinations()
- core/smart_bounds.py → compute_smart_bounds()
- models/search_parameters.py → SearchParameters
- models/combination.py → Combination
- models/number_item.py → NumberItem

USED BY:
- gui/find_tab.py → creates SolverManager, starts/stops SolverThread

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — solver orchestrator             | Sub-phase 1B solver engine       |
| 22-03-2026 | Updated — max_results caps approx only    | Exact matches must never be lost |
| 23-03-2026 | Added C solver support via solver_c.py    | Phase 6 performance solver       |
| 23-03-2026 | Enabled C solver in create_solver_thread   | Was hardcoded False, now auto    |
| 23-03-2026 | C solver now streams results live          | Batch dump felt broken to user   |
"""

# Group 1: Python standard library
import os
import time
from typing import List, Optional

# Group 2: Third-party libraries
from PyQt5.QtCore import QThread, pyqtSignal

# Group 3: This project's modules
from config.constants import (
    BATCH_SIZE,
    BATCH_INTERVAL,
    EXACT_MATCH_THRESHOLD,
    PROGRESS_SIGNAL_MIN_INTERVAL,
)
from config.settings import get_dll_path
from core.solver_python import find_combinations
from core.solver_c import load_solver_dll, find_combinations_c_streaming
from core.smart_bounds import compute_smart_bounds
from models.search_parameters import SearchParameters
from models.number_item import NumberItem
from models.combination import Combination


class SolverThread(QThread):
    """
    WHAT:
        Background thread that runs the solver and emits results in
        batches. Separates the computation from the UI thread so the
        app stays responsive during long searches.

    WHY ADDED:
        Running the solver on the UI thread freezes the entire app —
        buttons become unclickable, the window can't be resized, and
        Windows shows "Not Responding". A QThread with batched signals
        keeps the UI fully interactive.

    CALLED BY:
        → gui/find_tab.py → creates and starts this thread

    SIGNALS EMITTED:
        → results_batch(list)    — List of Combination objects (batch)
        → progress(int, int)     — (iterations_checked, current_size)
        → search_complete(dict)  — Final stats when search finishes
        → error_occurred(str)    — Error message if solver crashes

    ASSUMPTIONS:
        - Only one SolverThread runs at a time. The find_tab must stop
          any running thread before starting a new one.
        - Items passed in are already filtered (no finalized items).
        - SearchParameters are already validated.
        *** ASSUMPTION: Batching uses BATCH_SIZE (20) and BATCH_INTERVAL
            (0.1s) from constants. These values balance smoothness vs
            responsiveness — tested in the original tool. ***
    """

    # --- Signals ---
    # Emitted with a list of Combination objects (up to BATCH_SIZE at a time)
    results_batch = pyqtSignal(list)

    # Emitted periodically: (iterations_checked, current_combination_size)
    # Uses 'object' for iterations to avoid PyQt5 int→C int (32-bit) truncation.
    # Iteration counts above 2.1 billion would wrap to negative with pyqtSignal(int, int).
    progress = pyqtSignal(object, int)

    # Emitted when search finishes normally or is stopped.
    # Dict contains: {"total_found": int, "exact_count": int,
    #                  "approximate_count": int, "was_stopped": bool,
    #                  "iterations_checked": int}
    search_complete = pyqtSignal(dict)

    # Emitted if the solver throws an unexpected exception
    error_occurred = pyqtSignal(str)

    def request_stop(self):
        """
        WHAT: Sets the stop flag so the solver exits at next check point.
        CALLED BY: gui/find_tab.py → when user clicks Stop button.
        """
        self._stop_requested = True

    def _should_stop(self) -> bool:
        """
        WHAT: Returns True if stop has been requested.
        CALLED BY: Passed as stop_flag callable to find_combinations().
        """
        return self._stop_requested

    def __init__(
        self,
        items: List[NumberItem],
        params: SearchParameters,
        use_c_solver: bool = False,
        c_dll=None,
        parent=None,
    ):
        """
        WHAT: Initialises the solver thread with items and search parameters.

        PARAMETERS:
            items (List[NumberItem]): Available (non-finalized) items.
            params (SearchParameters): Validated search settings.
            use_c_solver (bool): True to use C DLL solver.
            c_dll: The loaded ctypes.CDLL (required if use_c_solver=True).
            parent: Qt parent object.
        """
        super().__init__(parent)
        self._items = items
        self._params = params
        self._use_c_solver = use_c_solver
        self._c_dll = c_dll
        self._stop_requested = False

    def run(self):
        """
        WHAT:
            Main thread execution. Runs the solver and collects results
            into batches before emitting signals.

        WHY:
            Called automatically by QThread.start(). Runs in a separate
            thread — NEVER call this directly (use start() instead).

        EDGE CASES HANDLED:
            - Solver throws exception → emits error_occurred signal
            - Stop requested → emits search_complete with was_stopped=True
            - Zero results → emits search_complete with total_found=0
        """
        try:
            if self._use_c_solver and self._c_dll is not None:
                self._run_c_solver()
            else:
                self._run_python_solver()
        except Exception as e:
            self.error_occurred.emit(f"Solver error: {str(e)}")

    def _run_python_solver(self):
        """
        WHAT:
            Runs the Python itertools solver and batches results.
            Emits results_batch every BATCH_SIZE items or BATCH_INTERVAL
            seconds, whichever comes first.

        CALLED BY:
            → self.run()

        CALLS:
            → core/solver_python.py → find_combinations()
        """
        batch = []
        last_emit_time = time.time()
        total_found = 0
        exact_count = 0
        approximate_count = 0
        last_iteration_count = 0
        last_progress_emit_time = 0.0

        def on_progress(iterations: int, current_size: int):
            """
            Progress callback passed to the solver.
            Throttled to emit at most every PROGRESS_SIGNAL_MIN_INTERVAL
            seconds to prevent flooding the UI signal queue on large searches.
            """
            nonlocal last_iteration_count, last_progress_emit_time
            last_iteration_count = iterations
            now = time.time()
            if (now - last_progress_emit_time) >= PROGRESS_SIGNAL_MIN_INTERVAL:
                self.progress.emit(iterations, current_size)
                last_progress_emit_time = now

        # Run the solver generator
        for combo in find_combinations(
            items=self._items,
            target=self._params.target,
            buffer=self._params.buffer,
            min_size=self._params.min_size,
            max_size=self._params.max_size,
            max_results=self._params.max_results,
            search_order=self._params.search_order,
            seed_indices=self._params.seed_indices,
            stop_flag=self._should_stop,
            progress_callback=on_progress,
        ):
            # Classify as exact or approximate
            if abs(combo.difference) < EXACT_MATCH_THRESHOLD:
                exact_count += 1
            else:
                approximate_count += 1
            total_found += 1

            batch.append(combo)

            # Emit batch when it reaches BATCH_SIZE or time interval elapsed
            now = time.time()
            if len(batch) >= BATCH_SIZE or (now - last_emit_time) >= BATCH_INTERVAL:
                self.results_batch.emit(batch)
                batch = []
                last_emit_time = now

        # Emit any remaining results in the final partial batch
        if batch:
            self.results_batch.emit(batch)

        # Emit completion signal
        self.search_complete.emit({
            "total_found": total_found,
            "exact_count": exact_count,
            "approximate_count": approximate_count,
            "was_stopped": self._stop_requested,
            "iterations_checked": last_iteration_count,
        })

    def _run_c_solver(self):
        """
        WHAT:
            Runs the C DLL solver with live result streaming. Each result
            is delivered from C → Python via a result callback, then
            batched and emitted to the UI — identical experience to the
            Python solver.

        WHY:
            Without streaming, the C solver collects all results internally
            and dumps them at once when complete. This makes the UI appear
            frozen during search and then floods it with thousands of
            results at once. Streaming fixes both problems.

        CALLED BY:
            → self.run() → when use_c_solver is True

        CALLS:
            → core/solver_c.py → find_combinations_c(on_result=...)
        """
        batch = []
        last_emit_time = time.time()
        total_found = 0
        exact_count = 0
        approximate_count = 0
        last_iteration_count = 0
        last_progress_emit_time = 0.0

        def on_progress(iterations: int, current_size: int):
            """Progress callback — throttled to avoid flooding UI signals."""
            nonlocal last_iteration_count, last_progress_emit_time
            last_iteration_count = iterations
            now = time.time()
            if (now - last_progress_emit_time) >= PROGRESS_SIGNAL_MIN_INTERVAL:
                self.progress.emit(iterations, current_size)
                last_progress_emit_time = now

        def on_result(combo: Combination):
            """
            WHAT: Result callback — called by C solver for each valid
                  combination found. Classifies, batches, and emits
                  to UI in real-time.
            """
            nonlocal batch, last_emit_time, total_found, exact_count
            nonlocal approximate_count

            # Classify as exact or approximate
            if abs(combo.difference) < EXACT_MATCH_THRESHOLD:
                exact_count += 1
            else:
                approximate_count += 1
            total_found += 1

            batch.append(combo)

            # Emit batch when full or time interval elapsed
            now = time.time()
            if len(batch) >= BATCH_SIZE or (now - last_emit_time) >= BATCH_INTERVAL:
                self.results_batch.emit(batch)
                batch = []
                last_emit_time = now

        # Run the C solver with streaming — results delivered via on_result
        find_combinations_c_streaming(
            dll=self._c_dll,
            items=self._items,
            target=self._params.target,
            buffer=self._params.buffer,
            min_size=self._params.min_size,
            max_size=self._params.max_size,
            max_results=self._params.max_results,
            search_order=self._params.search_order,
            seed_indices=self._params.seed_indices,
            stop_flag=self._should_stop,
            progress_callback=on_progress,
            on_result=on_result,
        )

        # Emit any remaining results in the final partial batch
        if batch:
            self.results_batch.emit(batch)

        # Emit completion signal
        self.search_complete.emit({
            "total_found": total_found,
            "exact_count": exact_count,
            "approximate_count": approximate_count,
            "was_stopped": self._stop_requested,
            "iterations_checked": last_iteration_count,
        })


class SolverManager:
    """
    WHAT:
        Manages solver selection and thread creation. Checks whether
        the C DLL is available, computes smart bounds, and creates
        SolverThread instances.

    WHY ADDED:
        The GUI should never call solver functions directly. This
        manager provides a clean interface: "here are items and params,
        give me a thread I can start." Swapping between C and Python
        solvers is invisible to the UI.

    CALLED BY:
        → gui/find_tab.py → to prepare and start searches

    CALLS:
        → core/smart_bounds.py → compute_smart_bounds()
        → config/settings.py → get_dll_path()
        → SolverThread constructor

    ASSUMPTIONS:
        - Only one search runs at a time. The caller must stop any
          existing thread before requesting a new one.
        *** ASSUMPTION: C DLL detection is by file existence only.
            If the DLL exists but is corrupt, the solver will fall
            back to Python after a load failure. ***
    """

    def __init__(self):
        """
        WHAT: Initialises the solver manager and attempts to load the DLL.
        """
        self._c_dll = None
        self._dll_available = False
        self._load_dll()
        self._current_thread: Optional[SolverThread] = None

    def _load_dll(self):
        """
        WHAT:
            Attempts to load solver.dll via ctypes. Sets _dll_available
            to True only if the DLL loads AND its function is callable.

        CALLED BY: __init__(), refresh_dll_status()
        """
        self._c_dll = load_solver_dll()
        self._dll_available = self._c_dll is not None

    def _check_dll(self) -> bool:
        """
        WHAT: Checks if solver.dll exists at the expected path.
        CALLED BY: Legacy — kept for compatibility.
        RETURNS: True if the DLL file exists.
        """
        dll_path = get_dll_path()
        return os.path.isfile(dll_path)

    def refresh_dll_status(self):
        """
        WHAT: Re-loads the DLL. Call after auto-compilation
              or if the user places a new DLL.
        CALLED BY: gui/settings_tab.py, main_window.py startup
        """
        self._load_dll()

    @property
    def is_dll_available(self) -> bool:
        """
        WHAT: Whether the C DLL solver is available.
        CALLED BY: gui/find_tab.py → to show solver status in UI.
        """
        return self._dll_available

    def compute_bounds(
        self,
        items: List[NumberItem],
        params: SearchParameters,
    ) -> dict:
        """
        WHAT:
            Computes smart bounds for the given items and parameters.
            Returns viable min/max sizes, search space estimate, and
            no-solution flag.

        CALLED BY:
            → gui/find_tab.py → before starting search and for UI hints

        CALLS:
            → core/smart_bounds.py → compute_smart_bounds()

        PARAMETERS:
            items (List[NumberItem]): Available (non-finalized) items.
            params (SearchParameters): Validated search settings.

        RETURNS:
            dict: Smart bounds result (see smart_bounds.py for format).
        """
        values = [item.value for item in items]
        return compute_smart_bounds(
            values=values,
            target=params.target,
            buffer=params.buffer,
            user_min_size=params.min_size,
            user_max_size=params.max_size,
        )

    def create_solver_thread(
        self,
        items: List[NumberItem],
        params: SearchParameters,
        smart_bounds_result: Optional[dict] = None,
    ) -> SolverThread:
        """
        WHAT:
            Creates a SolverThread ready to be started. Applies smart
            bounds to narrow the search range if provided. Selects
            C or Python solver based on DLL availability.

        WHY ADDED:
            Centralises thread creation so the GUI doesn't need to know
            about solver internals, smart bounds application, or
            C vs Python selection.

        CALLED BY:
            → gui/find_tab.py → when user clicks Find

        CALLS:
            → SolverThread constructor

        EDGE CASES HANDLED:
            - No smart bounds provided → uses params as-is
            - Smart bounds narrower than params → uses smart bounds
            - DLL available → sets use_c_solver=True (Phase 6)
            - DLL missing → uses Python solver (current default)

        PARAMETERS:
            items (List[NumberItem]): Available (non-finalized) items.
            params (SearchParameters): Validated search settings.
            smart_bounds_result (Optional[dict]): Result from compute_bounds().

        RETURNS:
            SolverThread: Ready to connect signals and start().
        """
        # Apply smart bounds to narrow the search range
        effective_params = params
        if smart_bounds_result and not smart_bounds_result["no_solution"]:
            # Create a new SearchParameters with narrowed min/max
            effective_params = SearchParameters(
                target=params.target,
                buffer=params.buffer,
                min_size=smart_bounds_result["smart_min"],
                max_size=smart_bounds_result["smart_max"],
                max_results=params.max_results,
                search_order=params.search_order,
                seed_indices=params.seed_indices,
            )

        # C solver selection — use DLL if available, else Python fallback
        use_c = self._dll_available

        thread = SolverThread(
            items=items,
            params=effective_params,
            use_c_solver=use_c,
            c_dll=self._c_dll,
        )
        self._current_thread = thread
        return thread

    def stop_current_search(self):
        """
        WHAT: Requests the current solver thread to stop.
        CALLED BY: gui/find_tab.py → when user clicks Stop button.
        """
        if self._current_thread is not None and self._current_thread.isRunning():
            self._current_thread.request_stop()
