"""
FILE: gui/find_tab.py

PURPOSE: The conductor of the Find tab. Wires together the input panel,
         results panel, and source panel. Manages the solver thread
         lifecycle: validate → compute bounds → start solver → receive
         batches → display results. Also manages the finalization flow:
         select combo → Finalize → label dialog → mark items → update UI.

CONTAINS:
- FindTab — QWidget that assembles and coordinates the three sub-panels

DEPENDS ON:
- config/constants.py → SPLITTER_SIZES
- gui/input_panel.py → InputPanel
- gui/results_panel.py → ResultsPanel
- gui/source_panel.py → SourcePanel
- gui/dialogs.py → ask_label, confirm_clear_all
- core/parameter_validator.py → validate_search_parameters
- core/solver_manager.py → SolverManager, SolverThread
- core/finalization_manager.py → FinalizationManager

USED BY:
- gui/main_window.py → embeds FindTab as the first tab

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — find tab assembly               | Sub-phase 1E results + find tab  |
| 22-03-2026 | Added finalization flow + undo support    | Sub-phase 2B finalization UI     |
| 22-03-2026 | Undo now restores removed results combos  | UI fixes before Phase 3          |
| 22-03-2026 | Added grab from Excel, WorkbookManager ref | Sub-phase 3A Excel integration   |
| 22-03-2026 | Added live smart bounds hints on param change | User couldn't see viable sizes  |
| 23-03-2026 | Wired seed indices into search flow       | Phase 4A seed numbers            |
| 23-03-2026 | Bounds hints now guidance, not validation  | User couldn't see viable range   |
| 23-03-2026 | Added solver indicator (C/Python)          | User needs to know which solver  |
| 23-03-2026 | Added get_session_state() and restore_session() | Phase 4C session persistence  |
"""

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QWidget, QHBoxLayout, QSplitter, QMessageBox,
)
from PyQt5.QtCore import Qt, pyqtSignal

# Group 3: This project's modules
from config.constants import SPLITTER_RATIOS
from gui.styles import scaled_px, get_screen_width
from gui.input_panel import InputPanel
from gui.results_panel import ResultsPanel
from gui.source_panel import SourcePanel
from gui.dialogs import ask_label, confirm_clear_all, confirm_multi_column, show_error, show_info
from core.parameter_validator import validate_search_parameters
from core.target_parser import parse_target, parse_buffer
from core.solver_manager import SolverManager, SolverThread
from core.finalization_manager import FinalizationManager
from readers.excel_workbook_manager import WorkbookManager


class FindTab(QWidget):
    """
    WHAT:
        The main working tab of CombiMatch. Assembles three panels
        (input, results, source) in a horizontal splitter and wires
        all signals together. Manages the full search lifecycle:
        load → validate → bounds → search → display → select.

    WHY ADDED:
        The original find_tab.py was ~1000 lines doing everything.
        This version delegates display to sub-panels and focuses
        solely on coordination — connecting signals, managing the
        solver thread, and routing data between panels.

    CALLED BY:
        → gui/main_window.py → added as the "Find" tab

    CALLS:
        → gui/input_panel.py → InputPanel
        → gui/results_panel.py → ResultsPanel
        → gui/source_panel.py → SourcePanel
        → gui/dialogs.py → ask_label(), confirm_clear_all()
        → core/parameter_validator.py → validate_search_parameters()
        → core/solver_manager.py → SolverManager
        → core/finalization_manager.py → FinalizationManager

    SIGNALS EMITTED:
        → finalization_changed() — Emitted after finalize, undo, or clear.
          MainWindow listens to refresh the Summary tab.

    ASSUMPTIONS:
        - Only one solver thread runs at a time. A new Find request
          stops any running search first.
        - The conductor pattern: FindTab does not display data directly,
          it routes data between panels via their public methods.
        *** ASSUMPTION: FinalizationManager is owned by FindTab because
            all finalization actions originate here (Finalize button,
            undo via MainWindow signal). The SummaryTab reads state
            via get_finalization_manager(). ***
    """

    finalization_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._solver_manager = SolverManager()
        self._finalization_manager = FinalizationManager()
        self._workbook_manager = None  # Set by MainWindow via set_workbook_manager()
        self._excel_highlighter = None  # Set by MainWindow via set_excel_highlighter()
        self._solver_thread = None
        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self):
        """
        WHAT: Creates the three-panel horizontal splitter layout.
        CALLED BY: __init__()
        """
        layout = QHBoxLayout(self)
        m = scaled_px(8)
        layout.setContentsMargins(m, m, m, m)

        self._splitter = QSplitter(Qt.Horizontal)

        self._input_panel = InputPanel()
        self._results_panel = ResultsPanel()
        self._source_panel = SourcePanel()

        self._splitter.addWidget(self._input_panel)
        self._splitter.addWidget(self._results_panel)
        self._splitter.addWidget(self._source_panel)
        # Compute proportional splitter widths from available screen width
        total = get_screen_width()
        self._splitter.setSizes([int(r * total) for r in SPLITTER_RATIOS])

        layout.addWidget(self._splitter)

    def _connect_signals(self):
        """
        WHAT: Wires all signals between panels and the solver manager.
        CALLED BY: __init__()
        """
        # Input panel signals
        self._input_panel.numbers_loaded.connect(self._on_numbers_loaded)
        self._input_panel.find_requested.connect(self._on_find_requested)
        self._input_panel.stop_requested.connect(self._on_stop_requested)
        self._input_panel.clear_all_requested.connect(self._on_clear_all)

        # Grab from Excel signal
        self._input_panel.grab_excel_requested.connect(self._on_grab_from_excel)

        # Smart bounds: recompute when target or buffer changes
        self._input_panel.target_changed.connect(self._update_bounds_hint)
        self._input_panel.buffer_changed.connect(self._update_bounds_hint)

        # Source panel signals
        self._source_panel.seeds_changed.connect(self._on_seeds_changed)

        # Results panel signals
        self._results_panel.combination_selected.connect(self._on_combination_selected)
        self._results_panel.deselection_requested.connect(self._on_deselection_requested)
        self._results_panel.finalize_requested.connect(self._on_finalize_requested)

        # Solver indicator — show which solver is active on startup
        self._input_panel.update_solver_indicator(
            self._solver_manager.is_dll_available
        )

    # -----------------------------------------------------------------------
    # Signal Handlers
    # -----------------------------------------------------------------------

    def set_workbook_manager(self, manager: WorkbookManager):
        """
        WHAT: Sets the WorkbookManager reference for Excel grab operations.
        CALLED BY: gui/main_window.py → after creating both FindTab and WorkbookManager.

        PARAMETERS:
            manager (WorkbookManager): The shared workbook manager.
        """
        self._workbook_manager = manager

    def set_excel_highlighter(self, highlighter):
        """
        WHAT: Sets the ExcelHighlighter reference for cell highlighting.
        CALLED BY: gui/main_window.py → after creating both FindTab and ExcelHighlighter.

        PARAMETERS:
            highlighter (ExcelHighlighter): The shared highlighter.
        """
        self._excel_highlighter = highlighter

    def _on_numbers_loaded(self, items):
        """
        WHAT:
            Handles the numbers_loaded signal from InputPanel.
            Passes items to the source panel for display.
            Computes smart bounds if target is already set.

        CALLED BY:
            → InputPanel.numbers_loaded signal

        PARAMETERS:
            items (list): List of NumberItem objects.
        """
        self._source_panel.load_items(items)
        self._results_panel.clear()
        # Show smart bounds immediately if target is already filled
        self._update_bounds_hint()

    def _on_grab_from_excel(self):
        """
        WHAT:
            Handles the Grab from Excel button click. Reads numbers from
            all checked workbook/sheet pairs via the WorkbookManager.
            Adds items to the existing source list (accumulative grab).

        CALLED BY:
            → InputPanel.grab_excel_requested signal

        EDGE CASES HANDLED:
            - No WorkbookManager set → error message
            - Not connected to Excel → error message (from WorkbookManager)
            - No sheets checked → error message pointing to Settings tab
            - Some sheets have no selection → continues with others
            - All sheets fail → error with details
            - Warnings (multi-column, filter fallback) → shown after success
        """
        if self._workbook_manager is None:
            show_error(self, "Not Available",
                       "Excel integration is not configured.")
            return

        # Force a live connection check — not just the cached flag.
        # is_connected does a COM call internally, but wrap it to catch
        # any unexpected COM exception (e.g., Excel crashed).
        try:
            is_alive = self._workbook_manager.excel_handler.is_connected
        except Exception:
            is_alive = False

        if not is_alive:
            show_error(self, "Not Connected",
                       "Excel is not running or the connection was lost.\n\n"
                       "Go to the Settings tab and click 'Connect to Excel' first.")
            return

        # Determine start index for accumulative grab
        existing_count = len(self._source_panel.get_all_items())
        start_index = existing_count

        result = self._workbook_manager.grab_from_checked(start_index)

        if not result["success"]:
            show_error(self, "Grab Failed", result["error"])
            return

        items = result["data"]["items"]
        errors = result["data"].get("errors", [])
        warnings = result["data"].get("warnings", [])

        # --- Multi-column check: ask user before proceeding ---
        has_multi_column = any(
            "multi-column" in w.lower() or "Multi-column" in w
            for w in warnings
        )
        if has_multi_column:
            if not confirm_multi_column(self):
                return
            # Remove the multi-column warning since user acknowledged it
            warnings = [
                w for w in warnings
                if "multi-column" not in w.lower() and "Multi-column" not in w
            ]

        # --- Duplicate detection: skip items already loaded from same cell ---
        if existing_count > 0:
            existing_items = self._source_panel.get_all_items()
            existing_sources = set()
            for ei in existing_items:
                if ei.source is not None:
                    existing_sources.add((
                        ei.source.workbook_name,
                        ei.source.sheet_name,
                        ei.source.cell_address,
                    ))

            if existing_sources:
                new_items = []
                skipped = 0
                next_index = existing_count
                for item in items:
                    if item.source is not None:
                        key = (
                            item.source.workbook_name,
                            item.source.sheet_name,
                            item.source.cell_address,
                        )
                        if key in existing_sources:
                            skipped += 1
                            continue
                    # Re-index to be sequential after existing items
                    item.index = next_index
                    next_index += 1
                    new_items.append(item)

                items = new_items
                if skipped > 0:
                    warnings.append(
                        f"{skipped} duplicate cell(s) skipped (already loaded)"
                    )

        if not items:
            show_info(self, "No New Data",
                      "All selected cells are already loaded.\n\n"
                      "Select different cells in Excel to grab new data.")
            return

        # Add to source panel (accumulative — doesn't replace)
        if existing_count > 0:
            self._source_panel.add_items(items)
        else:
            self._source_panel.load_items(items)

        # Clear previous results (new items may change combinations)
        self._results_panel.clear()
        # Show smart bounds immediately if target is already filled
        self._update_bounds_hint()

        # Show status
        status = f"Grabbed {len(items)} number{'s' if len(items) != 1 else ''} from Excel"
        if errors:
            status += f" ({len(errors)} skipped)"
        self._input_panel.set_load_status(status)

        # Show warnings if any
        if warnings:
            warning_text = "\n".join(f"  - {w}" for w in warnings)
            show_info(self, "Grab Warnings",
                      f"Numbers were loaded successfully, but with warnings:\n\n"
                      f"{warning_text}")

    def _on_find_requested(self):
        """
        WHAT:
            Handles the Find button click. Validates parameters,
            computes smart bounds, warns if search space is huge,
            and starts the solver thread.

        CALLED BY:
            → InputPanel.find_requested signal

        EDGE CASES HANDLED:
            - No numbers loaded → error message
            - Invalid parameters → error message with all issues
            - No solution possible → warning message
            - Huge search space → confirmation dialog
            - Already searching → stops previous search first
        """
        # Stop any running search (non-blocking — just set the flag and
        # disconnect old signals so stale results don't arrive)
        if self._solver_thread is not None and self._solver_thread.isRunning():
            self._solver_thread.request_stop()
            self._disconnect_solver_signals(self._solver_thread)
            # Reset UI from the old search's "searching" state
            self._input_panel.set_searching_state(False)

        # Get available items
        available_items = self._source_panel.get_available_items()
        item_count = len(available_items)

        # Collect seed indices from source panel
        seed_indices = self._source_panel.get_seed_indices()

        # Validate parameters
        result = validate_search_parameters(
            target_text=self._input_panel.get_target_text(),
            buffer_text=self._input_panel.get_buffer_text(),
            min_size=self._input_panel.get_min_size(),
            max_size=self._input_panel.get_max_size(),
            max_results=self._input_panel.get_max_results(),
            item_count=item_count,
            search_order=self._input_panel.get_search_order(),
            seed_indices=seed_indices,
        )

        if not result["success"]:
            errors = result["data"]["errors"]
            error_text = "\n".join(f"  - {e}" for e in errors)
            QMessageBox.warning(
                self,
                "Invalid Parameters",
                f"{result['error']}\n\n{error_text}",
            )
            return

        params = result["data"]["params"]

        # Compute smart bounds
        bounds = self._solver_manager.compute_bounds(available_items, params)
        self._input_panel.update_bounds_hint(bounds)

        # Check no-solution
        if bounds["no_solution"]:
            QMessageBox.information(
                self,
                "No Solution",
                "No combination of the loaded numbers can match the target "
                "with the current parameters.\n\n"
                "Try adjusting the target, buffer, or size range.",
            )
            return

        # Check search space warning
        if bounds["exceeds_warning_limit"]:
            est = bounds["estimated_combinations"]
            reply = QMessageBox.question(
                self,
                "Large Search Space",
                f"This search will check approximately {est:,} combinations.\n\n"
                f"This may take a very long time. You can click Stop at any "
                f"time to cancel.\n\n"
                f"Proceed anyway?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return

        # Clear previous results
        self._results_panel.clear()
        self._source_panel.highlight_combination(None)

        # Create and start solver thread
        self._solver_thread = self._solver_manager.create_solver_thread(
            available_items, params, bounds
        )
        self._solver_thread.results_batch.connect(self._on_results_batch)
        self._solver_thread.progress.connect(self._on_progress)
        self._solver_thread.search_complete.connect(self._on_search_complete)
        self._solver_thread.error_occurred.connect(self._on_solver_error)

        # Update solver indicator (may change if DLL was added/removed)
        self._input_panel.update_solver_indicator(
            self._solver_manager.is_dll_available
        )

        # Update UI state
        self._input_panel.set_searching_state(True)
        self._solver_thread.start()

    def _on_stop_requested(self):
        """
        WHAT:
            Stops the currently running solver thread and resets UI
            immediately. The thread continues running briefly until
            it hits the next stop-flag check, but signals are still
            connected so search_complete will fire normally.

        CALLED BY: InputPanel.stop_requested signal.
        """
        if self._solver_thread is not None and self._solver_thread.isRunning():
            self._solver_thread.request_stop()
            # Don't disconnect signals here — let search_complete fire
            # so stats are displayed. Just disable the Stop button
            # immediately for visual feedback.
            self._input_panel._stop_button.setEnabled(False)

    def _disconnect_solver_signals(self, thread):
        """
        WHAT:
            Disconnects all signals from a solver thread so stale results
            don't arrive after a new search starts. Non-blocking alternative
            to wait(2000) which froze the UI.

        CALLED BY:
            → _on_find_requested() → before starting a new search
            → _on_clear_all() → before clearing state

        PARAMETERS:
            thread (SolverThread): The thread to disconnect signals from.
        """
        try:
            thread.results_batch.disconnect()
            thread.progress.disconnect()
            thread.search_complete.disconnect()
            thread.error_occurred.disconnect()
        except TypeError:
            # Already disconnected — safe to ignore
            pass

    def _on_clear_all(self):
        """
        WHAT:
            Handles Clear All button. Confirms with user (stronger
            warning if finalized combos exist), then resets everything:
            input fields, results, source list, finalization state.

        CALLED BY:
            → InputPanel.clear_all_requested signal

        EDGE CASES HANDLED:
            - Search running → stops it first
            - No numbers loaded → still clears input fields
            - Finalized combos exist → stronger warning via dialog
        """
        # Stop any running search (non-blocking)
        if self._solver_thread is not None and self._solver_thread.isRunning():
            self._solver_thread.request_stop()
            self._disconnect_solver_signals(self._solver_thread)
            self._input_panel.set_searching_state(False)

        # Confirm with user (only if numbers are loaded)
        if self._source_panel.get_item_count() > 0:
            has_finalized = self._finalization_manager.get_finalized_count() > 0
            if not confirm_clear_all(self, has_finalized):
                return

        self._input_panel.clear_all()
        self._results_panel.clear()
        self._source_panel.clear_all_seeds()
        self._source_panel.clear()
        self._finalization_manager.clear()
        self.finalization_changed.emit()

    def _update_bounds_hint(self):
        """
        WHAT:
            Recomputes and displays smart bounds hints as GUIDANCE —
            showing the user what min/max sizes are viable given the
            current numbers, target, and buffer. Not gated on full
            parameter validation. Only needs a valid target to compute.

        CALLED BY:
            → _on_numbers_loaded() → after loading numbers
            → InputPanel.target_changed signal
            → InputPanel.buffer_changed signal
            → _on_seeds_changed() → when seeds are toggled

        EDGE CASES HANDLED:
            - No numbers loaded → clears hint
            - Target empty or unparseable → clears hint silently
            - Buffer empty → defaults to 0.0
        """
        available_items = self._source_panel.get_available_items()
        if not available_items:
            self._input_panel.clear_bounds_hint()
            return

        # Parse target — only need a valid number, skip full validation
        target_text = self._input_panel.get_target_text().strip()
        if not target_text:
            self._input_panel.clear_bounds_hint()
            return

        target_result = parse_target(target_text)
        if not target_result["success"]:
            self._input_panel.clear_bounds_hint()
            return

        target = target_result["data"]

        # Parse buffer — default to 0.0 if empty or invalid
        buffer_text = self._input_panel.get_buffer_text().strip()
        if buffer_text:
            buffer_result = parse_buffer(buffer_text)
            buffer_val = buffer_result["data"] if buffer_result["success"] else 0.0
        else:
            buffer_val = 0.0

        # Adjust target for seeds
        seed_indices = self._source_panel.get_seed_indices()
        seed_sum = 0.0
        if seed_indices:
            for item in available_items:
                if item.index in seed_indices:
                    seed_sum += item.value

        # Compute smart bounds over the full possible range (1 to item count)
        from core.smart_bounds import compute_smart_bounds
        values = [item.value for item in available_items]
        bounds = compute_smart_bounds(
            values=values,
            target=target,
            buffer=buffer_val,
            user_min_size=1,
            user_max_size=len(available_items),
        )
        self._input_panel.update_bounds_hint(bounds)

    def _on_combination_selected(self, combo):
        """
        WHAT:
            Handles combination selection in the results panel.
            Highlights the combo's items in the source panel.

        CALLED BY:
            → ResultsPanel.combination_selected signal

        PARAMETERS:
            combo (Combination): The selected combination.
        """
        self._source_panel.highlight_combination(combo)

    def _on_deselection_requested(self):
        """
        WHAT:
            Handles the deselection signal from ResultsPanel.
            Clears orange highlights from the source panel.

        CALLED BY:
            → ResultsPanel.deselection_requested signal
        """
        self._source_panel.highlight_combination(None)

    def _on_seeds_changed(self):
        """
        WHAT:
            Handles the seeds_changed signal from SourcePanel.
            Updates the seed info hint and recomputes smart bounds.

        CALLED BY:
            → SourcePanel.seeds_changed signal
        """
        # Update seed info display
        seed_indices = self._source_panel.get_seed_indices()
        if seed_indices:
            available = self._source_panel.get_available_items()
            items_by_index = {item.index: item for item in available}
            seed_sum = sum(
                items_by_index[idx].value for idx in seed_indices
                if idx in items_by_index
            )
            self._input_panel.update_seed_info(len(seed_indices), seed_sum)
        else:
            self._input_panel.clear_seed_info()

        # Recompute bounds with new seed set
        self._update_bounds_hint()

    def _on_finalize_requested(self):
        """
        WHAT:
            Handles the Finalize button click. Gets the selected combo,
            asks for a label, finalizes via FinalizationManager, then
            updates all three panels.

        CALLED BY:
            → ResultsPanel.finalize_requested signal

        EDGE CASES HANDLED:
            - No combo selected → shows error message
            - User cancels label dialog → aborts finalization
            - Combo contains already-finalized items → shouldn't happen
              because remove_invalid_combinations() cleans these up
        """
        combo = self._results_panel.get_selected_combination()
        if combo is None:
            QMessageBox.warning(
                self,
                "No Selection",
                "Please select a combination to finalize.",
            )
            return

        # Ask for optional label
        combo_number = self._finalization_manager.next_combo_number
        label_result = ask_label(self, combo_number)
        if not label_result["accepted"]:
            return

        # Finalize
        all_items = self._source_panel.get_all_items()
        finalized = self._finalization_manager.finalize_combination(
            combo, all_items, label=label_result["label"]
        )

        # Highlight cells in Excel immediately
        if self._excel_highlighter is not None:
            highlight_result = self._excel_highlighter.highlight_combination(finalized)
            if not highlight_result["success"] and highlight_result["error"]:
                # Non-fatal — show in status bar, don't block finalization
                self._input_panel.set_load_status(
                    f"Excel highlight: {highlight_result['error']}"
                )

        # Save results snapshot before removing (for undo)
        self._results_panel.save_results_snapshot()

        # Remove invalid combos from results (those containing finalized items)
        finalized_indices = self._finalization_manager.get_finalized_indices()
        self._results_panel.remove_invalid_combinations(finalized_indices)

        # Refresh source panel to show finalized colors
        self._source_panel.refresh_display()

        # Notify MainWindow to refresh Summary tab
        self.finalization_changed.emit()

    def undo_last_finalization(self):
        """
        WHAT:
            Undoes the most recent finalization. Called by MainWindow
            when the Summary tab's Undo button is clicked.

        CALLED BY:
            → gui/main_window.py → forwards SummaryTab.undo_requested

        RETURNS:
            bool: True if an undo was performed, False if nothing to undo.
        """
        all_items = self._source_panel.get_all_items()
        undone = self._finalization_manager.undo_last(all_items)

        if undone is None:
            return False

        # Remove highlights from Excel cells
        if self._excel_highlighter is not None:
            self._excel_highlighter.remove_highlight(undone)

        # Restore results to pre-finalization state
        self._results_panel.restore_results_snapshot()

        # Refresh source panel to remove finalized colors
        self._source_panel.refresh_display()

        # Notify MainWindow to refresh Summary tab
        self.finalization_changed.emit()

        return True

    def get_finalization_manager(self) -> FinalizationManager:
        """
        WHAT: Returns the FinalizationManager for external access.
        CALLED BY: gui/main_window.py → to pass state to SummaryTab.
        """
        return self._finalization_manager

    def get_all_items(self) -> list:
        """
        WHAT: Returns all loaded NumberItems from the source panel.
        CALLED BY: gui/main_window.py → to compute unmatched items for Mark Unmatched.
        """
        return self._source_panel.get_all_items()

    # -----------------------------------------------------------------------
    # Session Save / Restore
    # -----------------------------------------------------------------------

    def get_session_state(self) -> dict:
        """
        WHAT:
            Collects all state needed for session save: loaded items,
            finalization state, and current search parameters.

        CALLED BY:
            → gui/main_window.py → auto-save timer and close-event save

        RETURNS:
            dict with keys: items, finalized_list, next_color_index,
            next_combo_number, search_params
        """
        manager = self._finalization_manager
        return {
            "items": self._source_panel.get_all_items(),
            "finalized_list": manager.get_finalized_list(),
            "next_color_index": manager.next_color_index,
            "next_combo_number": manager.next_combo_number,
            "search_params": self._input_panel.get_search_params(),
        }

    def restore_session(
        self,
        items: list,
        finalized_list: list,
        next_color_index: int,
        next_combo_number: int,
        search_params: dict,
    ):
        """
        WHAT:
            Restores the Find tab state from a loaded session. Loads items
            into the source panel, restores finalization state, and sets
            search parameters.

        CALLED BY:
            → gui/main_window.py → after user confirms session restore

        PARAMETERS:
            items (list): List of NumberItem objects to load.
            finalized_list (list): List of FinalizedCombination objects.
            next_color_index (int): Next color index for finalization.
            next_combo_number (int): Next combo sequence number.
            search_params (dict): Search parameter values to restore.
        """
        # Load items into source panel
        if items:
            self._source_panel.load_items(items)

        # Restore finalization state
        self._finalization_manager.restore_state(
            next_color_index=next_color_index,
            next_combo_number=next_combo_number,
            finalized_list=finalized_list,
        )

        # Restore search parameters
        if search_params:
            self._input_panel.set_search_params(search_params)

        # Refresh source panel to show finalized colors
        self._source_panel.refresh_display()

        # Update bounds hints if target is set
        self._update_bounds_hint()

        # Notify MainWindow to refresh Summary tab
        self.finalization_changed.emit()

    # -----------------------------------------------------------------------
    # Solver Thread Signal Handlers
    # -----------------------------------------------------------------------

    def _on_results_batch(self, batch):
        """
        WHAT: Adds a batch of results to the results panel.
        CALLED BY: SolverThread.results_batch signal.
        """
        self._results_panel.add_results(batch)

    def _on_progress(self, iterations, current_size):
        """
        WHAT: Updates the progress display in the input panel.
        CALLED BY: SolverThread.progress signal.
        """
        self._input_panel.update_progress(iterations, current_size)

    def _on_search_complete(self, stats):
        """
        WHAT: Handles search completion. Updates UI and shows stats.
        CALLED BY: SolverThread.search_complete signal.
        """
        self._input_panel.set_searching_state(False)
        self._input_panel.show_search_complete(stats)

    def _on_solver_error(self, error_message):
        """
        WHAT: Handles solver errors. Shows error dialog.
        CALLED BY: SolverThread.error_occurred signal.
        """
        self._input_panel.set_searching_state(False)
        QMessageBox.critical(
            self,
            "Solver Error",
            f"An error occurred during the search:\n\n{error_message}",
        )
