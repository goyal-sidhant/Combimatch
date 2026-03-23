"""
FILE: gui/main_window.py

PURPOSE: The main application window. Contains the tab widget with
         Find, Summary, and Settings tabs. Handles window sizing,
         close events, and top-level coordination.

CONTAINS:
- MainWindow — QMainWindow subclass, the root of the GUI

DEPENDS ON:
- config/constants.py → WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT
- gui/styles.py → color constants (for status bar)
- gui/find_tab.py → FindTab (real Find tab)
- gui/summary_tab.py → SummaryTab (real Summary tab)
- gui/settings_tab.py → SettingsTab
- writers/excel_highlighter.py → ExcelHighlighter
- readers/excel_monitor.py → ExcelMonitor

USED BY:
- main.py → creates and shows MainWindow

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — main window shell with 3 tabs   | Sub-phase 1C app shell           |
| 22-03-2026 | Replaced Find placeholder with FindTab    | Sub-phase 1E find tab assembly   |
| 22-03-2026 | Replaced Summary placeholder, wired undo  | Sub-phase 2B finalization UI     |
| 22-03-2026 | Replaced Settings placeholder, wired Excel | Sub-phase 3A Excel integration   |
| 22-03-2026 | Enhanced close event, monitor reopen, highlighter | Sub-phase 3B Excel highlighting |
| 23-03-2026 | Wired Mark Unmatched button to ExcelHighlighter | Phase 4B unmatched numbers view  |
| 23-03-2026 | Added session save/restore + auto-save timer | Phase 4C session persistence     |
"""

# Group 1: Python standard library
# (none)

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QMainWindow, QTabWidget, QWidget, QVBoxLayout, QLabel,
    QStatusBar,
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QScreen

# Group 3: This project's modules
from config.constants import (
    WINDOW_MIN_WIDTH_FRACTION, WINDOW_MIN_HEIGHT_FRACTION,
    WINDOW_MIN_WIDTH_FALLBACK, WINDOW_MIN_HEIGHT_FALLBACK,
    EXCEL_MONITOR_INTERVAL, UNMATCHED_COLOR, SESSION_AUTOSAVE_INTERVAL,
)
from core.session_manager import SessionManager
from gui.styles import COLOR_TEXT_SECONDARY, scaled_size, get_screen_width, get_screen_height
from gui.find_tab import FindTab
from gui.summary_tab import SummaryTab
from gui.settings_tab import SettingsTab
from readers.excel_reader import ExcelHandler
from readers.excel_workbook_manager import WorkbookManager
from readers.excel_monitor import ExcelMonitor
from writers.excel_highlighter import ExcelHighlighter


class MainWindow(QMainWindow):
    """
    WHAT:
        The root window of CombiMatch. Contains a QTabWidget with
        three tabs: Find, Summary, and Settings. Manages window
        geometry, close events, and status bar messages.

    WHY ADDED:
        Every PyQt5 app needs a main window. This one holds the tab
        container and coordinates top-level events (close, resize).

    CALLED BY:
        → main.py → creates, shows, and enters event loop

    CALLS:
        → (Will call tab constructors in later sub-phases)

    ASSUMPTIONS:
        - Only one MainWindow exists per app instance.
        - The window starts centered on the primary screen.
        *** ASSUMPTION: Using QScreen instead of the deprecated
            QDesktopWidget for screen geometry. QDesktopWidget was
            removed in Qt 6 and generates warnings in recent Qt 5. ***
    """

    def __init__(self, parent=None):
        """
        WHAT: Sets up the main window with tabs and status bar.
        """
        super().__init__(parent)
        self.setWindowTitle("CombiMatch")

        # Compute proportional minimum size from screen dimensions
        sw = get_screen_width()
        sh = get_screen_height()
        min_w = max(WINDOW_MIN_WIDTH_FALLBACK, int(sw * WINDOW_MIN_WIDTH_FRACTION))
        min_h = max(WINDOW_MIN_HEIGHT_FALLBACK, int(sh * WINDOW_MIN_HEIGHT_FRACTION))
        self.setMinimumSize(min_w, min_h)

        self._setup_ui()
        self._center_on_screen()

    def _setup_ui(self):
        """
        WHAT: Creates the tab widget, real tabs, status bar, and Excel monitor.
        CALLED BY: __init__()
        """
        # --- Excel infrastructure ---
        self._excel_handler = ExcelHandler()
        self._workbook_manager = WorkbookManager(self._excel_handler)
        self._excel_highlighter = ExcelHighlighter(self._excel_handler)
        self._excel_monitor = ExcelMonitor(self._excel_handler)

        # --- Tab widget ---
        self._tab_widget = QTabWidget()
        self._tab_widget.setDocumentMode(True)

        # Find tab — real panel (Sub-phase 1E)
        self._find_tab = FindTab()
        self._find_tab.set_workbook_manager(self._workbook_manager)
        self._find_tab.set_excel_highlighter(self._excel_highlighter)

        # Summary tab — real panel (Sub-phase 2B)
        self._summary_tab = SummaryTab()

        # Settings tab — real panel (Sub-phase 3A)
        self._settings_tab = SettingsTab(self._workbook_manager)

        self._tab_widget.addTab(self._find_tab, "Find")
        self._tab_widget.addTab(self._summary_tab, "Summary")
        self._tab_widget.addTab(self._settings_tab, "Settings")

        self.setCentralWidget(self._tab_widget)

        # --- Wire finalization signals ---
        # When FindTab finalizes/undoes/clears, refresh the Summary tab
        self._find_tab.finalization_changed.connect(self._on_finalization_changed)
        # When SummaryTab Undo is clicked, forward to FindTab
        self._summary_tab.undo_requested.connect(self._on_undo_requested)
        # When SummaryTab Mark Unmatched is clicked, highlight grey in Excel
        self._summary_tab.mark_unmatched_requested.connect(self._on_mark_unmatched)

        # --- Wire Excel signals ---
        self._settings_tab.excel_connected.connect(self._on_excel_connected)
        self._settings_tab.excel_disconnected.connect(self._on_excel_disconnected)

        # --- Excel monitor timer ---
        self._excel_monitor_timer = QTimer(self)
        self._excel_monitor_timer.timeout.connect(self._check_excel_connection)
        # Timer is started when Excel connects, stopped on disconnect

        # --- Session auto-save timer ---
        self._autosave_timer = QTimer(self)
        self._autosave_timer.timeout.connect(self._auto_save_session)
        self._autosave_timer.start(SESSION_AUTOSAVE_INTERVAL)

        # --- Status bar ---
        self._status_bar = QStatusBar()
        self.setStatusBar(self._status_bar)
        self._status_bar.showMessage("Ready")

    def _create_placeholder_tab(self, title: str, description: str) -> QWidget:
        """
        WHAT: Creates a temporary placeholder tab with a centered label.
              These will be replaced by real panel widgets in Sub-phases
              1D, 1E, 2B, and 3A.
        CALLED BY: _setup_ui()
        RETURNS: QWidget — the placeholder tab widget.
        """
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setAlignment(Qt.AlignCenter)

        title_label = QLabel(title)
        title_label.setStyleSheet(f"font-size: {scaled_size(20)}px; font-weight: bold; color: {COLOR_TEXT_SECONDARY};")
        title_label.setAlignment(Qt.AlignCenter)

        desc_label = QLabel(description)
        desc_label.setStyleSheet(f"font-size: {scaled_size(15)}px; color: {COLOR_TEXT_SECONDARY};")
        desc_label.setAlignment(Qt.AlignCenter)

        layout.addWidget(title_label)
        layout.addSpacing(10)
        layout.addWidget(desc_label)

        return tab

    def _on_finalization_changed(self):
        """
        WHAT:
            Refreshes the Summary tab after any finalization change
            (finalize, undo, or clear). Passes unmatched count so
            the Summary tab can show it in stats and enable/disable
            the Mark Unmatched button. Also triggers immediate session save.

        CALLED BY:
            → FindTab.finalization_changed signal
        """
        manager = self._find_tab.get_finalization_manager()
        all_items = self._find_tab.get_all_items()
        unmatched = manager.get_unmatched_items(all_items)
        self._summary_tab.refresh(
            manager.get_finalized_list(),
            manager.can_undo(),
            unmatched_count=len(unmatched),
            total_items=len(all_items),
        )

        # Immediate save on every finalization/undo (not just timer)
        self._save_session()

    def _on_undo_requested(self):
        """
        WHAT:
            Forwards the SummaryTab's Undo button click to FindTab,
            which performs the undo and emits finalization_changed
            to refresh the Summary tab.

        CALLED BY:
            → SummaryTab.undo_requested signal
        """
        self._find_tab.undo_last_finalization()

    def _on_mark_unmatched(self):
        """
        WHAT:
            Handles the Mark Unmatched button click from the Summary tab.
            Highlights all non-finalized Excel-sourced items in grey
            using the ExcelHighlighter.

        CALLED BY:
            → SummaryTab.mark_unmatched_requested signal

        EDGE CASES HANDLED:
            - Excel not connected → shows error message
            - No unmatched items → button should be disabled (but safe if called)
            - Some cells fail → reports count of successes
        """
        manager = self._find_tab.get_finalization_manager()
        all_items = self._find_tab.get_all_items()
        unmatched = manager.get_unmatched_items(all_items)

        if not unmatched:
            self._status_bar.showMessage("No unmatched items to highlight")
            return

        if not self._excel_highlighter:
            self._status_bar.showMessage("Excel highlighter not available")
            return

        grey_rgb = UNMATCHED_COLOR[0]  # (220, 220, 220)
        result = self._excel_highlighter.highlight_unmatched(unmatched, grey_rgb)

        if result["success"]:
            cells = result["data"]["cells_highlighted"]
            self._summary_tab.show_unmatched_result(cells)
            self._status_bar.showMessage(
                f"Marked {cells} unmatched cell{'s' if cells != 1 else ''} in grey"
            )
        else:
            from gui.dialogs import show_error
            show_error(
                self,
                "Mark Unmatched Failed",
                result["error"] or "Could not highlight unmatched cells in Excel.",
            )

    def _on_excel_connected(self):
        """
        WHAT:
            Starts the Excel monitor timer after a successful connection.
            Stores workbook paths for reopen functionality.

        CALLED BY: SettingsTab.excel_connected signal.
        """
        self._excel_monitor_timer.start(EXCEL_MONITOR_INTERVAL)
        self._excel_monitor.update_workbook_paths()
        self._status_bar.showMessage("Connected to Excel")

    def _on_excel_disconnected(self):
        """
        WHAT: Stops the Excel monitor timer after disconnection.
        CALLED BY: SettingsTab.excel_disconnected signal.
        """
        self._excel_monitor_timer.stop()
        self._excel_monitor.clear()
        self._status_bar.showMessage("Disconnected from Excel")

    def _check_excel_connection(self):
        """
        WHAT:
            Periodically checks if Excel is still running. If the
            connection dropped (user closed Excel), updates Settings
            tab, stops the monitor timer, and offers to reopen.

        CALLED BY:
            → _excel_monitor_timer QTimer (every EXCEL_MONITOR_INTERVAL ms)

        EDGE CASES HANDLED:
            - COM call raises unexpected exception → treated as disconnected
            - Monitor has no stored names (unsaved workbooks) → shows
              generic disconnect message instead of reopen dialog
        """
        try:
            still_alive = self._settings_tab.check_connection_alive()
        except Exception:
            # Any COM error during the check means connection is dead
            still_alive = False

        if not still_alive:
            self._excel_monitor_timer.stop()
            self._status_bar.showMessage("Excel connection lost")

            # Check if we have workbook paths to offer reopen
            last_names = self._excel_monitor.get_last_known_names()
            if last_names:
                from gui.dialogs import warn_excel_closed
                display_name = ", ".join(last_names[:3])
                if len(last_names) > 3:
                    display_name += f" (+{len(last_names) - 3} more)"
                wants_reopen = warn_excel_closed(self, display_name)

                if wants_reopen:
                    result = self._excel_monitor.reopen_workbooks()
                    if result["success"]:
                        self._status_bar.showMessage(
                            f"Reopening {result['data']['files_opened']} workbook(s) — "
                            f"click Connect in Settings when Excel finishes loading"
                        )
                    else:
                        from gui.dialogs import show_error
                        show_error(self, "Reopen Failed", result["error"])
            else:
                from gui.dialogs import show_info
                show_info(
                    self,
                    "Excel Disconnected",
                    "Excel appears to have been closed.\n\n"
                    "The connection has been lost. You can reconnect "
                    "from the Settings tab when Excel is running again."
                )

    # -------------------------------------------------------------------
    # Session Save / Restore
    # -------------------------------------------------------------------

    def try_restore_session(self):
        """
        WHAT:
            Checks for a saved session on disk and offers to restore it.
            Called by main.py AFTER the window is shown, so the restore
            dialog appears on top of the visible window.

        CALLED BY:
            → main.py → after window.show()

        EDGE CASES HANDLED:
            - No session file → does nothing
            - Corrupt session file → shows error, deletes file
            - User chooses Restore → loads session into FindTab
            - User chooses Discard → deletes session file
            - User chooses Start Fresh → keeps file for next time
        """
        if not SessionManager.has_saved_session():
            return

        summary = SessionManager.get_session_summary()
        if summary is None:
            return

        # Only offer restore if there is meaningful state
        if summary["item_count"] == 0 and summary["finalized_count"] == 0:
            SessionManager.delete_session()
            return

        from gui.dialogs import ask_session_restore, show_error
        choice = ask_session_restore(self, summary)

        if choice == "restore":
            result = SessionManager.load_session()
            if result["success"]:
                data = result["data"]
                self._find_tab.restore_session(
                    items=data["items"],
                    finalized_list=data["finalized"],
                    next_color_index=data["next_color_index"],
                    next_combo_number=data["next_combo_number"],
                    search_params=data["search_params"],
                )
                self._status_bar.showMessage(
                    f"Session restored — {summary['item_count']} items, "
                    f"{summary['finalized_count']} finalized"
                )
            else:
                show_error(self, "Restore Failed", result["error"])
                SessionManager.delete_session()

        elif choice == "discard":
            SessionManager.delete_session()
            self._status_bar.showMessage("Previous session discarded")

        # "cancel" (Start Fresh) — keep file, do nothing

    def _save_session(self):
        """
        WHAT:
            Saves the current application state to disk. Called by the
            auto-save timer and immediately after finalization/undo.

        CALLED BY:
            → _autosave_timer timeout
            → _on_finalization_changed()
            → closeEvent()
        """
        state = self._find_tab.get_session_state()

        # Only save if there is something to save
        if not state["items"] and not state["finalized_list"]:
            return

        SessionManager.save_session(
            items=state["items"],
            finalized_list=state["finalized_list"],
            next_color_index=state["next_color_index"],
            next_combo_number=state["next_combo_number"],
            search_params=state["search_params"],
        )

    def _auto_save_session(self):
        """
        WHAT: Timer callback for periodic auto-save. Silently saves state.
        CALLED BY: _autosave_timer timeout (every SESSION_AUTOSAVE_INTERVAL ms)
        """
        self._save_session()

    def _center_on_screen(self):
        """
        WHAT: Centers the window on the primary screen.
        CALLED BY: __init__()

        WHY:
            Users expect the app to appear centered, not in the top-left
            corner. Uses QScreen (not deprecated QDesktopWidget).
        """
        screen = self.screen()
        if screen is not None:
            screen_geometry = screen.availableGeometry()
            window_geometry = self.frameGeometry()
            center_point = screen_geometry.center()
            window_geometry.moveCenter(center_point)
            self.move(window_geometry.topLeft())

    def update_status(self, message: str):
        """
        WHAT: Updates the status bar message.
        CALLED BY: Various tabs and managers for status updates.

        PARAMETERS:
            message (str): The message to display in the status bar.
        """
        self._status_bar.showMessage(message)

    def closeEvent(self, event):
        """
        WHAT:
            Handles the window close event. If finalized combos exist,
            asks whether to save session before closing. Always saves
            session on "save", deletes on "discard", and disconnects
            Excel cleanly.

        CALLED BY:
            → Qt framework when user clicks X or Alt+F4

        EDGE CASES HANDLED:
            - No finalized combos → save session silently and close
            - User cancels → abort close
            - User saves → save session to disk, then close
            - User discards → delete session file, then close
        """
        manager = self._find_tab.get_finalization_manager()
        has_finalized = manager.get_finalized_count() > 0
        has_items = len(self._find_tab.get_all_items()) > 0

        # If there is meaningful state, ask the user what to do
        if has_finalized:
            from gui.dialogs import confirm_close_unsaved
            choice = confirm_close_unsaved(self)

            if choice == "cancel":
                event.ignore()
                return
            elif choice == "save":
                # Save session so it can be restored on next launch
                self._save_session()
            elif choice == "discard":
                # Delete session file — user doesn't want to resume
                SessionManager.delete_session()
        elif has_items:
            # Items loaded but nothing finalized — save silently
            # so user can resume if they close by accident
            self._save_session()

        # Stop timers
        self._autosave_timer.stop()
        self._excel_monitor_timer.stop()

        # Disconnect Excel cleanly
        if self._excel_handler.is_connected:
            try:
                self._excel_handler.disconnect()
            except Exception:
                pass  # Best effort — we're closing anyway

        event.accept()
