"""
FILE: readers/excel_monitor.py

PURPOSE: Monitors the Excel connection and provides reopen functionality.
         Tracks the last known workbook paths so the app can offer to
         reopen them if Excel is closed.

CONTAINS:
- ExcelMonitor — Tracks workbook paths, detects closure, offers reopen

DEPENDS ON:
- readers/excel_reader.py → ExcelHandler

USED BY:
- gui/main_window.py → called by the Excel monitor timer

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — Excel connection monitor        | Sub-phase 3B Excel monitoring    |
| 22-03-2026 | Store names even for unsaved workbooks     | Reopen dialog wasn't appearing   |
"""

# Group 1: Python standard library
import os
import subprocess
from typing import List, Optional

# Group 3: This project's modules
from readers.excel_reader import ExcelHandler, PYWIN32_AVAILABLE


class ExcelMonitor:
    """
    WHAT:
        Tracks open workbook file paths and provides reopen functionality.
        When Excel closes unexpectedly, the monitor can offer to reopen
        the last known workbooks.

    WHY ADDED:
        Users sometimes accidentally close Excel during a reconciliation
        session. Without monitoring, they lose their connection and have
        to manually reopen files and reconnect. This provides a smoother
        recovery path.

    CALLED BY:
        → gui/main_window.py → timer-based periodic check

    ASSUMPTIONS:
        - Workbook FullName paths are available via COM while connected.
        - Reopening uses os.startfile() which launches the default app
          (Excel) for .xlsx files.
        *** ASSUMPTION: After reopening, the user must click "Connect"
            in Settings again. Auto-reconnect is not attempted because
            Excel needs time to fully load the workbook. ***
    """

    def __init__(self, excel_handler: ExcelHandler):
        """
        WHAT: Creates the monitor with a reference to the Excel handler.

        PARAMETERS:
            excel_handler (ExcelHandler): The shared handler for COM access.
        """
        self._handler = excel_handler
        self._last_known_paths: List[str] = []
        self._last_known_names: List[str] = []

    # -------------------------------------------------------------------
    # Public Methods
    # -------------------------------------------------------------------

    def update_workbook_paths(self):
        """
        WHAT:
            Reads and stores the names and file paths of all currently open
            workbooks. Should be called after connecting and after each
            refresh, so the monitor has current data for the reopen dialog.

        CALLED BY:
            → gui/main_window.py → after Excel connects or refreshes

        EDGE CASES HANDLED:
            - Unsaved workbooks (FullName is just the name, not a real path)
              → name is stored for the dialog, but path is skipped so
              reopen won't try to open a nonexistent file
            - COM error reading a single workbook → skips it, continues
        """
        if not PYWIN32_AVAILABLE or not self._handler.is_connected:
            return

        try:
            app = self._handler._excel_app
            paths = []
            names = []
            for i in range(1, app.Workbooks.Count + 1):
                wb = app.Workbooks(i)
                try:
                    name = wb.Name
                    full_path = wb.FullName
                    # Always store the name (for display in reopen dialog)
                    names.append(name)
                    # Only store the path if it's a real saved file
                    if full_path and os.path.isfile(full_path):
                        paths.append(full_path)
                except Exception:
                    pass
            self._last_known_paths = paths
            self._last_known_names = names
        except Exception:
            pass  # Connection may have dropped

    def get_last_known_names(self) -> List[str]:
        """
        WHAT: Returns the names of the last known open workbooks.
        CALLED BY: gui/main_window.py → for the reopen dialog message.
        """
        return list(self._last_known_names)

    def get_last_known_paths(self) -> List[str]:
        """
        WHAT: Returns the full paths of the last known open workbooks.
        CALLED BY: gui/main_window.py → for reopening files.
        """
        return list(self._last_known_paths)

    def reopen_workbooks(self) -> dict:
        """
        WHAT:
            Attempts to reopen the last known workbooks by launching
            them with the default application (Excel). Does NOT
            reconnect — the user must click Connect in Settings after
            Excel finishes loading.

        CALLED BY:
            → gui/main_window.py → when user clicks "Yes" on reopen dialog

        RETURNS:
            dict: {"success": bool, "error": str or None, "data": {
                "files_opened": int,
                "errors": List[str]
            }}
        """
        if not self._last_known_paths:
            return {
                "success": False,
                "error": "No workbook paths available to reopen",
                "data": None,
            }

        files_opened = 0
        errors = []

        for path in self._last_known_paths:
            if not os.path.isfile(path):
                errors.append(f"File not found: {path}")
                continue

            try:
                os.startfile(path)
                files_opened += 1
            except Exception as e:
                errors.append(f"Could not open {os.path.basename(path)}: {str(e)}")

        return {
            "success": files_opened > 0,
            "error": "; ".join(errors) if errors else None,
            "data": {
                "files_opened": files_opened,
                "errors": errors,
            },
        }

    def clear(self):
        """
        WHAT: Clears stored workbook paths. Called on disconnect.
        CALLED BY: gui/main_window.py → on Excel disconnect.
        """
        self._last_known_paths = []
        self._last_known_names = []
