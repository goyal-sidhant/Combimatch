"""
FILE: main.py

PURPOSE: Entry point for CombiMatch. Does three things only:
         1. Single instance check (prevents duplicate launches)
         2. Creates QApplication with stylesheet
         3. Creates and shows MainWindow

CONTAINS:
- main() — Application entry point

DEPENDS ON:
- utils/single_instance.py → SingleInstanceGuard
- gui/styles.py → get_stylesheet()
- gui/main_window.py → MainWindow

USED BY:
- The user runs this file: python main.py

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — application entry point         | Sub-phase 1C app shell           |
"""

# Group 1: Python standard library
import sys

# Group 2: Third-party libraries
from PyQt5.QtWidgets import QApplication, QMessageBox

# Group 3: This project's modules
from utils.single_instance import SingleInstanceGuard
from gui.styles import get_stylesheet, compute_font_scale
from gui.main_window import MainWindow


def main():
    """
    WHAT:
        Application entry point. Checks for duplicate instances,
        creates the Qt application, applies styling, and launches
        the main window.

    WHY ADDED:
        Every app needs an entry point. This one is deliberately
        minimal — all logic lives in the appropriate modules.

    CALLED BY:
        → __name__ == "__main__" block at the bottom of this file

    CALLS:
        → utils/single_instance.py → SingleInstanceGuard
        → gui/styles.py → get_stylesheet()
        → gui/main_window.py → MainWindow

    EDGE CASES HANDLED:
        - Another instance already running → shows error dialog, exits
        - Normal startup → single instance guard stays active until exit

    ASSUMPTIONS:
        - Python 3.x with PyQt5 installed.
        - Windows OS (for single instance mutex).
    """
    # --- Single instance check ---
    guard = SingleInstanceGuard()

    if guard.is_running:
        # Need a temporary QApplication to show the error dialog
        temp_app = QApplication(sys.argv)
        QMessageBox.warning(
            None,
            "CombiMatch",
            "CombiMatch is already running.\n\n"
            "Only one instance can run at a time on this computer.\n"
            "Check the taskbar for the existing window.",
        )
        guard.release()
        sys.exit(1)

    # --- Create application ---
    app = QApplication(sys.argv)
    app.setApplicationName("CombiMatch")
    compute_font_scale()
    app.setStyleSheet(get_stylesheet())

    # --- Create and show main window ---
    window = MainWindow()
    window.show()

    # --- Restore previous session (if any) ---
    # Called after show() so the dialog appears on top of the visible window
    window.try_restore_session()

    # --- Run event loop ---
    exit_code = app.exec_()

    # --- Cleanup ---
    guard.release()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
