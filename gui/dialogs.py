"""
FILE: gui/dialogs.py

PURPOSE: All popup dialogs used throughout the application. Centralised
         here so every dialog has consistent styling and behaviour.

CONTAINS:
- ask_label()                — Label input dialog for finalization
- confirm_clear_all()        — Clear All confirmation
- confirm_close_unsaved()    — Close with unsaved changes
- ask_session_restore()      — Session restore on startup
- warn_no_solution()         — No solution possible
- warn_large_search_space()  — Large search space confirmation
- warn_excel_closed()        — Excel closed, offer to reopen
- show_error()               — Generic error message
- show_info()                — Generic info message

DEPENDS ON:
- gui/styles.py → color constants

USED BY:
- gui/find_tab.py → search warnings, label input, clear all
- gui/main_window.py → close confirmation
- gui/settings_tab.py → Excel reopen offer

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — all popup dialogs               | Sub-phase 2A finalization core   |
| 23-03-2026 | Added session restore dialog               | Phase 4C session persistence     |
"""

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QMessageBox, QInputDialog, QWidget,
)
from PyQt5.QtCore import Qt

# Group 3: This project's modules
from gui.styles import COLOR_ACCENT


def ask_label(parent: QWidget, combo_number: int) -> dict:
    """
    WHAT:
        Shows a text input dialog asking the user for an optional label
        for the finalized combination. Returns the label text or None
        if the user cancelled.

    WHY ADDED:
        Labels like "Dec Payment" or "TDS Adjustment" help users
        identify finalized groups later. The label is optional —
        pressing OK with empty text is valid.

    CALLED BY:
        → gui/find_tab.py → after user clicks Finalize

    PARAMETERS:
        parent (QWidget): Parent widget for the dialog.
        combo_number (int): The finalization number (for the dialog title).

    RETURNS:
        dict: {
            "accepted": bool,  — True if user clicked OK
            "label": str       — The entered label (may be empty)
        }
    """
    text, ok = QInputDialog.getText(
        parent,
        f"Label — Combination #{combo_number}",
        "Enter an optional label for this combination:\n"
        "(e.g., \"Dec Payment\", \"TDS Adj\", or leave blank)",
    )

    return {
        "accepted": ok,
        "label": text.strip() if ok else "",
    }


def confirm_clear_all(parent: QWidget, has_finalized: bool = False) -> bool:
    """
    WHAT:
        Asks the user to confirm clearing all numbers and results.
        If finalized combinations exist, warns more strongly.

    CALLED BY:
        → gui/find_tab.py → Clear All button

    PARAMETERS:
        parent (QWidget): Parent widget.
        has_finalized (bool): True if there are finalized combinations.

    RETURNS:
        bool: True if user confirmed, False if cancelled.
    """
    if has_finalized:
        message = (
            "This will remove ALL loaded numbers, search results, "
            "AND finalized combinations.\n\n"
            "Finalized combinations cannot be recovered after clearing.\n\n"
            "Are you sure?"
        )
    else:
        message = (
            "This will remove all loaded numbers and search results.\n\n"
            "Are you sure?"
        )

    reply = QMessageBox.question(
        parent,
        "Clear All",
        message,
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No,
    )
    return reply == QMessageBox.Yes


def confirm_close_unsaved(parent: QWidget) -> str:
    """
    WHAT:
        Asks the user what to do when closing with unsaved work.
        Returns "save", "discard", or "cancel".

    CALLED BY:
        → gui/main_window.py → closeEvent (Phase 4)

    RETURNS:
        str: "save", "discard", or "cancel"
    """
    msg = QMessageBox(parent)
    msg.setWindowTitle("Unsaved Work")
    msg.setText("You have unsaved finalized combinations.")
    msg.setInformativeText("Do you want to save your session before closing?")
    msg.setStandardButtons(
        QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
    )
    msg.setDefaultButton(QMessageBox.Save)

    result = msg.exec_()

    if result == QMessageBox.Save:
        return "save"
    elif result == QMessageBox.Discard:
        return "discard"
    else:
        return "cancel"


def ask_session_restore(parent: QWidget, summary: dict) -> str:
    """
    WHAT:
        Asks the user whether to restore a previous session found on disk.
        Shows the session timestamp, number count, and finalization count.
        Returns "restore", "discard", or "cancel".

    WHY ADDED:
        When the app starts and finds a saved session, the user should
        choose whether to resume where they left off or start fresh.

    CALLED BY:
        → gui/main_window.py → on startup when a session file exists

    PARAMETERS:
        parent (QWidget): Parent widget.
        summary (dict): {"timestamp": str, "item_count": int,
                         "finalized_count": int}

    RETURNS:
        str: "restore" to load session, "discard" to delete and start fresh,
             "cancel" to do nothing (start fresh, keep file for next time).
    """
    timestamp = summary.get("timestamp", "Unknown")
    item_count = summary.get("item_count", 0)
    finalized_count = summary.get("finalized_count", 0)

    # Format timestamp for display
    try:
        from datetime import datetime
        dt = datetime.fromisoformat(timestamp)
        display_time = dt.strftime("%d-%b-%Y %I:%M %p")
    except Exception:
        display_time = timestamp

    msg = QMessageBox(parent)
    msg.setWindowTitle("Restore Previous Session")
    msg.setText("A previous session was found.")
    msg.setInformativeText(
        f"Saved: {display_time}\n"
        f"Numbers loaded: {item_count}\n"
        f"Combinations finalized: {finalized_count}\n\n"
        f"Would you like to restore this session?"
    )

    restore_btn = msg.addButton("Restore", QMessageBox.AcceptRole)
    discard_btn = msg.addButton("Discard", QMessageBox.DestructiveRole)
    msg.addButton("Start Fresh", QMessageBox.RejectRole)
    msg.setDefaultButton(restore_btn)

    msg.exec_()
    clicked = msg.clickedButton()

    if clicked == restore_btn:
        return "restore"
    elif clicked == discard_btn:
        return "discard"
    else:
        return "cancel"


def warn_no_solution(parent: QWidget):
    """
    WHAT:
        Shows an informational dialog when smart bounds determine
        that no combination can match the target.

    CALLED BY:
        → gui/find_tab.py → after compute_bounds returns no_solution=True
    """
    QMessageBox.information(
        parent,
        "No Solution",
        "No combination of the loaded numbers can match the target "
        "with the current parameters.\n\n"
        "Try adjusting the target, buffer, or size range.",
    )


def warn_large_search_space(parent: QWidget, estimated: int) -> bool:
    """
    WHAT:
        Shows a warning when the estimated search space is very large.
        Returns True if the user wants to proceed anyway.

    CALLED BY:
        → gui/find_tab.py → before starting a large search

    PARAMETERS:
        parent (QWidget): Parent widget.
        estimated (int): Estimated number of combinations to check.

    RETURNS:
        bool: True if user clicked "Proceed", False if cancelled.
    """
    reply = QMessageBox.question(
        parent,
        "Large Search Space",
        f"This search will check approximately {estimated:,} combinations.\n\n"
        f"This may take a very long time. You can click Stop at any "
        f"time to cancel.\n\n"
        f"Proceed anyway?",
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No,
    )
    return reply == QMessageBox.Yes


def warn_excel_closed(parent: QWidget, workbook_name: str) -> bool:
    """
    WHAT:
        Shows a dialog when Excel has been closed, offering to reopen
        the last known workbook.

    CALLED BY:
        → gui/main_window.py → Excel monitor detects closure (Phase 3)

    PARAMETERS:
        parent (QWidget): Parent widget.
        workbook_name (str): Name of the workbook that was open.

    RETURNS:
        bool: True if user wants to reopen, False to ignore.
    """
    reply = QMessageBox.question(
        parent,
        "Excel Closed",
        f"Excel appears to have been closed.\n\n"
        f"Would you like to reopen '{workbook_name}'?",
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No,
    )
    return reply == QMessageBox.Yes


def show_error(parent: QWidget, title: str, message: str):
    """
    WHAT: Shows a generic error message dialog.
    CALLED BY: Various modules for error display.
    """
    QMessageBox.critical(parent, title, message)


def show_info(parent: QWidget, title: str, message: str):
    """
    WHAT: Shows a generic information message dialog.
    CALLED BY: Various modules for info display.
    """
    QMessageBox.information(parent, title, message)


def confirm_multi_column(parent: QWidget) -> bool:
    """
    WHAT:
        Shows a dialog when a multi-column selection is detected in Excel.
        Gives the user a choice to proceed with all columns or cancel
        and re-select a single column.

    CALLED BY:
        → gui/find_tab.py → after grab detects multi-column warning

    RETURNS:
        bool: True if user wants to proceed, False to cancel.
    """
    reply = QMessageBox.question(
        parent,
        "Multi-Column Selection",
        "The Excel selection spans multiple columns.\n\n"
        "CombiMatch works best with a single column of numbers.\n\n"
        "• Click 'Yes' to proceed with all columns\n"
        "• Click 'No' to cancel and re-select in Excel",
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No,
    )
    return reply == QMessageBox.Yes


def confirm_action(parent: QWidget, title: str, message: str) -> bool:
    """
    WHAT: Shows a generic Yes/No confirmation dialog.
    CALLED BY: Various modules for action confirmation.

    RETURNS:
        bool: True if user clicked Yes.
    """
    reply = QMessageBox.question(
        parent,
        title,
        message,
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No,
    )
    return reply == QMessageBox.Yes
