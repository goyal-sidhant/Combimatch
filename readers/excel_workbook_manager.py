"""
FILE: readers/excel_workbook_manager.py

PURPOSE: Manages workbook/sheet selections for multi-source grab (N8).
         Tracks which workbooks and sheets are checked, orchestrates
         reading from all checked sources, and supports accumulative grab.

CONTAINS:
- WorkbookManager — Orchestrates multi-workbook/sheet reading

DEPENDS ON:
- readers/excel_reader.py → ExcelHandler

USED BY:
- gui/settings_tab.py → UI for workbook/sheet selection
- gui/find_tab.py → grab_from_checked() for reading numbers

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — multi-workbook grab manager    | Sub-phase 3A Excel integration   |
"""

# Group 1: Python standard library
from typing import Dict, List, Tuple

# Group 3: This project's modules
from readers.excel_reader import ExcelHandler


class WorkbookManager:
    """
    WHAT:
        Tracks which workbooks and sheets the user has checked in the
        Settings tab. Provides grab_from_checked() which loops through
        all checked sources, reads selections from each, and returns
        a combined list of NumberItems.

    WHY ADDED:
        Real reconciliations span multiple workbooks (e.g., Apr-Jun in
        one file, Jul-Sep in another). The user checks the relevant
        sheets, makes cell selections on each, then clicks "Grab from
        Excel" once to load everything.

    CALLED BY:
        → gui/settings_tab.py → selection UI
        → gui/find_tab.py → grab button handler

    ASSUMPTIONS:
        - Each workbook/sheet pair has its own selection in Excel.
          Excel maintains per-sheet selections automatically.
        - Checked state is stored in memory only (not persisted).
          Session persistence (Phase 4) will save/restore this.
        *** ASSUMPTION: When refresh_workbooks() is called, any workbooks
            that were closed in Excel are removed from the selections
            dict. New workbooks are added with all sheets unchecked. ***
    """

    def __init__(self, excel_handler: ExcelHandler):
        """
        WHAT: Creates a WorkbookManager with a reference to the ExcelHandler.

        PARAMETERS:
            excel_handler (ExcelHandler): The shared COM handler.
        """
        self._handler = excel_handler
        # Structure: {workbook_name: {sheet_name: is_checked}}
        self._selections: Dict[str, Dict[str, bool]] = {}

    @property
    def excel_handler(self) -> ExcelHandler:
        """
        WHAT: Returns the underlying ExcelHandler.
        CALLED BY: gui/settings_tab.py → for connect/disconnect.
        """
        return self._handler

    # -------------------------------------------------------------------
    # Workbook/Sheet Management
    # -------------------------------------------------------------------

    def refresh_workbooks(self) -> dict:
        """
        WHAT:
            Refreshes the list of open workbooks and their sheets from
            the live Excel instance. Preserves check state for workbooks
            and sheets that are still open. Removes closed ones. Adds
            new ones as unchecked.

        CALLED BY:
            → gui/settings_tab.py → Refresh button, after connect

        RETURNS:
            dict: {"success": bool, "error": str or None,
                   "data": {"selections": dict} or None}
        """
        result = self._handler.get_workbooks()
        if not result["success"]:
            return result

        workbook_names = result["data"]["workbooks"]

        new_selections: Dict[str, Dict[str, bool]] = {}
        for wb_name in workbook_names:
            sheets_result = self._handler.get_sheets(wb_name)
            if not sheets_result["success"]:
                continue

            sheet_names = sheets_result["data"]["sheets"]
            old_wb = self._selections.get(wb_name, {})

            new_selections[wb_name] = {}
            for sheet_name in sheet_names:
                # Preserve existing check state; default unchecked
                new_selections[wb_name][sheet_name] = old_wb.get(sheet_name, False)

        self._selections = new_selections

        return {
            "success": True,
            "error": None,
            "data": {"selections": self._selections},
        }

    def set_sheet_checked(self, workbook_name: str, sheet_name: str,
                          checked: bool):
        """
        WHAT: Sets the checked state of a specific sheet.

        CALLED BY:
            → gui/settings_tab.py → when user toggles a sheet checkbox

        PARAMETERS:
            workbook_name (str): The workbook containing the sheet.
            sheet_name (str): The sheet to check/uncheck.
            checked (bool): True to check, False to uncheck.
        """
        if workbook_name in self._selections:
            if sheet_name in self._selections[workbook_name]:
                self._selections[workbook_name][sheet_name] = checked

    def set_workbook_checked(self, workbook_name: str, checked: bool):
        """
        WHAT: Checks or unchecks all sheets in a workbook.

        CALLED BY:
            → gui/settings_tab.py → when user toggles a workbook checkbox

        PARAMETERS:
            workbook_name (str): The workbook to check/uncheck all sheets.
            checked (bool): True to check all, False to uncheck all.
        """
        if workbook_name in self._selections:
            for sheet_name in self._selections[workbook_name]:
                self._selections[workbook_name][sheet_name] = checked

    def get_checked_sources(self) -> List[Tuple[str, str]]:
        """
        WHAT: Returns list of (workbook_name, sheet_name) pairs that are checked.

        CALLED BY:
            → grab_from_checked()
            → gui/settings_tab.py → for display

        RETURNS:
            List[Tuple[str, str]]: Checked workbook/sheet pairs.
        """
        sources = []
        for wb_name, sheets in self._selections.items():
            for sheet_name, is_checked in sheets.items():
                if is_checked:
                    sources.append((wb_name, sheet_name))
        return sources

    def get_selections(self) -> Dict[str, Dict[str, bool]]:
        """
        WHAT: Returns the full selections dictionary.
        CALLED BY: gui/settings_tab.py → for building the tree display.

        RETURNS:
            dict: {workbook_name: {sheet_name: is_checked}}
        """
        return self._selections

    def has_any_checked(self) -> bool:
        """
        WHAT: Returns True if at least one sheet is checked.
        CALLED BY: gui/find_tab.py → to enable/disable Grab button.
        """
        return len(self.get_checked_sources()) > 0

    # -------------------------------------------------------------------
    # Grab from Checked Sources
    # -------------------------------------------------------------------

    def grab_from_checked(self, start_index: int = 0) -> dict:
        """
        WHAT:
            Reads numbers from all checked workbook/sheet pairs. For each
            checked source, activates the sheet and reads the current
            Excel selection on that sheet. Combines all items into one list.

        WHY ADDED:
            This is the core of N8 (multi-workbook grab). Instead of
            reading from one sheet at a time, the user checks multiple
            sheets and clicks Grab once.

        CALLED BY:
            → gui/find_tab.py → on Grab from Excel button click

        PARAMETERS:
            start_index (int): Index to start numbering items from.
                               Enables accumulative grab — if user already
                               has 50 items loaded, start_index=50 so new
                               items get indices 50, 51, 52, ...

        EDGE CASES HANDLED:
            - No sheets checked → clear error message pointing to Settings
            - Some sheets fail → continues reading others, collects errors
            - All sheets fail → returns success=False
            - Empty selection on a sheet → skipped, counted as error

        RETURNS:
            dict: {"success": bool, "error": str or None,
                   "data": {"items": List[NumberItem],
                            "errors": List[str],
                            "warnings": List[str]} or None}
        """
        sources = self.get_checked_sources()
        if not sources:
            return {
                "success": False,
                "error": (
                    "No workbook/sheet selected.\n\n"
                    "Go to the Settings tab and check at least one sheet "
                    "to grab from."
                ),
                "data": None,
            }

        all_items = []
        all_errors = []
        all_warnings = []
        current_index = start_index

        for wb_name, sheet_name in sources:
            result = self._handler.read_selection(
                wb_name, sheet_name, current_index,
            )

            if result["success"] and result["data"]:
                items = result["data"]["items"]
                all_items.extend(items)
                all_errors.extend(result["data"].get("errors", []))
                all_warnings.extend(result["data"].get("warnings", []))
                current_index += len(items)
            elif not result["success"]:
                all_errors.append(
                    f"{wb_name} / {sheet_name}: {result['error']}"
                )

        if not all_items:
            error_detail = ""
            if all_errors:
                error_detail = "\n\nDetails:\n" + "\n".join(
                    f"  - {e}" for e in all_errors
                )
            return {
                "success": False,
                "error": f"No numeric values found in any checked sheet.{error_detail}",
                "data": None,
            }

        return {
            "success": True,
            "error": None,
            "data": {
                "items": all_items,
                "errors": all_errors,
                "warnings": all_warnings,
            },
        }

    def clear(self):
        """
        WHAT: Resets all selections (used on disconnect).
        CALLED BY: gui/settings_tab.py → on Disconnect.
        """
        self._selections.clear()
