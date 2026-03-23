"""
FILE: readers/excel_reader.py

PURPOSE: Handles all Excel COM automation for reading data. Connects to
         a running Excel instance, reads selected cells, and returns
         NumberItem objects with SourceTag metadata.

CONTAINS:
- PYWIN32_AVAILABLE          — Flag: True if pywin32 is installed
- ExcelHandler               — COM wrapper for reading from running Excel

DEPENDS ON:
- config/constants.py → ROUNDING_PRECISION, THOUSAND_SEPARATOR
- models/number_item.py → NumberItem
- models/source_tag.py → SourceTag

USED BY:
- readers/excel_workbook_manager.py → uses ExcelHandler for reading
- gui/settings_tab.py → connect/disconnect, get workbooks/sheets
- gui/find_tab.py → grab numbers from Excel (via WorkbookManager)

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — Excel COM reader               | Sub-phase 3A Excel integration   |
| 22-03-2026 | Skip zero values silently                  | Zeros don't contribute to sums   |
"""

# Group 1: Python standard library
from typing import List, Optional, Tuple

# Group 3: This project's modules
from config.constants import ROUNDING_PRECISION, THOUSAND_SEPARATOR
from models.number_item import NumberItem
from models.source_tag import SourceTag


# Try to import pywin32 — if not available, Excel features are disabled.
# The app still works for manual input without pywin32.
try:
    import pythoncom
    import win32com.client
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False


class ExcelHandler:
    """
    WHAT:
        Low-level COM wrapper for interacting with a running Excel instance.
        Provides connect/disconnect, workbook/sheet listing, sheet activation,
        and cell selection reading with SourceTag metadata.

    WHY ADDED:
        CombiMatch reads numbers from live Excel selections (not saved files).
        COM automation via pywin32 is the only way to interact with a running
        Excel instance on Windows.

    CALLED BY:
        → readers/excel_workbook_manager.py → WorkbookManager uses this
        → gui/settings_tab.py → connect/disconnect buttons

    ASSUMPTIONS:
        - Windows only. COM automation does not work on Mac/Linux.
        - pywin32 must be installed. If missing, PYWIN32_AVAILABLE is False
          and all methods return graceful errors.
        - pythoncom.CoInitialize() is called in connect(). All COM calls
          must happen on the same thread that called CoInitialize().
          Currently all COM work is on the GUI (main) thread.
        *** ASSUMPTION: Excel maintains per-sheet selections. When we
            activate a sheet and read Selection, we get the selection
            the user made on that sheet, even if they switched away.
            This is standard Excel behaviour and is relied upon for
            multi-workbook grab (N8). ***
    """

    def __init__(self):
        """
        WHAT: Initialises the handler with no active connection.
        """
        self._excel_app = None
        self._connected = False

    # -------------------------------------------------------------------
    # Connection Management
    # -------------------------------------------------------------------

    def connect(self) -> dict:
        """
        WHAT:
            Connects to the currently running Excel instance via COM.
            Does NOT start a new Excel — requires Excel to already be open.

        WHY ADDED:
            The entry point for all Excel interaction. Must be called before
            any workbook/sheet/selection reading.

        CALLED BY:
            → gui/settings_tab.py → Connect button

        CALLS:
            → pythoncom.CoInitialize()
            → win32com.client.GetActiveObject("Excel.Application")

        EDGE CASES HANDLED:
            - pywin32 not installed → error message with install instructions
            - Excel not running → clear error message
            - Any other COM error → generic error with details

        RETURNS:
            dict: {"success": bool, "error": str or None, "data": None}
        """
        if not PYWIN32_AVAILABLE:
            return {
                "success": False,
                "error": "pywin32 is not installed. Install it with: pip install pywin32",
                "data": None,
            }

        try:
            pythoncom.CoInitialize()
            self._excel_app = win32com.client.GetActiveObject("Excel.Application")
            self._connected = True
            return {"success": True, "error": None, "data": None}

        except pythoncom.com_error:
            self._excel_app = None
            self._connected = False
            return {
                "success": False,
                "error": "Excel is not running. Please open Excel first.",
                "data": None,
            }
        except Exception as e:
            self._excel_app = None
            self._connected = False
            return {
                "success": False,
                "error": f"Could not connect to Excel: {str(e)}",
                "data": None,
            }

    def disconnect(self):
        """
        WHAT: Releases the COM connection to Excel.
        CALLED BY: gui/settings_tab.py → Disconnect button, app close.

        WHY:
            COM objects hold references. Releasing them allows Excel to
            close cleanly if the user wants to.
        """
        self._excel_app = None
        self._connected = False

    @property
    def is_connected(self) -> bool:
        """
        WHAT:
            Checks if the connection to Excel is still alive by attempting
            multiple COM calls. Returns False if Excel was closed.

        CALLED BY:
            → gui/settings_tab.py → status display
            → readers/excel_workbook_manager.py → before operations

        EDGE CASES HANDLED:
            - Excel closed → Workbooks.Count raises com_error → False
            - COM proxy stale but Count succeeds → Name check catches it
            - Any unexpected exception → treated as disconnected

        RETURNS:
            bool: True if connected and Excel is responsive.
        """
        if not self._connected or self._excel_app is None:
            return False
        try:
            # Two-step liveness check: both must succeed.
            # Workbooks.Count can sometimes succeed briefly after Excel
            # closes (COM proxy lag). Name is a tighter check — it
            # requires the application object to be fully responsive.
            _ = self._excel_app.Name
            _ = self._excel_app.Workbooks.Count
            return True
        except Exception:
            self._connected = False
            self._excel_app = None
            return False

    # -------------------------------------------------------------------
    # Workbook and Sheet Listing
    # -------------------------------------------------------------------

    def get_workbooks(self) -> dict:
        """
        WHAT: Returns a list of names of all currently open workbooks.

        CALLED BY:
            → readers/excel_workbook_manager.py → refresh_workbooks()
            → gui/settings_tab.py → populate workbook list

        RETURNS:
            dict: {"success": bool, "error": str or None,
                   "data": {"workbooks": List[str]} or None}
        """
        if not self.is_connected:
            return {"success": False, "error": "Not connected to Excel", "data": None}

        try:
            workbooks = []
            for i in range(1, self._excel_app.Workbooks.Count + 1):
                wb = self._excel_app.Workbooks(i)
                workbooks.append(wb.Name)
            return {"success": True, "error": None, "data": {"workbooks": workbooks}}
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not read workbooks: {str(e)}",
                "data": None,
            }

    def get_sheets(self, workbook_name: str) -> dict:
        """
        WHAT: Returns a list of sheet names in the specified workbook.

        CALLED BY:
            → readers/excel_workbook_manager.py → refresh_workbooks()
            → gui/settings_tab.py → populate sheet list

        PARAMETERS:
            workbook_name (str): Name of the workbook (e.g., "Apr-Jun.xlsx").

        RETURNS:
            dict: {"success": bool, "error": str or None,
                   "data": {"sheets": List[str]} or None}
        """
        if not self.is_connected:
            return {"success": False, "error": "Not connected to Excel", "data": None}

        try:
            wb = self._excel_app.Workbooks(workbook_name)
            sheets = []
            for i in range(1, wb.Sheets.Count + 1):
                sheets.append(wb.Sheets(i).Name)
            return {"success": True, "error": None, "data": {"sheets": sheets}}
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not read sheets from '{workbook_name}': {str(e)}",
                "data": None,
            }

    def activate_sheet(self, workbook_name: str, sheet_name: str) -> dict:
        """
        WHAT: Activates a specific sheet in a specific workbook in Excel.

        CALLED BY:
            → read_selection() → before reading the selection

        PARAMETERS:
            workbook_name (str): Workbook name.
            sheet_name (str): Sheet name within the workbook.

        RETURNS:
            dict: {"success": bool, "error": str or None, "data": None}
        """
        if not self.is_connected:
            return {"success": False, "error": "Not connected to Excel", "data": None}

        try:
            wb = self._excel_app.Workbooks(workbook_name)
            ws = wb.Sheets(sheet_name)
            ws.Activate()
            return {"success": True, "error": None, "data": None}
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not activate '{sheet_name}' in '{workbook_name}': {str(e)}",
                "data": None,
            }

    # -------------------------------------------------------------------
    # Cell Selection Reading
    # -------------------------------------------------------------------

    def read_selection(self, workbook_name: str, sheet_name: str,
                       start_index: int = 0) -> dict:
        """
        WHAT:
            Reads the currently selected cells from a specific workbook/sheet.
            Returns NumberItem objects tagged with SourceTag metadata.
            Handles filtered cells (only reads visible), single-cell edge
            case, and bulk-read with cell-by-cell fallback.

        WHY ADDED:
            Core reading functionality. Production-hardened after 5 commits
            in the original tool. Single-cell bypass, SpecialCells(12) for
            filtered data, and bulk-read fallback are all critical patterns.

        CALLED BY:
            → readers/excel_workbook_manager.py → grab_from_checked()

        PARAMETERS:
            workbook_name (str): Which workbook to read from.
            sheet_name (str): Which sheet to read from.
            start_index (int): Index to start numbering items from
                               (for accumulative grab across sheets).

        EDGE CASES HANDLED:
            - No selection → error message
            - Single cell → bypasses SpecialCells (would hang)
            - Filtered data → uses SpecialCells(12) for visible cells only
            - Multi-column selection → warning but still reads all columns
            - Bulk read fails → falls back to cell-by-cell iteration
            - String values with commas → strips thousand separators
            - Non-numeric cells → skipped, counted as errors

        RETURNS:
            dict: {"success": bool, "error": str or None,
                   "data": {"items": List[NumberItem],
                            "errors": List[str],
                            "warnings": List[str]} or None}
        """
        if not self.is_connected:
            return {"success": False, "error": "Not connected to Excel", "data": None}

        try:
            # Activate the target sheet first
            activate_result = self.activate_sheet(workbook_name, sheet_name)
            if not activate_result["success"]:
                return activate_result

            selection = self._excel_app.Selection
            if selection is None:
                return {
                    "success": False,
                    "error": f"No cells selected on '{sheet_name}' in '{workbook_name}'",
                    "data": None,
                }

            cell_count = selection.Count
            if cell_count == 0:
                return {
                    "success": False,
                    "error": "No cells selected in Excel",
                    "data": None,
                }

            items: List[NumberItem] = []
            warnings: List[str] = []
            errors: List[str] = []

            # --- Single cell: bypass SpecialCells to avoid hang ---
            # SpecialCells(12) hangs when called on a single cell.
            # This workaround was hardened over 5 commits in the original.
            if cell_count == 1:
                self._read_single_cell(
                    selection, workbook_name, sheet_name,
                    start_index, items, errors,
                )
                return {
                    "success": True,
                    "error": None,
                    "data": {"items": items, "errors": errors, "warnings": warnings},
                }

            # --- Multiple cells: use SpecialCells(12) for visible only ---
            try:
                visible_cells = selection.SpecialCells(12)  # xlCellTypeVisible
            except Exception:
                # SpecialCells failed — fall back to raw selection
                visible_cells = selection
                warnings.append(
                    "Could not filter for visible cells — reading all selected cells"
                )

            # Warn about multi-column (still reads everything)
            if selection.Columns.Count > 1:
                warnings.append(
                    "Multi-column selection detected — reading all columns"
                )

            # --- Read each area (filtered ranges create multiple areas) ---
            current_index = start_index
            for area_idx in range(1, visible_cells.Areas.Count + 1):
                area = visible_cells.Areas(area_idx)
                area_items, area_errors = self._read_area(
                    area, workbook_name, sheet_name, current_index,
                )
                items.extend(area_items)
                errors.extend(area_errors)
                current_index += len(area_items)

            if not items and not errors:
                return {
                    "success": False,
                    "error": "No numeric values found in selection",
                    "data": None,
                }

            return {
                "success": True,
                "error": None,
                "data": {"items": items, "errors": errors, "warnings": warnings},
            }

        except Exception as e:
            return {
                "success": False,
                "error": f"Error reading from Excel: {str(e)}",
                "data": None,
            }

    # -------------------------------------------------------------------
    # Internal Reading Helpers
    # -------------------------------------------------------------------

    def _read_single_cell(self, selection, workbook_name: str,
                          sheet_name: str, index: int,
                          items: List[NumberItem], errors: List[str]):
        """
        WHAT: Reads a single selected cell directly (no SpecialCells).
        CALLED BY: read_selection() → when selection.Count == 1.
        """
        value = selection.Value
        address = selection.Address.replace("$", "")
        row = selection.Row
        col = selection.Column

        parsed = self._parse_cell_value(value)
        if parsed is not None:
            # Skip zeros — they don't contribute to sums
            if parsed == 0.0:
                return
            source = SourceTag(
                workbook_name=workbook_name,
                sheet_name=sheet_name,
                cell_address=address,
                row=row,
                column=col,
            )
            items.append(NumberItem(
                value=parsed,
                index=index,
                source=source,
                original_text=str(value) if value is not None else "",
            ))
        elif value is not None and str(value).strip():
            errors.append(f"Cell {address}: could not parse '{value}'")

    def _read_area(self, area, workbook_name: str, sheet_name: str,
                   start_index: int) -> Tuple[List[NumberItem], List[str]]:
        """
        WHAT:
            Reads all cells from a COM Range area. Tries bulk read first
            (area.Value returns everything at once), falls back to
            cell-by-cell iteration if that fails.

        CALLED BY:
            → read_selection() → for each area in visible_cells

        WHY:
            Bulk read is much faster for large selections. But some COM
            configurations or edge cases cause it to fail, so the
            cell-by-cell fallback ensures we always get the data.

        PARAMETERS:
            area: COM Range object (one contiguous area).
            workbook_name (str): For SourceTag.
            sheet_name (str): For SourceTag.
            start_index (int): Starting NumberItem.index.

        RETURNS:
            Tuple[List[NumberItem], List[str]]: (items, errors)
        """
        try:
            # Try bulk read
            values = area.Value
            if values is None:
                return [], []

            base_row = area.Row
            base_col = area.Column

            # Normalize to list of (value, row, col) tuples
            cells = []

            if not isinstance(values, tuple):
                # Single value (one cell in this area)
                cells.append((values, base_row, base_col))
            elif len(values) > 0 and isinstance(values[0], tuple):
                # 2D block: tuple of row-tuples
                for row_idx, row_data in enumerate(values):
                    for col_idx, cell_value in enumerate(row_data):
                        cells.append((
                            cell_value,
                            base_row + row_idx,
                            base_col + col_idx,
                        ))
            else:
                # 1D: single row or single column
                for idx, cell_value in enumerate(values):
                    if area.Rows.Count > 1:
                        # Column vector
                        cells.append((cell_value, base_row + idx, base_col))
                    else:
                        # Row vector
                        cells.append((cell_value, base_row, base_col + idx))

            # Process all cells
            return self._process_cell_list(
                cells, workbook_name, sheet_name, start_index,
            )

        except Exception:
            # Bulk read failed — fall back to cell-by-cell
            return self._read_area_cell_by_cell(
                area, workbook_name, sheet_name, start_index,
            )

    def _read_area_cell_by_cell(self, area, workbook_name: str,
                                sheet_name: str,
                                start_index: int) -> Tuple[List[NumberItem], List[str]]:
        """
        WHAT: Fallback reader that iterates one cell at a time via COM.
        CALLED BY: _read_area() → when bulk read fails.

        WHY:
            Some COM configurations or protected sheets cause bulk Value
            reads to throw exceptions. Individual cell reads are slower
            but almost always work.
        """
        cells = []
        for row_idx in range(1, area.Rows.Count + 1):
            for col_idx in range(1, area.Columns.Count + 1):
                try:
                    cell = area.Cells(row_idx, col_idx)
                    value = cell.Value
                    row = cell.Row
                    col = cell.Column
                    cells.append((value, row, col))
                except Exception:
                    continue

        return self._process_cell_list(
            cells, workbook_name, sheet_name, start_index,
        )

    def _process_cell_list(self, cells: list, workbook_name: str,
                           sheet_name: str,
                           start_index: int) -> Tuple[List[NumberItem], List[str]]:
        """
        WHAT: Converts a list of (value, row, col) tuples into NumberItems.
        CALLED BY: _read_area(), _read_area_cell_by_cell()
        """
        items: List[NumberItem] = []
        errors: List[str] = []
        current_index = start_index

        for value, row, col in cells:
            address = self._column_letter(col) + str(row)
            parsed = self._parse_cell_value(value)

            if parsed is not None:
                # Skip zeros — they don't contribute to sums
                if parsed == 0.0:
                    continue
                source = SourceTag(
                    workbook_name=workbook_name,
                    sheet_name=sheet_name,
                    cell_address=address,
                    row=row,
                    column=col,
                )
                items.append(NumberItem(
                    value=parsed,
                    index=current_index,
                    source=source,
                    original_text=str(value) if value is not None else "",
                ))
                current_index += 1
            elif value is not None and str(value).strip():
                errors.append(f"Cell {address}: could not parse '{value}'")

        return items, errors

    # -------------------------------------------------------------------
    # Value Parsing
    # -------------------------------------------------------------------

    @staticmethod
    def _parse_cell_value(value) -> Optional[float]:
        """
        WHAT:
            Parses a cell value to float. Handles int, float, and string
            values. Strips thousand separators from strings before parsing.

        CALLED BY:
            → _read_single_cell(), _process_cell_list()

        EDGE CASES HANDLED:
            - None → returns None (empty cell)
            - int/float → rounds and returns
            - String "1,00,000.50" → strips commas → 100000.50
            - Non-numeric string → returns None

        PARAMETERS:
            value: Raw cell value from Excel COM (any type).

        RETURNS:
            float or None: Parsed value, or None if not a number.
        """
        if value is None:
            return None

        if isinstance(value, (int, float)):
            return round(float(value), ROUNDING_PRECISION)

        # String value: strip thousand separators and try to parse
        text = str(value).strip()
        if not text:
            return None

        text = text.replace(THOUSAND_SEPARATOR, "")

        try:
            return round(float(text), ROUNDING_PRECISION)
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _column_letter(col_num: int) -> str:
        """
        WHAT: Converts a 1-based column number to Excel letter notation.
        CALLED BY: _process_cell_list()

        EXAMPLES:
            1 → "A", 2 → "B", 26 → "Z", 27 → "AA", 28 → "AB"

        PARAMETERS:
            col_num (int): 1-based column number.

        RETURNS:
            str: Column letter(s) in Excel notation.
        """
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result
