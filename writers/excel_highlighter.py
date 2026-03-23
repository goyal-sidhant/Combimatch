"""
FILE: writers/excel_highlighter.py

PURPOSE: Highlights finalized combination cells in Excel with assigned
         colors and writes labels as Shift+F2 notes (cell comments).
         Also provides undo (color removal) and unmatched grey marking.

CONTAINS:
- ExcelHighlighter — Writes colors and notes to Excel cells via COM

DEPENDS ON:
- readers/excel_reader.py → ExcelHandler (for COM access to Excel app)
- models/finalized_combination.py → FinalizedCombination
- models/number_item.py → NumberItem
- models/source_tag.py → SourceTag
- utils/format_helpers.py → format_number_indian

USED BY:
- gui/find_tab.py → highlights cells immediately on finalization
- gui/main_window.py → undo color removal on undo

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — Excel cell highlighter          | Sub-phase 3B Excel highlighting  |
| 22-03-2026 | Notes on ALL cells, not just first         | User found only first cell noted |
"""

# Group 1: Python standard library
from typing import List, Optional

# Group 3: This project's modules
from readers.excel_reader import ExcelHandler, PYWIN32_AVAILABLE
from models.finalized_combination import FinalizedCombination
from models.number_item import NumberItem
from utils.format_helpers import format_number_indian, format_difference


class ExcelHighlighter:
    """
    WHAT:
        Writes highlight colors and cell comments to Excel via COM
        automation. Each finalized combination's source cells get
        the assigned pastel color, and EVERY cell gets a Shift+F2
        note with the label, sum, and cross-references.

    WHY ADDED:
        The core value proposition of CombiMatch is visual — users see
        colored cells in Excel that show which invoices were matched.
        Without highlighting, the user would have to manually color
        cells, defeating the purpose.

    CALLED BY:
        → gui/find_tab.py → after finalization
        → gui/main_window.py → for undo (clear colors)

    ASSUMPTIONS:
        - The ExcelHandler is connected and the workbooks are still open.
          If Excel was closed, all methods return graceful errors.
        - Cell addresses in SourceTag match current Excel state. If the
          user moved/deleted rows after grabbing, highlighting will hit
          the wrong cells. This is an inherent limitation of live COM.
        *** ASSUMPTION: Excel COM uses BGR color format (Blue * 65536 +
            Green * 256 + Red), NOT standard RGB. The _rgb_to_excel()
            method handles this conversion. ***
    """

    def __init__(self, excel_handler: ExcelHandler):
        """
        WHAT: Creates the highlighter with a reference to the Excel handler.

        PARAMETERS:
            excel_handler (ExcelHandler): The shared handler for COM access.
        """
        self._handler = excel_handler

    # -------------------------------------------------------------------
    # Public Methods
    # -------------------------------------------------------------------

    def highlight_combination(self, finalized: FinalizedCombination) -> dict:
        """
        WHAT:
            Highlights all Excel-sourced cells in a finalized combination
            with the assigned color. Writes a Shift+F2 note on EVERY
            Excel-sourced cell with label, sum, and cross-references.

        CALLED BY:
            → gui/find_tab.py → immediately after finalization

        PARAMETERS:
            finalized (FinalizedCombination): The combo to highlight.

        RETURNS:
            dict: {"success": bool, "error": str or None, "data": {
                "cells_highlighted": int,
                "notes_written": int,
                "errors": List[str]
            }}
        """
        if not PYWIN32_AVAILABLE:
            return {
                "success": False,
                "error": "pywin32 is not installed — cannot highlight Excel cells",
                "data": None,
            }

        if not self._handler.is_connected:
            return {
                "success": False,
                "error": "Not connected to Excel",
                "data": None,
            }

        excel_color = self._rgb_to_excel(finalized.color_rgb)
        note_text = self._build_note_text(finalized)
        cells_highlighted = 0
        notes_written = 0
        errors = []

        for item in finalized.combination.items:
            if item.source is None:
                continue  # Manual item — no Excel cell to highlight

            # Highlight cell with color
            result = self._highlight_cell(
                item.source.workbook_name,
                item.source.sheet_name,
                item.source.cell_address,
                excel_color,
            )
            if result["success"]:
                cells_highlighted += 1
            else:
                errors.append(result["error"])

            # Write Shift+F2 note on EVERY cell (not just first)
            note_result = self._write_cell_note(
                item.source.workbook_name,
                item.source.sheet_name,
                item.source.cell_address,
                note_text,
            )
            if note_result["success"]:
                notes_written += 1
            else:
                errors.append(f"Note: {note_result['error']}")

        return {
            "success": cells_highlighted > 0 or len(errors) == 0,
            "error": "; ".join(errors) if errors else None,
            "data": {
                "cells_highlighted": cells_highlighted,
                "notes_written": notes_written,
                "errors": errors,
            },
        }

    def remove_highlight(self, finalized: FinalizedCombination) -> dict:
        """
        WHAT:
            Removes highlight colors and notes from a finalized
            combination's cells. Used during undo.

        CALLED BY:
            → gui/find_tab.py → after undo_last_finalization()

        PARAMETERS:
            finalized (FinalizedCombination): The combo to unhighlight.

        RETURNS:
            dict: {"success": bool, "error": str or None, "data": {
                "cells_cleared": int
            }}
        """
        if not PYWIN32_AVAILABLE or not self._handler.is_connected:
            return {"success": False, "error": "Not connected to Excel", "data": None}

        cells_cleared = 0
        errors = []

        for item in finalized.combination.items:
            if item.source is None:
                continue

            result = self._clear_cell_highlight(
                item.source.workbook_name,
                item.source.sheet_name,
                item.source.cell_address,
            )
            if result["success"]:
                cells_cleared += 1
            else:
                errors.append(result["error"])

        return {
            "success": cells_cleared > 0 or len(errors) == 0,
            "error": "; ".join(errors) if errors else None,
            "data": {"cells_cleared": cells_cleared},
        }

    def highlight_unmatched(self, items: List[NumberItem], grey_rgb: tuple) -> dict:
        """
        WHAT:
            Highlights remaining (non-finalized) Excel-sourced items
            in grey to visually mark them as unmatched.

        CALLED BY:
            → gui/summary_tab.py → "Mark Unmatched" button (Phase 4B)

        PARAMETERS:
            items (List[NumberItem]): Items to mark as unmatched.
            grey_rgb (tuple): The grey RGB color (e.g., (220, 220, 220)).

        RETURNS:
            dict: {"success": bool, "error": str or None, "data": {
                "cells_highlighted": int
            }}
        """
        if not PYWIN32_AVAILABLE or not self._handler.is_connected:
            return {"success": False, "error": "Not connected to Excel", "data": None}

        excel_color = self._rgb_to_excel(grey_rgb)
        cells_highlighted = 0
        errors = []

        for item in items:
            if item.source is None or item.is_finalized:
                continue

            result = self._highlight_cell(
                item.source.workbook_name,
                item.source.sheet_name,
                item.source.cell_address,
                excel_color,
            )
            if result["success"]:
                cells_highlighted += 1
            else:
                errors.append(result["error"])

        return {
            "success": True,
            "error": "; ".join(errors) if errors else None,
            "data": {"cells_highlighted": cells_highlighted},
        }

    # -------------------------------------------------------------------
    # Internal Methods
    # -------------------------------------------------------------------

    def _highlight_cell(
        self,
        workbook_name: str,
        sheet_name: str,
        cell_address: str,
        excel_color: int,
    ) -> dict:
        """
        WHAT: Sets the background color of a single Excel cell.
        CALLED BY: highlight_combination(), highlight_unmatched()

        RETURNS:
            dict: {"success": bool, "error": str or None}
        """
        try:
            app = self._handler._excel_app
            wb = app.Workbooks(workbook_name)
            ws = wb.Sheets(sheet_name)
            cell = ws.Range(cell_address)
            cell.Interior.Color = excel_color
            return {"success": True, "error": None}
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not highlight {workbook_name}→{sheet_name}:{cell_address}: {str(e)}",
            }

    def _clear_cell_highlight(
        self,
        workbook_name: str,
        sheet_name: str,
        cell_address: str,
    ) -> dict:
        """
        WHAT: Removes background color and comment from a single cell.
        CALLED BY: remove_highlight()

        RETURNS:
            dict: {"success": bool, "error": str or None}
        """
        try:
            app = self._handler._excel_app
            wb = app.Workbooks(workbook_name)
            ws = wb.Sheets(sheet_name)
            cell = ws.Range(cell_address)
            # xlNone = -4142 — removes fill
            cell.Interior.ColorIndex = -4142
            # Remove comment if it exists
            try:
                if cell.Comment is not None:
                    cell.Comment.Delete()
            except Exception:
                pass  # No comment to remove
            return {"success": True, "error": None}
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not clear {workbook_name}→{sheet_name}:{cell_address}: {str(e)}",
            }

    def _write_cell_note(
        self,
        workbook_name: str,
        sheet_name: str,
        cell_address: str,
        note_text: str,
    ) -> dict:
        """
        WHAT:
            Writes a Shift+F2 comment (note) on a cell. Replaces any
            existing comment. This is visible when the user hovers
            or presses Shift+F2.

        CALLED BY: highlight_combination()

        RETURNS:
            dict: {"success": bool, "error": str or None}
        """
        try:
            app = self._handler._excel_app
            wb = app.Workbooks(workbook_name)
            ws = wb.Sheets(sheet_name)
            cell = ws.Range(cell_address)
            # Remove existing comment if any
            try:
                if cell.Comment is not None:
                    cell.Comment.Delete()
            except Exception:
                pass
            # Add new comment
            cell.AddComment(note_text)
            # Auto-size the comment box for readability
            cell.Comment.Shape.TextFrame.AutoSize = True
            return {"success": True, "error": None}
        except Exception as e:
            return {
                "success": False,
                "error": f"Could not write note on {cell_address}: {str(e)}",
            }

    def _build_note_text(self, finalized: FinalizedCombination) -> str:
        """
        WHAT:
            Builds the text for the Shift+F2 cell note. Includes:
            - CombiMatch marker
            - Combo number and label
            - Sum and difference
            - All values with source references

        CALLED BY: highlight_combination()

        RETURNS:
            str: The formatted note text.
        """
        combo = finalized.combination
        lines = []

        # Header
        lines.append(f"CombiMatch #{finalized.combo_number}")
        if finalized.label:
            lines.append(f'Label: "{finalized.label}"')

        # Sum and difference
        sum_str = format_number_indian(combo.sum_value)
        diff_str = format_difference(combo.difference)
        lines.append(f"Sum: {sum_str}")
        if diff_str != "Exact":
            lines.append(f"Difference: {diff_str}")
        else:
            lines.append("Match: Exact")

        # Item count
        lines.append(f"Items: {combo.size}")

        # Values with cross-references
        lines.append("")
        lines.append("Values:")
        for item in combo.items:
            val_str = format_number_indian(item.value)
            if item.source is not None:
                src = item.source
                lines.append(f"  {val_str} ({src.workbook_name}→{src.sheet_name}:{src.cell_address})")
            else:
                lines.append(f"  {val_str} (manual input)")

        return "\n".join(lines)

    @staticmethod
    def _rgb_to_excel(rgb: tuple) -> int:
        """
        WHAT:
            Converts RGB tuple to Excel's BGR integer color format.
            Excel uses Blue * 65536 + Green * 256 + Red (BGR),
            NOT the standard R * 65536 + G * 256 + B (RGB).

        CALLED BY: highlight_combination(), highlight_unmatched()

        PARAMETERS:
            rgb (tuple): (R, G, B) values, each 0-255.

        RETURNS:
            int: Excel-format BGR color integer.
        """
        r, g, b = rgb
        return b * 65536 + g * 256 + r
