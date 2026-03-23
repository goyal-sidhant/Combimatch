"""
FILE: models/source_tag.py

PURPOSE: Defines the SourceTag dataclass — tracks where a number came from
         when grabbed from Excel (workbook, sheet, cell address).

CONTAINS:
- SourceTag — dataclass with workbook_name, sheet_name, cell_address fields

DEPENDS ON:
- Nothing from this project (uses only Python standard library).

USED BY:
- models/number_item.py → NumberItem has an optional SourceTag field
- readers/excel_reader.py → creates SourceTag for each grabbed number
- writers/excel_highlighter.py → uses SourceTag to find cells for highlighting
- writers/report_writer.py → includes source info in audit trail

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — source tracking for Excel items | Project skeleton setup (Phase 1A)|
"""

# Group 1: Python standard library
from dataclasses import dataclass
from typing import Optional


@dataclass
class SourceTag:
    """
    WHAT:
        Tracks the origin of a number grabbed from Excel.
        Attached to each NumberItem so that finalization can highlight
        the correct cell across multiple workbooks and sheets.

    WHY ADDED:
        Multi-workbook/sheet grab (N8) requires knowing exactly where
        each number came from. Without source tracking, cross-workbook
        highlighting and audit trail reports would be impossible.

    ASSUMPTIONS:
        - workbook_name is the filename (e.g., "Apr-Jun 2025.xlsx"),
          not the full path. The full path is resolved at highlight time
          via the Excel COM connection.
        - cell_address uses Excel notation (e.g., "A5", "B12").
    """

    # The Excel workbook filename (e.g., "Apr-Jun 2025.xlsx")
    workbook_name: str = ""

    # The sheet name within the workbook (e.g., "Sheet1", "GSTR-2B")
    sheet_name: str = ""

    # The cell address in Excel notation (e.g., "A5")
    cell_address: str = ""

    # The row number in Excel (1-based, for display purposes)
    row: Optional[int] = None

    # The column number in Excel (1-based, for internal use)
    column: Optional[int] = None
