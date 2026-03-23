"""
FILE: models/finalized_combination.py

PURPOSE: Defines the FinalizedCombination dataclass — a locked-in
         combination with an assigned color, optional label, and metadata.

CONTAINS:
- FinalizedCombination — dataclass wrapping a Combination with finalization
                         details (color, label, sequence number, timestamp)

DEPENDS ON:
- models/combination.py → Combination

USED BY:
- core/finalization_manager.py → creates FinalizedCombination on finalize
- gui/summary_tab.py → displays as colored cards
- core/session_manager.py → serializes for session save/restore
- writers/report_writer.py → includes in audit trail report

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — finalized combination model     | Project skeleton setup (Phase 1A)|
"""

# Group 1: Python standard library
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

# Group 3: This project's modules
from models.combination import Combination


@dataclass
class FinalizedCombination:
    """
    WHAT:
        A Combination that the user has locked in ("finalized"). Carries
        the assigned highlight color, optional user label, sequence number,
        and timestamp. This is the record that appears on the Summary tab
        and in the audit trail report.

    WHY ADDED:
        Finalization adds metadata (color, label) that doesn't belong on
        the Combination itself. A separate dataclass keeps the solver's
        output clean and the finalization record self-contained.

    ASSUMPTIONS:
        - combo_number is 1-based (first finalization = #1).
        - timestamp is set at creation time, not at display time.
    """

    # The combination that was finalized.
    combination: Combination = field(default_factory=Combination)

    # The RGB color tuple assigned to this finalization.
    # e.g., (173, 216, 230) for Light Blue.
    color_rgb: tuple[int, int, int] = (0, 0, 0)

    # The human-readable color name (e.g., "Light Blue").
    color_name: str = ""

    # Optional user-provided label (e.g., "Dec Payment", "TDS Adj").
    # Empty string if user skipped the label popup.
    label: str = ""

    # Sequential finalization number (1-based). First finalization = #1.
    combo_number: int = 0

    # When this combination was finalized. Used for session restore ordering.
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())
