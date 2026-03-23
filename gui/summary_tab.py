"""
FILE: gui/summary_tab.py

PURPOSE: The Summary tab. Shows colored finalization cards, an Undo button,
         reconciliation statistics, and (Phase 5) an export button.

CONTAINS:
- SummaryTab — QWidget with finalization cards, stats, and unmatched marking

DEPENDS ON:
- config/constants.py → HIGHLIGHT_COLORS
- models/finalized_combination.py → FinalizedCombination
- utils/format_helpers.py → format_number_indian, format_difference
- gui/styles.py → color constants

USED BY:
- gui/main_window.py → embeds SummaryTab as the second tab

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — summary tab with cards + undo   | Sub-phase 2B finalization UI     |
| 23-03-2026 | Added Mark Unmatched button + count display | Phase 4B unmatched numbers view  |
"""

# Group 1: Python standard library
from typing import List

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QScrollArea, QFrame, QGroupBox,
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor

# Group 3: This project's modules
from models.finalized_combination import FinalizedCombination
from utils.format_helpers import format_number_indian, format_difference
from gui.styles import (
    COLOR_ACCENT, COLOR_TEXT_SECONDARY, COLOR_TEXT_PRIMARY,
    COLOR_ERROR, COLOR_SUCCESS, COLOR_PANEL_BG, COLOR_BORDER,
    scaled_size,
)


class SummaryTab(QWidget):
    """
    WHAT:
        Displays a scrollable list of colored finalization cards and
        reconciliation statistics. Each card shows the combo number,
        color, label, sum, difference, item count, and values.

    WHY ADDED:
        After finalizing combinations, users need to review what they've
        matched. The summary tab gives an at-a-glance view with colored
        cards matching the highlight colors in Excel and the source list.

    CALLED BY:
        → gui/main_window.py → added as the "Summary" tab

    SIGNALS EMITTED:
        → undo_requested() — Emitted when user clicks Undo
        → mark_unmatched_requested() — Emitted when user clicks Mark Unmatched

    ASSUMPTIONS:
        - Cards are rebuilt from scratch each time refresh() is called.
          This is simpler than incremental updates and fast enough for
          the expected number of finalizations (typically < 50).
    """

    undo_requested = pyqtSignal()
    mark_unmatched_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        """
        WHAT: Creates the summary layout with header, cards area, and stats.
        CALLED BY: __init__()
        """
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # --- Header with title and undo button ---
        header_layout = QHBoxLayout()

        title = QLabel("Finalization Summary")
        title.setStyleSheet(f"font-weight: bold; color: {COLOR_ACCENT}; font-size: {scaled_size(18)}px;")
        header_layout.addWidget(title)

        header_layout.addStretch()

        self._mark_unmatched_button = QPushButton("Mark Unmatched")
        self._mark_unmatched_button.setEnabled(False)
        self._mark_unmatched_button.setToolTip(
            "Highlight remaining (non-finalized) Excel cells in grey"
        )
        self._mark_unmatched_button.setStyleSheet(
            f"QPushButton {{ background-color: #9CA3AF; color: white; }}"
            f"QPushButton:hover {{ background-color: #6B7280; }}"
            f"QPushButton:disabled {{ background-color: #D1D5DB; color: #9CA3AF; }}"
        )
        self._mark_unmatched_button.clicked.connect(self.mark_unmatched_requested.emit)
        header_layout.addWidget(self._mark_unmatched_button)

        self._undo_button = QPushButton("Undo Last")
        self._undo_button.setEnabled(False)
        self._undo_button.setStyleSheet(
            f"QPushButton {{ background-color: {COLOR_ERROR}; }}"
            f"QPushButton:hover {{ background-color: #DC2626; }}"
            f"QPushButton:disabled {{ background-color: #D1D5DB; color: #9CA3AF; }}"
        )
        self._undo_button.clicked.connect(self.undo_requested.emit)
        header_layout.addWidget(self._undo_button)

        layout.addLayout(header_layout)

        # --- Stats bar ---
        self._stats_label = QLabel("No combinations finalized yet")
        self._stats_label.setStyleSheet(
            f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(14)}px; padding: 6px;"
        )
        layout.addWidget(self._stats_label)

        # --- Scrollable cards area ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        self._cards_container = QWidget()
        self._cards_layout = QVBoxLayout(self._cards_container)
        self._cards_layout.setContentsMargins(0, 0, 0, 0)
        self._cards_layout.setSpacing(10)
        self._cards_layout.addStretch()  # Push cards to top

        scroll.setWidget(self._cards_container)
        layout.addWidget(scroll, 1)

    # -----------------------------------------------------------------------
    # Public Methods
    # -----------------------------------------------------------------------

    def refresh(
        self,
        finalized_list: List[FinalizedCombination],
        can_undo: bool,
        unmatched_count: int = 0,
        total_items: int = 0,
    ):
        """
        WHAT:
            Rebuilds all finalization cards and updates stats.
            Called after every finalization, undo, or clear.

        CALLED BY:
            → gui/main_window.py → after finalization events

        PARAMETERS:
            finalized_list (List[FinalizedCombination]): All finalized combos.
            can_undo (bool): Whether the Undo button should be enabled.
            unmatched_count (int): Number of items not in any finalized combo.
            total_items (int): Total number of loaded items.
        """
        self._undo_button.setEnabled(can_undo)
        # Enable Mark Unmatched only when there are finalized combos AND unmatched items
        self._mark_unmatched_button.setEnabled(
            len(finalized_list) > 0 and unmatched_count > 0
        )
        self._rebuild_cards(finalized_list)
        self._update_stats(finalized_list, unmatched_count, total_items)

    def clear(self):
        """
        WHAT: Clears all cards and resets stats.
        CALLED BY: gui/find_tab.py → on clear all.
        """
        self._clear_cards()
        self._undo_button.setEnabled(False)
        self._mark_unmatched_button.setEnabled(False)
        self._stats_label.setText("No combinations finalized yet")

    def show_unmatched_result(self, cells_highlighted: int):
        """
        WHAT:
            Updates the Mark Unmatched button text to show how many cells
            were highlighted grey. Called by MainWindow after the highlighting
            completes.

        CALLED BY:
            → gui/main_window.py → after highlight_unmatched() completes

        PARAMETERS:
            cells_highlighted (int): Number of Excel cells marked grey.
        """
        if cells_highlighted > 0:
            self._mark_unmatched_button.setText(
                f"Marked {cells_highlighted} cell{'s' if cells_highlighted != 1 else ''}"
            )
            self._mark_unmatched_button.setEnabled(False)
        else:
            self._mark_unmatched_button.setText("Mark Unmatched")

    # -----------------------------------------------------------------------
    # Internal Methods
    # -----------------------------------------------------------------------

    def _rebuild_cards(self, finalized_list: List[FinalizedCombination]):
        """
        WHAT: Clears and rebuilds all finalization cards.
        CALLED BY: refresh()
        """
        self._clear_cards()

        for finalized in finalized_list:
            card = self._create_card(finalized)
            # Insert before the stretch at the end
            self._cards_layout.insertWidget(
                self._cards_layout.count() - 1, card
            )

    def _clear_cards(self):
        """
        WHAT: Removes all card widgets from the cards layout.
        CALLED BY: _rebuild_cards(), clear()
        """
        while self._cards_layout.count() > 1:  # Keep the stretch
            item = self._cards_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def _create_card(self, finalized: FinalizedCombination) -> QFrame:
        """
        WHAT:
            Creates a single finalization card widget showing the combo
            number, color, label, sum, difference, and item values.

        CALLED BY:
            → _rebuild_cards()

        PARAMETERS:
            finalized (FinalizedCombination): The finalization to display.

        RETURNS:
            QFrame: The card widget.
        """
        card = QFrame()
        card.setFrameShape(QFrame.StyledPanel)

        # Set card background to the finalization color (lighter version)
        r, g, b = finalized.color_rgb
        card.setStyleSheet(
            f"QFrame {{ background-color: rgba({r}, {g}, {b}, 80); "
            f"border: 2px solid rgb({r}, {g}, {b}); "
            f"border-radius: 6px; padding: 8px; }}"
            f"QLabel {{ background-color: transparent; }}"
        )

        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(6)

        # --- Top line: number, color name, label ---
        top_layout = QHBoxLayout()

        combo_label = QLabel(f"#{finalized.combo_number}")
        combo_label.setStyleSheet(
            f"font-weight: bold; font-size: {scaled_size(16)}px; color: {COLOR_TEXT_PRIMARY};"
        )
        top_layout.addWidget(combo_label)

        color_label = QLabel(f"({finalized.color_name})")
        color_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px;")
        top_layout.addWidget(color_label)

        if finalized.label:
            user_label = QLabel(f'"{finalized.label}"')
            user_label.setStyleSheet(
                f"font-style: italic; color: {COLOR_TEXT_PRIMARY}; font-size: {scaled_size(14)}px;"
            )
            top_layout.addWidget(user_label)

        top_layout.addStretch()
        card_layout.addLayout(top_layout)

        # --- Middle line: sum, difference, item count ---
        combo = finalized.combination
        sum_str = format_number_indian(combo.sum_value)
        diff_str = format_difference(combo.difference)
        count = combo.size

        diff_color = COLOR_SUCCESS if diff_str == "Exact" else COLOR_TEXT_SECONDARY
        middle_label = QLabel(
            f"Sum: {sum_str}  |  Diff: {diff_str}  |  "
            f"{count} item{'s' if count != 1 else ''}"
        )
        middle_label.setStyleSheet(f"font-size: {scaled_size(14)}px; color: {COLOR_TEXT_PRIMARY};")
        card_layout.addWidget(middle_label)

        # --- Bottom line: item values with source info ---
        value_parts = []
        for item in combo.items:
            val_str = format_number_indian(item.value)
            if item.source is not None:
                # Include source reference: (Book→Sheet:Cell)
                src = item.source
                short_wb = src.workbook_name
                # Truncate long workbook names
                if len(short_wb) > 20:
                    short_wb = short_wb[:17] + "..."
                val_str += f" ({short_wb}→{src.sheet_name}:{src.cell_address})"
            value_parts.append(val_str)

        if len(value_parts) > 8:
            values_text = ", ".join(value_parts[:8]) + f", ... ({len(value_parts)} total)"
        else:
            values_text = ", ".join(value_parts)

        values_label = QLabel(f"Values: {values_text}")
        values_label.setStyleSheet(
            f"font-size: {scaled_size(13)}px; color: {COLOR_TEXT_SECONDARY};"
        )
        values_label.setWordWrap(True)
        card_layout.addWidget(values_label)

        return card

    def _update_stats(
        self,
        finalized_list: List[FinalizedCombination],
        unmatched_count: int = 0,
        total_items: int = 0,
    ):
        """
        WHAT: Updates the statistics label with totals and unmatched count.
        CALLED BY: refresh()

        PARAMETERS:
            finalized_list (List[FinalizedCombination]): All finalized combos.
            unmatched_count (int): Number of items not in any finalized combo.
            total_items (int): Total number of loaded items.
        """
        if not finalized_list:
            self._stats_label.setText("No combinations finalized yet")
            return

        total_combos = len(finalized_list)
        matched_items = sum(f.combination.size for f in finalized_list)
        total_value = sum(f.combination.sum_value for f in finalized_list)

        exact = sum(1 for f in finalized_list
                    if format_difference(f.combination.difference) == "Exact")
        approx = total_combos - exact

        stats_parts = [
            f"{total_combos} combination{'s' if total_combos != 1 else ''}",
            f"{matched_items} items matched",
            f"Total: {format_number_indian(total_value)}",
        ]
        if exact > 0:
            stats_parts.append(f"{exact} exact")
        if approx > 0:
            stats_parts.append(f"{approx} approximate")
        if unmatched_count > 0:
            stats_parts.append(f"{unmatched_count} unmatched")

        self._stats_label.setText("  |  ".join(stats_parts))
        self._stats_label.setStyleSheet(
            f"color: {COLOR_SUCCESS}; font-size: {scaled_size(14)}px; padding: 6px; font-weight: bold;"
        )
