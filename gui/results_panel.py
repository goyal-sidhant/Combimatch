"""
FILE: gui/results_panel.py

PURPOSE: The middle panel of the Find tab. Displays search results in
         two sections: exact matches and approximate matches. Results
         are grouped by combination size with collapsible headers
         showing counts.

CONTAINS:
- ResultsPanel — QWidget with exact and approximate result lists

DEPENDS ON:
- config/constants.py → EXACT_MATCH_THRESHOLD
- models/combination.py → Combination
- utils/format_helpers.py → format_number_indian, format_difference
- gui/styles.py → color constants

USED BY:
- gui/find_tab.py → embeds ResultsPanel in the middle section of the splitter

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — results panel with size groups  | Sub-phase 1E results + find tab  |
| 22-03-2026 | Updated — max_results caps approx only    | Exact matches must never be lost |
| 22-03-2026 | Added Finalize button + signal            | Sub-phase 2B finalization UI     |
| 22-03-2026 | Added results undo stack, new combo format | UI fixes before Phase 3          |
| 23-03-2026 | Added Deselect button + deselection support | User couldn't deselect a combo  |
"""

# Group 1: Python standard library
from typing import List, Optional, Dict, Tuple
from copy import copy

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QAbstractItemView, QSplitter, QGroupBox, QPushButton,
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QBrush

# Group 3: This project's modules
from config.constants import EXACT_MATCH_THRESHOLD
from models.combination import Combination
from utils.format_helpers import format_number_indian, format_difference
from gui.styles import (
    COLOR_ACCENT, COLOR_SUCCESS, COLOR_WARNING,
    COLOR_TEXT_SECONDARY, COLOR_HEADER_BG, COLOR_TEXT_PRIMARY,
    COLOR_PANEL_BG, scaled_size,
)


# Qt data roles for storing data on list items
COMBO_DATA_ROLE = Qt.UserRole          # Stores the Combination object
IS_HEADER_ROLE = Qt.UserRole + 1       # True if this item is a size-group header
IS_COLLAPSED_ROLE = Qt.UserRole + 2    # True if this header's children are collapsed


class ResultsPanel(QWidget):
    """
    WHAT:
        Displays search results in two lists: exact matches (sum equals
        target within threshold) and approximate matches (within buffer
        but not exact). Results are grouped by combination size with
        collapsible headers.

    WHY ADDED:
        The original tool had this but with bugs: header counts drifted
        after finalization, and the index mismatch caused wrong combos
        to be removed. This version stores Combination objects on each
        list item via Qt.UserRole, avoiding index-based lookups entirely.

    CALLED BY:
        → gui/find_tab.py → creates and embeds this panel

    SIGNALS EMITTED:
        → combination_selected(object) — Emitted when user clicks a result.
                                          Payload is the Combination object.
        → finalize_requested() — Emitted when user clicks the Finalize button.

    ASSUMPTIONS:
        - Results arrive in batches via add_results().
        - Each list item stores its Combination in COMBO_DATA_ROLE.
        - Header items have IS_HEADER_ROLE = True and are non-selectable.
        *** ASSUMPTION: Results within each size group are sorted by
            closeness to target (ascending absolute difference). The
            solver may not guarantee this, so we sort on insertion. ***
    """

    combination_selected = pyqtSignal(object)  # Combination or None
    deselection_requested = pyqtSignal()       # Emitted when user deselects
    finalize_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._exact_combos: List[Combination] = []
        self._approx_combos: List[Combination] = []
        self._results_undo_stack: List[Tuple[List[Combination], List[Combination]]] = []
        self._setup_ui()

    def _setup_ui(self):
        """
        WHAT: Creates the two result lists (exact and approximate).
        CALLED BY: __init__()
        """
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # Title
        title = QLabel("Results")
        title.setStyleSheet(f"font-weight: bold; color: {COLOR_ACCENT}; font-size: {scaled_size(17)}px;")
        layout.addWidget(title)

        # Splitter for exact / approximate sections
        splitter = QSplitter(Qt.Vertical)

        # --- Exact matches ---
        exact_container = QWidget()
        exact_layout = QVBoxLayout(exact_container)
        exact_layout.setContentsMargins(0, 0, 0, 0)
        exact_layout.setSpacing(2)

        self._exact_header = QLabel("Exact Matches (0)")
        self._exact_header.setStyleSheet(
            f"font-weight: bold; color: {COLOR_SUCCESS}; font-size: {scaled_size(14)}px;"
        )
        exact_layout.addWidget(self._exact_header)

        self._exact_list = QListWidget()
        self._exact_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self._exact_list.currentItemChanged.connect(self._on_exact_item_changed)
        self._exact_list.clicked.connect(lambda idx: self._on_list_clicked(self._exact_list, idx))
        exact_layout.addWidget(self._exact_list)

        splitter.addWidget(exact_container)

        # --- Approximate matches ---
        approx_container = QWidget()
        approx_layout = QVBoxLayout(approx_container)
        approx_layout.setContentsMargins(0, 0, 0, 0)
        approx_layout.setSpacing(2)

        self._approx_header = QLabel("Approximate Matches (0)")
        self._approx_header.setStyleSheet(
            f"font-weight: bold; color: {COLOR_WARNING}; font-size: {scaled_size(14)}px;"
        )
        approx_layout.addWidget(self._approx_header)

        self._approx_list = QListWidget()
        self._approx_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self._approx_list.currentItemChanged.connect(self._on_approx_item_changed)
        self._approx_list.clicked.connect(lambda idx: self._on_list_clicked(self._approx_list, idx))
        approx_layout.addWidget(self._approx_list)

        splitter.addWidget(approx_container)

        # Give exact matches more space by default
        splitter.setSizes([300, 200])
        layout.addWidget(splitter, 1)

        # --- Finalize / Deselect buttons ---
        button_row = QHBoxLayout()

        self._finalize_button = QPushButton("Finalize Selected")
        self._finalize_button.setEnabled(False)
        self._finalize_button.setStyleSheet(
            f"QPushButton {{ background-color: {COLOR_SUCCESS}; color: white; "
            f"font-weight: bold; padding: 10px; font-size: {scaled_size(15)}px; }}"
            f"QPushButton:hover {{ background-color: #059669; }}"
            f"QPushButton:disabled {{ background-color: #D1D5DB; color: #9CA3AF; }}"
        )
        self._finalize_button.clicked.connect(self.finalize_requested.emit)
        button_row.addWidget(self._finalize_button, 2)

        self._deselect_button = QPushButton("Deselect")
        self._deselect_button.setEnabled(False)
        self._deselect_button.setToolTip(
            "Clear the current selection.\n"
            "Removes orange highlight from source list."
        )
        self._deselect_button.clicked.connect(self._on_deselect_clicked)
        button_row.addWidget(self._deselect_button, 1)

        layout.addLayout(button_row)

    # -----------------------------------------------------------------------
    # Public Methods
    # -----------------------------------------------------------------------

    def add_results(self, combinations: List[Combination]):
        """
        WHAT:
            Adds a batch of combinations incrementally to the appropriate
            list (exact or approximate). Uses incremental insertion into
            size-group sections instead of rebuilding the entire list,
            keeping the UI responsive during long searches.

        CALLED BY:
            → gui/find_tab.py → on solver results_batch signal

        PARAMETERS:
            combinations (List[Combination]): Batch of results to add.
        """
        exact_batch = []
        approx_batch = []

        for combo in combinations:
            if abs(combo.difference) < EXACT_MATCH_THRESHOLD:
                self._exact_combos.append(combo)
                exact_batch.append(combo)
            else:
                self._approx_combos.append(combo)
                approx_batch.append(combo)

        # Incrementally insert new combos (no full rebuild)
        if exact_batch:
            self._add_combos_incremental(self._exact_list, exact_batch)
        if approx_batch:
            self._add_combos_incremental(self._approx_list, approx_batch)
        self._update_headers()

    def clear(self):
        """
        WHAT: Removes all results from both lists.
        CALLED BY: gui/find_tab.py → on new search or clear all.
        """
        self._exact_combos.clear()
        self._approx_combos.clear()
        self._exact_list.clear()
        self._approx_list.clear()
        self._results_undo_stack.clear()
        self._update_headers()
        self._finalize_button.setEnabled(False)
        self._deselect_button.setEnabled(False)

    def get_selected_combination(self) -> Optional[Combination]:
        """
        WHAT: Returns the currently selected Combination, or None.
        CALLED BY: gui/find_tab.py → for finalization.
        """
        # Check exact list first
        item = self._exact_list.currentItem()
        if item and item.data(COMBO_DATA_ROLE) is not None:
            return item.data(COMBO_DATA_ROLE)

        # Then approximate list
        item = self._approx_list.currentItem()
        if item and item.data(COMBO_DATA_ROLE) is not None:
            return item.data(COMBO_DATA_ROLE)

        return None

    def save_results_snapshot(self):
        """
        WHAT:
            Saves a copy of the current results state onto an undo stack.
            Called before finalization removes invalid combos, so undo
            can restore them.

        CALLED BY:
            → gui/find_tab.py → before remove_invalid_combinations()
        """
        self._results_undo_stack.append(
            (list(self._exact_combos), list(self._approx_combos))
        )

    def restore_results_snapshot(self) -> bool:
        """
        WHAT:
            Restores the most recent results snapshot from the undo stack.
            Returns True if a snapshot was restored, False if stack was empty.

        CALLED BY:
            → gui/find_tab.py → after undo_last_finalization()

        RETURNS:
            bool: True if restored, False if nothing to restore.
        """
        if not self._results_undo_stack:
            return False

        self._exact_combos, self._approx_combos = self._results_undo_stack.pop()
        self._rebuild_list(self._exact_list, self._exact_combos)
        self._rebuild_list(self._approx_list, self._approx_combos)
        self._update_headers()
        self._finalize_button.setEnabled(False)
        self._deselect_button.setEnabled(False)
        return True

    def remove_invalid_combinations(self, finalized_indices: set):
        """
        WHAT:
            Removes combinations that contain any finalized item.
            Called after finalization to clean up stale results.

        WHY:
            After an item is finalized (locked into a matched combo),
            any other combination containing that item is no longer
            valid. This method removes them.

        CALLED BY:
            → gui/find_tab.py → after finalization completes

        PARAMETERS:
            finalized_indices (set): Set of item indices that were just finalized.

        EDGE CASES HANDLED:
            - Combination with ALL items finalized → removed
            - Combination with ONE finalized item → removed
            - Empty headers after removal → removed
            - No overlapping combos → nothing changes
        """
        # Filter out combos that overlap with finalized indices
        self._exact_combos = [
            combo for combo in self._exact_combos
            if not combo.item_indices.intersection(finalized_indices)
        ]
        self._approx_combos = [
            combo for combo in self._approx_combos
            if not combo.item_indices.intersection(finalized_indices)
        ]

        # Rebuild both lists
        self._rebuild_list(self._exact_list, self._exact_combos)
        self._rebuild_list(self._approx_list, self._approx_combos)
        self._update_headers()

    # -----------------------------------------------------------------------
    # Internal Methods
    # -----------------------------------------------------------------------

    def _add_combos_incremental(self, list_widget: QListWidget, new_combos: List[Combination]):
        """
        WHAT:
            Inserts new combinations into existing size-group sections
            without clearing and rebuilding the entire list. If a size
            group doesn't exist yet, creates a new header and section.
            Much faster than full rebuild for streaming search results.

        CALLED BY:
            → add_results()

        PARAMETERS:
            list_widget (QListWidget): The list to update.
            new_combos (List[Combination]): New combinations to insert.
        """
        list_widget.setUpdatesEnabled(False)
        list_widget.blockSignals(True)

        # Group new combos by size
        by_size: Dict[int, List[Combination]] = {}
        for combo in new_combos:
            size = combo.size
            if size not in by_size:
                by_size[size] = []
            by_size[size].append(combo)

        for size in sorted(by_size.keys()):
            combos_for_size = by_size[size]

            # Find existing header for this size
            header_row = -1
            section_end = list_widget.count()  # Default: end of list

            for row in range(list_widget.count()):
                item = list_widget.item(row)
                if item.data(IS_HEADER_ROLE):
                    # Parse size from header text like "  ▼ 3 Numbers (5 found)"
                    text = item.text()
                    # Extract the number after ▼ or ▶
                    for part in text.split():
                        if part.isdigit():
                            header_size = int(part)
                            if header_size == size:
                                header_row = row
                            elif header_size > size and header_row >= 0:
                                # Found next header after our section
                                section_end = row
                            elif header_size > size and header_row < 0:
                                # Insert new header before this larger one
                                section_end = row
                            break

            if header_row >= 0:
                # Existing section — find where it ends
                end_row = header_row + 1
                while end_row < list_widget.count():
                    next_item = list_widget.item(end_row)
                    if next_item.data(IS_HEADER_ROLE):
                        break
                    end_row += 1

                # Count existing items in this section
                existing_count = end_row - header_row - 1
                new_start_num = existing_count + 1

                # Insert new items at end of section
                for i, combo in enumerate(combos_for_size):
                    combo_num = new_start_num + i
                    item_text = self._format_combo_text(combo, combo_num)
                    list_item = QListWidgetItem(item_text)
                    list_item.setData(COMBO_DATA_ROLE, combo)
                    list_item.setData(IS_HEADER_ROLE, False)
                    list_widget.insertItem(end_row + i, list_item)

                # Update header count
                total_count = existing_count + len(combos_for_size)
                header_item = list_widget.item(header_row)
                is_collapsed = header_item.data(IS_COLLAPSED_ROLE)
                indicator = "▶" if is_collapsed else "▼"
                header_item.setText(
                    f"  {indicator} {size} Number{'s' if size != 1 else ''} ({total_count} found)"
                )

                # If section is collapsed, hide new items
                if is_collapsed:
                    for i in range(len(combos_for_size)):
                        list_widget.item(end_row + i).setHidden(True)
            else:
                # New section — create header and items
                # Insert at section_end (before larger groups, or at end)
                header_text = f"  ▼ {size} Number{'s' if size != 1 else ''} ({len(combos_for_size)} found)"
                header_item = QListWidgetItem(header_text)
                header_item.setData(IS_HEADER_ROLE, True)
                header_item.setData(COMBO_DATA_ROLE, None)
                header_item.setData(IS_COLLAPSED_ROLE, False)
                header_item.setFlags(Qt.ItemIsEnabled)

                header_font = QFont()
                header_font.setBold(True)
                header_font.setPointSize(13)
                header_item.setFont(header_font)
                header_item.setBackground(QBrush(QColor(COLOR_HEADER_BG)))
                header_item.setForeground(QBrush(QColor(COLOR_TEXT_PRIMARY)))

                list_widget.insertItem(section_end, header_item)

                for i, combo in enumerate(combos_for_size, start=1):
                    item_text = self._format_combo_text(combo, i)
                    list_item = QListWidgetItem(item_text)
                    list_item.setData(COMBO_DATA_ROLE, combo)
                    list_item.setData(IS_HEADER_ROLE, False)
                    list_widget.insertItem(section_end + i, list_item)

        list_widget.blockSignals(False)
        list_widget.setUpdatesEnabled(True)

    def _rebuild_list(self, list_widget: QListWidget, combos: List[Combination]):
        """
        WHAT:
            Clears and repopulates a QListWidget with size-grouped
            combinations. Each size group has a non-selectable header
            item followed by combination items.

        CALLED BY:
            → add_results(), remove_invalid_combinations()

        PARAMETERS:
            list_widget (QListWidget): The list to rebuild.
            combos (List[Combination]): The combinations to display.
        """
        # Block signals and visual updates during rebuild
        list_widget.setUpdatesEnabled(False)
        list_widget.blockSignals(True)
        list_widget.clear()

        if not combos:
            list_widget.blockSignals(False)
            list_widget.setUpdatesEnabled(True)
            return

        # Group by size
        size_groups: Dict[int, List[Combination]] = {}
        for combo in combos:
            size = combo.size
            if size not in size_groups:
                size_groups[size] = []
            size_groups[size].append(combo)

        # Sort each group by closeness to target
        for size in size_groups:
            size_groups[size].sort(key=lambda c: abs(c.difference))

        # Add items by size group (sorted by size)
        for size in sorted(size_groups.keys()):
            group = size_groups[size]

            # Size-group header (▼ expanded / ▶ collapsed)
            header_text = f"  ▼ {size} Number{'s' if size != 1 else ''} ({len(group)} found)"
            header_item = QListWidgetItem(header_text)
            header_item.setData(IS_HEADER_ROLE, True)
            header_item.setData(COMBO_DATA_ROLE, None)
            header_item.setData(IS_COLLAPSED_ROLE, False)
            header_item.setFlags(Qt.ItemIsEnabled)  # Not selectable

            # Style header
            header_font = QFont()
            header_font.setBold(True)
            header_font.setPointSize(13)
            header_item.setFont(header_font)
            header_item.setBackground(QBrush(QColor(COLOR_HEADER_BG)))
            header_item.setForeground(QBrush(QColor(COLOR_TEXT_PRIMARY)))

            list_widget.addItem(header_item)

            # Combination items — numbered within each size group
            for combo_num, combo in enumerate(group, start=1):
                item_text = self._format_combo_text(combo, combo_num)
                list_item = QListWidgetItem(item_text)
                list_item.setData(COMBO_DATA_ROLE, combo)
                list_item.setData(IS_HEADER_ROLE, False)
                list_widget.addItem(list_item)

        list_widget.blockSignals(False)
        list_widget.setUpdatesEnabled(True)

    def _format_combo_text(self, combo: Combination, combo_num: int) -> str:
        """
        WHAT:
            Formats a combination for display in the results list.
            Numbered list with values, equals sign, sum, difference.

        CALLED BY:
            → _rebuild_list()

        FORMAT:
            "1. 31.00, 542.00, 331.00 = 904.00"
            "2. 100.00, 200.00 = 300.00 (-96.00)"
        """
        sum_str = format_number_indian(combo.sum_value)
        diff_str = format_difference(combo.difference)

        # Individual values, comma-separated (truncate if too many)
        values = [format_number_indian(item.value) for item in combo.items]
        if len(values) > 8:
            values_str = ", ".join(values[:8]) + f", ... ({len(values)} total)"
        else:
            values_str = ", ".join(values)

        if diff_str == "Exact":
            return f"{combo_num}. {values_str} = {sum_str}"
        else:
            return f"{combo_num}. {values_str} = {sum_str} ({diff_str})"

    def _update_headers(self):
        """
        WHAT: Updates the section header labels with current counts.
        CALLED BY: add_results(), clear(), remove_invalid_combinations()
        """
        self._exact_header.setText(f"Exact Matches ({len(self._exact_combos)})")
        self._approx_header.setText(f"Approximate Matches ({len(self._approx_combos)})")

    def _on_list_clicked(self, list_widget: QListWidget, index):
        """
        WHAT:
            Handles click on a list item. If it's a header, toggles
            collapse/expand of the items below it (until the next header).

        CALLED BY:
            → QListWidget.clicked signal (both exact and approx lists)

        PARAMETERS:
            list_widget (QListWidget): The list that was clicked.
            index: The QModelIndex of the clicked item.
        """
        item = list_widget.item(index.row())
        if item is None:
            return

        # Only handle header clicks
        if not item.data(IS_HEADER_ROLE):
            return

        is_collapsed = item.data(IS_COLLAPSED_ROLE)
        new_collapsed = not is_collapsed
        item.setData(IS_COLLAPSED_ROLE, new_collapsed)

        # Update header text indicator (▼ expanded / ▶ collapsed)
        text = item.text()
        if new_collapsed:
            text = text.replace("▼", "▶", 1)
        else:
            text = text.replace("▶", "▼", 1)
        item.setText(text)

        # Toggle visibility of all items below until the next header
        row = index.row() + 1
        while row < list_widget.count():
            child_item = list_widget.item(row)
            if child_item.data(IS_HEADER_ROLE):
                break  # Stop at next header
            child_item.setHidden(new_collapsed)
            row += 1

    def _on_deselect_clicked(self):
        """
        WHAT:
            Clears the selection in both lists, disables Finalize and
            Deselect buttons, and emits deselection_requested so FindTab
            can clear orange highlights from the source panel.

        CALLED BY:
            → Deselect button click
        """
        self._exact_list.blockSignals(True)
        self._approx_list.blockSignals(True)
        self._exact_list.setCurrentItem(None)
        self._approx_list.setCurrentItem(None)
        self._exact_list.blockSignals(False)
        self._approx_list.blockSignals(False)
        self._finalize_button.setEnabled(False)
        self._deselect_button.setEnabled(False)
        self.deselection_requested.emit()

    def _on_exact_item_changed(self, current, previous):
        """
        WHAT: Handles selection change in the exact matches list.
        CALLED BY: Qt signal when user clicks an item.
        """
        if current is None:
            # Deselected — disable buttons if approx also has no selection
            approx_item = self._approx_list.currentItem()
            if approx_item is None or approx_item.data(COMBO_DATA_ROLE) is None:
                self._finalize_button.setEnabled(False)
                self._deselect_button.setEnabled(False)
            return

        # Skip headers
        if current.data(IS_HEADER_ROLE):
            return

        combo = current.data(COMBO_DATA_ROLE)
        if combo is not None:
            # Deselect in the other list
            self._approx_list.blockSignals(True)
            self._approx_list.setCurrentItem(None)
            self._approx_list.blockSignals(False)
            self._finalize_button.setEnabled(True)
            self._deselect_button.setEnabled(True)
            self.combination_selected.emit(combo)

    def _on_approx_item_changed(self, current, previous):
        """
        WHAT: Handles selection change in the approximate matches list.
        CALLED BY: Qt signal when user clicks an item.
        """
        if current is None:
            # Deselected — disable buttons if exact also has no selection
            exact_item = self._exact_list.currentItem()
            if exact_item is None or exact_item.data(COMBO_DATA_ROLE) is None:
                self._finalize_button.setEnabled(False)
                self._deselect_button.setEnabled(False)
            return

        # Skip headers
        if current.data(IS_HEADER_ROLE):
            self._finalize_button.setEnabled(False)
            self._deselect_button.setEnabled(False)
            return

        combo = current.data(COMBO_DATA_ROLE)
        if combo is not None:
            # Deselect in the other list
            self._exact_list.blockSignals(True)
            self._exact_list.setCurrentItem(None)
            self._exact_list.blockSignals(False)
            self._finalize_button.setEnabled(True)
            self._deselect_button.setEnabled(True)
            self.combination_selected.emit(combo)
