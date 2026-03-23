"""
FILE: gui/input_panel.py

PURPOSE: The left panel of the Find tab. Contains input mode selection,
         text area for pasting numbers, Load button, search parameters
         (target, buffer, min/max size, max results, search order),
         Find/Stop buttons, progress bar, and Clear All button (N3).

CONTAINS:
- InputPanel — QWidget with all input controls and search parameters

DEPENDS ON:
- config/constants.py → DEFAULT_BUFFER, DEFAULT_MIN_SIZE, DEFAULT_MAX_SIZE,
                         DEFAULT_MAX_RESULTS, MAX_RESULTS_UPPER_LIMIT,
                         SEARCH_ORDER_SMALLEST
- config/mappings.py → SEARCH_ORDER_OPTIONS
- core/number_parser.py → parse_numbers_line_separated, parse_numbers_semicolon_separated
- gui/styles.py → color constants

USED BY:
- gui/find_tab.py → embeds InputPanel in the left section of the splitter

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| 22-03-2026 | Created — input panel with all controls   | Sub-phase 1D input + source      |
| 22-03-2026 | Added Grab from Excel button + signal     | Sub-phase 3A Excel integration   |
| 22-03-2026 | Added target/buffer change signals          | Live smart bounds hints          |
| 23-03-2026 | Added seed info hint display               | Phase 4A seed numbers            |
| 23-03-2026 | Added solver indicator label               | Show C Solver vs Python Solver   |
"""

# Group 1: Python standard library
from typing import List

# Group 2: Third-party libraries
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
    QComboBox, QLineEdit, QSpinBox, QPushButton, QProgressBar,
    QGroupBox, QFormLayout, QMessageBox, QSizePolicy,
)
from PyQt5.QtCore import pyqtSignal, Qt

# Group 3: This project's modules
from config.constants import (
    DEFAULT_MIN_SIZE, DEFAULT_MAX_SIZE, DEFAULT_MAX_RESULTS,
    MAX_RESULTS_UPPER_LIMIT, SEARCH_ORDER_SMALLEST,
)
from config.mappings import SEARCH_ORDER_OPTIONS
from core.number_parser import parse_numbers_line_separated, parse_numbers_semicolon_separated
from utils.format_helpers import format_number_indian
from gui.styles import COLOR_ERROR, COLOR_SUCCESS, COLOR_TEXT_SECONDARY, scaled_size, scaled_px


class InputPanel(QWidget):
    """
    WHAT:
        The left panel of the Find tab. Provides all controls for
        loading numbers and configuring search parameters.

    WHY ADDED:
        Separating input controls into their own panel keeps find_tab.py
        manageable and gives each panel a single responsibility.

    CALLED BY:
        → gui/find_tab.py → creates and embeds this panel

    SIGNALS EMITTED:
        → numbers_loaded(list)    — Emitted with list of NumberItem when Load succeeds
        → find_requested()        — Emitted when user clicks Find
        → stop_requested()        — Emitted when user clicks Stop
        → clear_all_requested()   — Emitted when user clicks Clear All (N3)

    ASSUMPTIONS:
        - The Find tab connects to signals and handles solver orchestration.
        - This panel does NOT start the solver — it only signals the request.
    """

    # --- Signals ---
    numbers_loaded = pyqtSignal(list)     # List[NumberItem]
    find_requested = pyqtSignal()
    stop_requested = pyqtSignal()
    clear_all_requested = pyqtSignal()
    grab_excel_requested = pyqtSignal()   # Grab from Excel button clicked
    target_changed = pyqtSignal()         # Target text changed (for live bounds)
    buffer_changed = pyqtSignal()         # Buffer text changed (for live bounds)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        """
        WHAT: Creates all input controls and lays them out vertically.
        CALLED BY: __init__()
        """
        layout = QVBoxLayout(self)
        m = scaled_px(12)
        layout.setContentsMargins(m, m, m, m)
        layout.setSpacing(scaled_px(10))

        # --- Input Mode ---
        mode_layout = QHBoxLayout()
        mode_label = QLabel("Input Mode:")
        self._mode_combo = QComboBox()
        self._mode_combo.addItem("Line-separated", "line")
        self._mode_combo.addItem("Semicolon-separated", "semicolon")
        mode_layout.addWidget(mode_label)
        mode_layout.addWidget(self._mode_combo, 1)
        layout.addLayout(mode_layout)

        # --- Text Area ---
        self._text_area = QTextEdit()
        self._text_area.setPlaceholderText(
            "Paste numbers here...\n\n"
            "Line-separated mode:\n"
            "  100\n"
            "  200.50\n"
            "  1,00,000\n\n"
            "Or switch to semicolon mode:\n"
            "  100; 200.50; 1,00,000"
        )
        self._text_area.setMinimumHeight(scaled_px(100))
        layout.addWidget(self._text_area, 1)

        # --- Load / Grab Buttons + Status ---
        load_layout = QHBoxLayout()
        self._load_button = QPushButton("Load Numbers")
        self._load_button.clicked.connect(self._on_load_numbers)

        self._grab_excel_button = QPushButton("Grab from Excel")
        self._grab_excel_button.setToolTip(
            "Read numbers from Excel selection.\n"
            "Connect to Excel in the Settings tab first."
        )
        self._grab_excel_button.clicked.connect(self.grab_excel_requested.emit)

        load_layout.addWidget(self._load_button)
        load_layout.addWidget(self._grab_excel_button)
        layout.addLayout(load_layout)

        self._load_status = QLabel("")
        self._load_status.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px;")
        layout.addWidget(self._load_status)

        # --- Search Parameters Group ---
        params_group = QGroupBox("Search Parameters")
        params_layout = QFormLayout(params_group)
        params_layout.setSpacing(scaled_px(10))
        # Prevent fields from expanding the layout beyond the panel width
        params_layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        # Target
        self._target_input = QLineEdit()
        self._target_input.setPlaceholderText("e.g. 1,00,000")
        params_layout.addRow("Target Sum:", self._target_input)

        # Buffer
        self._buffer_input = QLineEdit()
        self._buffer_input.setPlaceholderText("0 = exact match")
        params_layout.addRow("Buffer (±):", self._buffer_input)

        # Min Size — with inline bounds hint
        min_size_row = QHBoxLayout()
        self._min_size_spin = QSpinBox()
        self._min_size_spin.setRange(DEFAULT_MIN_SIZE, 9999)
        self._min_size_spin.setValue(DEFAULT_MIN_SIZE)
        self._min_hint_label = QLabel("")
        self._min_hint_label.setStyleSheet(
            f"color: #1565C0; font-size: {scaled_size(12)}px; font-weight: bold;"
        )
        self._min_hint_label.setMinimumWidth(0)
        self._min_hint_label.setWordWrap(True)
        self._min_hint_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        min_size_row.addWidget(self._min_size_spin)
        min_size_row.addWidget(self._min_hint_label, 1)
        params_layout.addRow("Min Size:", min_size_row)

        # Max Size — with inline bounds hint
        max_size_row = QHBoxLayout()
        self._max_size_spin = QSpinBox()
        self._max_size_spin.setRange(DEFAULT_MIN_SIZE, 9999)
        self._max_size_spin.setValue(DEFAULT_MAX_SIZE)
        self._max_hint_label = QLabel("")
        self._max_hint_label.setStyleSheet(
            f"color: #1565C0; font-size: {scaled_size(12)}px; font-weight: bold;"
        )
        self._max_hint_label.setMinimumWidth(0)
        self._max_hint_label.setWordWrap(True)
        self._max_hint_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        max_size_row.addWidget(self._max_size_spin)
        max_size_row.addWidget(self._max_hint_label, 1)
        params_layout.addRow("Max Size:", max_size_row)

        # Max Results
        self._max_results_spin = QSpinBox()
        self._max_results_spin.setRange(1, MAX_RESULTS_UPPER_LIMIT)
        self._max_results_spin.setValue(DEFAULT_MAX_RESULTS)
        params_layout.addRow("Max Results:", self._max_results_spin)

        # Search Order
        self._search_order_combo = QComboBox()
        for key, label in SEARCH_ORDER_OPTIONS.items():
            self._search_order_combo.addItem(label, key)
        params_layout.addRow("Search Order:", self._search_order_combo)

        # Seed info hint (shown when seeds are pinned)
        self._seed_info_label = QLabel("")
        self._seed_info_label.setStyleSheet(
            f"color: #0D47A1; font-size: {scaled_size(13)}px; font-weight: bold;"
        )
        self._seed_info_label.setWordWrap(True)
        self._seed_info_label.setMinimumWidth(0)
        self._seed_info_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        params_layout.addRow("", self._seed_info_label)

        # Bounds summary — search space estimate + no-solution warning
        self._bounds_hint_label = QLabel("")
        self._bounds_hint_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px;")
        self._bounds_hint_label.setWordWrap(True)
        self._bounds_hint_label.setMinimumWidth(0)
        self._bounds_hint_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        params_layout.addRow("", self._bounds_hint_label)

        layout.addWidget(params_group)

        # --- Find / Stop Buttons ---
        button_layout = QHBoxLayout()
        self._find_button = QPushButton("Find Combinations")
        self._find_button.clicked.connect(self._on_find_clicked)
        self._stop_button = QPushButton("Stop")
        self._stop_button.setEnabled(False)
        self._stop_button.clicked.connect(self._on_stop_clicked)
        # Style stop button differently
        self._stop_button.setStyleSheet(
            f"QPushButton {{ background-color: {COLOR_ERROR}; }}"
            f"QPushButton:hover {{ background-color: #DC2626; }}"
            f"QPushButton:disabled {{ background-color: #D1D5DB; color: #9CA3AF; }}"
        )
        button_layout.addWidget(self._find_button, 2)
        button_layout.addWidget(self._stop_button, 1)
        layout.addLayout(button_layout)

        # --- Solver Indicator ---
        self._solver_label = QLabel("")
        self._solver_label.setStyleSheet(
            f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(11)}px;"
        )
        layout.addWidget(self._solver_label)

        # --- Progress Bar ---
        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 0)  # Indeterminate by default
        self._progress_bar.setVisible(False)
        self._progress_label = QLabel("")
        self._progress_label.setStyleSheet(f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px;")
        self._progress_label.setWordWrap(True)
        self._progress_label.setMinimumWidth(0)
        self._progress_label.setVisible(False)
        layout.addWidget(self._progress_bar)
        layout.addWidget(self._progress_label)

        # --- Clear All Button (N3) ---
        self._clear_all_button = QPushButton("Clear All Numbers")
        self._clear_all_button.setProperty("flat", True)
        self._clear_all_button.clicked.connect(self._on_clear_all_clicked)
        layout.addWidget(self._clear_all_button)

        # --- Wire target/buffer change signals for live bounds hints ---
        self._target_input.textChanged.connect(self._on_target_changed)
        self._buffer_input.textChanged.connect(self._on_buffer_changed)

    # -----------------------------------------------------------------------
    # Public Methods (called by find_tab)
    # -----------------------------------------------------------------------

    def get_target_text(self) -> str:
        """WHAT: Returns the raw target input text."""
        return self._target_input.text()

    def get_buffer_text(self) -> str:
        """WHAT: Returns the raw buffer input text."""
        return self._buffer_input.text()

    def get_min_size(self) -> int:
        """WHAT: Returns the min size spinbox value."""
        return self._min_size_spin.value()

    def get_max_size(self) -> int:
        """WHAT: Returns the max size spinbox value."""
        return self._max_size_spin.value()

    def get_max_results(self) -> int:
        """WHAT: Returns the max results spinbox value."""
        return self._max_results_spin.value()

    def get_search_order(self) -> str:
        """WHAT: Returns the selected search order constant."""
        return self._search_order_combo.currentData()

    def get_search_params(self) -> dict:
        """
        WHAT: Returns all search parameter values as a dict for session save.
        CALLED BY: gui/find_tab.py → get_session_state()
        """
        return {
            "target": self._target_input.text(),
            "buffer": self._buffer_input.text(),
            "min_size": self._min_size_spin.value(),
            "max_size": self._max_size_spin.value(),
            "max_results": self._max_results_spin.value(),
            "search_order_index": self._search_order_combo.currentIndex(),
        }

    def set_search_params(self, params: dict):
        """
        WHAT: Restores search parameter values from a saved session dict.
        CALLED BY: gui/find_tab.py → restore_session()

        PARAMETERS:
            params (dict): Keys match those from get_search_params().
        """
        if "target" in params:
            self._target_input.setText(str(params["target"]))
        if "buffer" in params:
            self._buffer_input.setText(str(params["buffer"]))
        if "min_size" in params:
            self._min_size_spin.setValue(int(params["min_size"]))
        if "max_size" in params:
            self._max_size_spin.setValue(int(params["max_size"]))
        if "max_results" in params:
            self._max_results_spin.setValue(int(params["max_results"]))
        if "search_order_index" in params:
            self._search_order_combo.setCurrentIndex(int(params["search_order_index"]))

    def set_searching_state(self, is_searching: bool):
        """
        WHAT: Toggles UI between searching and idle states.
        CALLED BY: gui/find_tab.py → when solver starts/stops.

        PARAMETERS:
            is_searching (bool): True when solver is running.
        """
        self._find_button.setEnabled(not is_searching)
        self._stop_button.setEnabled(is_searching)
        self._load_button.setEnabled(not is_searching)
        self._clear_all_button.setEnabled(not is_searching)
        self._progress_bar.setVisible(is_searching)
        self._progress_label.setVisible(is_searching)

        if is_searching:
            self._progress_bar.setRange(0, 0)  # Indeterminate
            self._progress_label.setText("Searching...")

    def update_progress(self, iterations: int, current_size: int):
        """
        WHAT: Updates the progress label with current search status.
        CALLED BY: gui/find_tab.py → on solver progress signal.
        """
        if current_size > 0:
            self._progress_label.setText(
                f"Checking {current_size}-number combinations... "
                f"({iterations:,} checked)"
            )

    def update_bounds_hint(self, bounds_result: dict):
        """
        WHAT:
            Shows smart bounds hints inline next to the Min/Max spinboxes
            and in the summary label. Makes viable range impossible to miss.

        CALLED BY: gui/find_tab.py → after computing smart bounds.

        DISPLAY:
            - Min Size: [1]  ← viable: 2
            - Max Size: [10] ← viable: 5
            - "Viable: 2 to 5 (1,200 combinations)" in summary
            - "No solution possible" in red if no sizes work
        """
        if bounds_result["no_solution"]:
            # Red warning on both spinbox hints and summary
            no_sol_style = f"color: {COLOR_ERROR}; font-size: {scaled_size(12)}px; font-weight: bold;"
            self._min_hint_label.setText("No solution")
            self._min_hint_label.setStyleSheet(no_sol_style)
            self._max_hint_label.setText("")
            self._bounds_hint_label.setText(
                "No solution possible with current parameters"
            )
            self._bounds_hint_label.setStyleSheet(
                f"color: {COLOR_ERROR}; font-size: {scaled_size(13)}px; font-weight: bold;"
            )
        else:
            smart_min = bounds_result["smart_min"]
            smart_max = bounds_result["smart_max"]
            est = bounds_result["estimated_combinations"]

            # Inline hints next to spinboxes
            hint_style = f"color: #1565C0; font-size: {scaled_size(12)}px; font-weight: bold;"
            self._min_hint_label.setText(f"← viable: {smart_min}")
            self._min_hint_label.setStyleSheet(hint_style)
            self._max_hint_label.setText(f"← viable: {smart_max}")
            self._max_hint_label.setStyleSheet(hint_style)

            # Summary label with combination count
            self._bounds_hint_label.setText(
                f"Viable: {smart_min} to {smart_max} "
                f"({est:,} combinations)"
            )
            self._bounds_hint_label.setStyleSheet(
                f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px;"
            )

    def clear_bounds_hint(self):
        """WHAT: Clears all bounds hint labels."""
        self._min_hint_label.setText("")
        self._max_hint_label.setText("")
        self._bounds_hint_label.setText("")

    def update_seed_info(self, seed_count: int, seed_sum: float):
        """
        WHAT: Shows the seed info hint (count and sum of pinned seeds).
        CALLED BY: gui/find_tab.py → when seeds change.

        PARAMETERS:
            seed_count (int): Number of pinned seeds.
            seed_sum (float): Sum of pinned seed values.
        """
        if seed_count == 0:
            self._seed_info_label.setText("")
        else:
            self._seed_info_label.setText(
                f"{seed_count} seed{'s' if seed_count != 1 else ''} pinned "
                f"(sum: {format_number_indian(seed_sum)})"
            )

    def clear_seed_info(self):
        """WHAT: Clears the seed info label."""
        self._seed_info_label.setText("")

    def update_solver_indicator(self, is_dll_available: bool):
        """
        WHAT: Shows which solver is active — C Solver or Python Solver.
        CALLED BY: gui/find_tab.py → on startup and when search starts.
        """
        if is_dll_available:
            self._solver_label.setText("Solver: C (solver.dll)")
            self._solver_label.setStyleSheet(
                f"color: #2E7D32; font-size: {scaled_size(11)}px;"
            )
        else:
            self._solver_label.setText("Solver: Python (fallback)")
            self._solver_label.setStyleSheet(
                f"color: {COLOR_TEXT_SECONDARY}; font-size: {scaled_size(11)}px;"
            )

    def show_search_complete(self, stats: dict):
        """
        WHAT: Updates progress area to show search completion stats.
        CALLED BY: gui/find_tab.py → on solver search_complete signal.
        """
        self._progress_bar.setVisible(False)

        total = stats.get("total_found", 0)
        exact = stats.get("exact_count", 0)
        approx = stats.get("approximate_count", 0)
        stopped = stats.get("was_stopped", False)

        parts = []
        if exact > 0:
            parts.append(f"{exact} exact")
        if approx > 0:
            parts.append(f"{approx} approximate")

        status = f"Found {total} combination{'s' if total != 1 else ''}"
        if parts:
            status += f" ({', '.join(parts)})"
        if stopped:
            status += " — stopped by user"

        self._progress_label.setText(status)
        self._progress_label.setVisible(True)
        self._progress_label.setStyleSheet(
            f"color: {COLOR_SUCCESS if total > 0 else COLOR_TEXT_SECONDARY}; font-size: {scaled_size(13)}px;"
        )

    def set_load_status(self, message: str, is_error: bool = False):
        """
        WHAT: Sets the load status label text and color.
        CALLED BY: gui/find_tab.py → after grab from Excel.
        """
        color = COLOR_ERROR if is_error else COLOR_SUCCESS
        self._load_status.setText(message)
        self._load_status.setStyleSheet(f"color: {color}; font-size: {scaled_size(13)}px;")

    def clear_all(self):
        """
        WHAT: Resets all input fields to defaults.
        CALLED BY: gui/find_tab.py → on clear_all_requested signal.
        """
        self._text_area.clear()
        self._target_input.clear()
        self._buffer_input.clear()
        self._min_size_spin.setValue(DEFAULT_MIN_SIZE)
        self._max_size_spin.setValue(DEFAULT_MAX_SIZE)
        self._max_results_spin.setValue(DEFAULT_MAX_RESULTS)
        self._search_order_combo.setCurrentIndex(0)
        self._load_status.setText("")
        self._progress_label.setText("")
        self._progress_label.setVisible(False)
        self._progress_bar.setVisible(False)
        self.clear_bounds_hint()
        self.clear_seed_info()

    # -----------------------------------------------------------------------
    # Internal Slots
    # -----------------------------------------------------------------------

    def _on_load_numbers(self):
        """
        WHAT:
            Parses the text area content and emits numbers_loaded signal
            with the parsed NumberItem list.

        CALLED BY:
            → Load Numbers button click

        CALLS:
            → core/number_parser.py → parse_numbers_line_separated or
              parse_numbers_semicolon_separated depending on mode

        EDGE CASES HANDLED:
            - Empty text → shows error message
            - Invalid lines → shows count of errors, still loads valid numbers
            - All lines invalid → shows error, does not emit signal
        """
        text = self._text_area.toPlainText().strip()

        if not text:
            self._load_status.setText("No numbers to load")
            self._load_status.setStyleSheet(f"color: {COLOR_ERROR}; font-size: {scaled_size(13)}px;")
            return

        # Parse based on selected mode
        mode = self._mode_combo.currentData()
        if mode == "semicolon":
            result = parse_numbers_semicolon_separated(text)
        else:
            result = parse_numbers_line_separated(text)

        if not result["success"]:
            self._load_status.setText(result["error"])
            self._load_status.setStyleSheet(f"color: {COLOR_ERROR}; font-size: {scaled_size(13)}px;")
            return

        items = result["data"]["items"]
        errors = result["data"]["errors"]

        if not items:
            self._load_status.setText("No valid numbers found")
            self._load_status.setStyleSheet(f"color: {COLOR_ERROR}; font-size: {scaled_size(13)}px;")
            return

        # Show status
        status = f"Loaded {len(items)} number{'s' if len(items) != 1 else ''}"
        if errors:
            status += f" ({len(errors)} skipped)"
        self._load_status.setText(status)
        self._load_status.setStyleSheet(f"color: {COLOR_SUCCESS}; font-size: {scaled_size(13)}px;")

        # Emit signal with parsed items
        self.numbers_loaded.emit(items)

    def _on_find_clicked(self):
        """
        WHAT: Emits find_requested signal when user clicks Find.
        CALLED BY: Find Combinations button click.
        """
        self.find_requested.emit()

    def _on_stop_clicked(self):
        """
        WHAT: Emits stop_requested signal when user clicks Stop.
        CALLED BY: Stop button click.
        """
        self.stop_requested.emit()

    def _on_clear_all_clicked(self):
        """
        WHAT: Confirms and emits clear_all_requested signal.
        CALLED BY: Clear All Numbers button click.
        """
        self.clear_all_requested.emit()

    def _on_target_changed(self):
        """
        WHAT: Emits target_changed signal for live smart bounds update.
        CALLED BY: Target QLineEdit textChanged signal.
        """
        self.target_changed.emit()

    def _on_buffer_changed(self):
        """
        WHAT: Emits buffer_changed signal for live smart bounds update.
        CALLED BY: Buffer QLineEdit textChanged signal.
        """
        self.buffer_changed.emit()
