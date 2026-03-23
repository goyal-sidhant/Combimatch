# PROJECT_AUDIT_COMBIMATCH.md — Complete Project Archaeology

**Audit Date:** 2026-03-20
**Auditor:** Claude Opus 4.6 (AI-assisted reverse-engineering)
**Project:** CombiMatch — Invoice Reconciliation Subset Sum Solver
**Audit Scope:** Every source file, every git commit, all configuration, full code path tracing

---

## SECTION 1: EXECUTIVE SUMMARY

CombiMatch is a Windows desktop application that helps accountants and finance professionals reconcile invoices by finding combinations of numbers that sum to a target value, with a configurable tolerance (buffer). It is purpose-built for a CA/GST consultant (Sidhant) who needs to match payment amounts against lists of invoice values — a common and tedious task in bank and invoice reconciliation. The tool is built with Python 3.8+, PyQt5 for the desktop GUI, and pywin32 for live Excel COM automation (reading selected cells from a running Excel instance and highlighting matched cells with colors). The core algorithm uses `itertools.combinations` to exhaustively test subsets of numbers, with smart bounds optimization to skip mathematically impossible combination sizes. The project appears **complete and functional** for its intended use case, with a coherent codebase across 10 source files (~2,400 lines of code), clear separation of concerns (models, solver, Excel handler, UI layer), and a comprehensive handover document from the original development conversation. One significant bug exists in the combination removal logic after finalization (Section 7), but the overall quality is solid for a tool built by a non-programmer through AI-assisted development.

---

## SECTION 2: PROJECT STRUCTURE MAP

```
combimatch/                              [ROOT]
├── main.py                              [CORE] Entry point — creates QApplication, applies styles, launches MainWindow
├── models.py                            [CORE] Data classes: NumberItem, Combination, FinalizedCombination, etc.
├── solver.py                            [CORE] SubsetSumSolver — combination-finding algorithm with smart bounds
├── excel_handler.py                     [CORE] ExcelHandler — COM automation for reading/highlighting Excel cells
├── utils.py                             [UTIL] Color management, number parsing, parameter validation
├── requirements.txt                     [CONFIG] Dependencies: PyQt5>=5.15.0, pywin32>=304
├── README.md                            [DOCS] User-facing documentation
├── .gitignore                           [CONFIG] Standard Python/PyQt ignores + .xlsx/.xls/.csv exclusions
├── HANDOVER.md                          [DOCS] Development session handover — decisions, bugs, reasoning
├── FILE_INDEX.md                        [DOCS] File inventory with descriptions
├── ui/                                  [UI PACKAGE]
│   ├── __init__.py                      [UI] Package init — re-exports MainWindow, FindTab, SummaryTab, SettingsTab, apply_styles
│   ├── main_window.py                   [UI] MainWindow — tab container, close handling, Excel monitor timer
│   ├── find_tab.py                      [UI/CORE] FindTab — main workspace (input, search, results, finalization)
│   ├── summary_tab.py                   [UI] SummaryTab — displays all finalized combinations with colors
│   ├── settings_tab.py                  [UI] SettingsTab — Excel connection, workbook/sheet selection, color preview
│   └── styles.py                        [UI] QSS stylesheet and color constants
├── .claude/                             [TOOLING] Claude Code AI assistant configuration
│   └── settings.local.json              [TOOLING] Permissions for bash commands (includes references to a csolver/ directory that does not exist in the repo)
├── __pycache__/                         [GENERATED] Python bytecode cache (safe to delete)
│   ├── excel_handler.cpython-313.pyc
│   ├── models.cpython-313.pyc
│   ├── solver.cpython-313.pyc
│   └── utils.cpython-313.pyc
└── ui/__pycache__/                      [GENERATED] UI package bytecode cache (safe to delete)
    ├── __init__.cpython-313.pyc
    ├── find_tab.cpython-313.pyc
    ├── main_window.cpython-313.pyc
    ├── settings_tab.cpython-313.pyc
    ├── styles.cpython-313.pyc
    └── summary_tab.cpython-313.pyc
```

**Entry Point:** `python main.py`

### Dead Code Inventory

| Item | Location | Type | Evidence |
|------|----------|------|----------|
| `format_number()` | `utils.py:233` | Dead function | Defined but never called from any file |
| `_column_number()` | `excel_handler.py:537` | Dead method | Defined but never called from any file (the reverse conversion `_column_letter` IS used) |
| `clear_highlight()` | `excel_handler.py:481` | Dead method | Defined in ExcelHandler but never called from any file (highlights are applied but never individually cleared) |
| `find_combinations()` | `solver.py:100` | Dead method | The non-generator version; only `find_combinations_generator()` is used by the app |
| `CombinationResult` class | `models.py:120` | Partially dead class | Only referenced inside the dead `find_combinations()` method and its own `__init__`; the generator path bypasses it entirely |
| `_get_combo_avg_position()` | `find_tab.py:92` | Dead method | Static method on `SelectedComboInfoPanel`, never called (remnant of abandoned sorting feature) |
| `check_excel_connection()` | `find_tab.py:1069` | Dead method | Defined on `FindTab` but never called from any file |
| `get_color()` | `styles.py:308` | Dead function | Defined but never called from any file (referenced only in FILE_INDEX.md documentation) |
| `placeholder` attribute | `find_tab.py:153` | Dead attribute | Created in `SelectedComboInfoPanel._init_ui()` but never added to any layout — it is invisible |

---

## SECTION 3: FEATURE INVENTORY

### F1: Manual Number Input (Line-Separated)

- **What it does:** User pastes or types numbers into a text box, one per line. The tool parses them into `NumberItem` objects.
- **Where it lives:** `utils.py:76-111` (`parse_numbers_line_separated`), triggered from `find_tab.py:459-479` (`_on_load_numbers`)
- **Input → Output:** Raw text → list of `NumberItem` objects + error messages for unparseable lines
- **User interaction:** Select "Line Separated" mode, type/paste numbers, click "Load Numbers"
- **Concrete trace:** Input: `"100.50\n200\nabc\n-50"`. Line 1 → strip → `"100.50"` → remove commas/spaces → `float("100.50")` = 100.5 → `NumberItem(value=100.5, index=0)`. Line 2 → 200.0 → `NumberItem(value=200.0, index=1)`. Line 3 → `"abc"` → `float()` raises `ValueError` → error: `"Line 3: 'abc' is not a valid number"`. Line 4 → `"-50"` → -50.0 → `NumberItem(value=-50.0, index=2)`. Returns 3 items + 1 error.

### F2: Manual Number Input (Comma-Separated)

- **What it does:** Same as F1 but splits on commas instead of newlines.
- **Where it lives:** `utils.py:114-150` (`parse_numbers_comma_separated`), same trigger path
- **Input → Output:** Comma-separated text → list of `NumberItem` objects + errors
- **User interaction:** Select "Comma Separated" mode, type/paste, click "Load Numbers"
- **Concrete trace:** Input: `"100, 200, 300"`. Split by comma → `["100", " 200", " 300"]`. Each stripped and parsed. Returns 3 `NumberItem` objects with indices 0, 1, 2.

> **Note:** Comma-separated mode strips spaces but does NOT handle thousand-separator commas (e.g., `"1,000"` would be split into `"1"` and `"000"` → parsed as 1.0 and 0.0). Line-separated mode DOES handle thousand-separator commas by removing them before parsing. This is an inconsistency.

### F3: Excel Selection Reading

- **What it does:** Reads the currently selected cells from a running Excel instance, respecting filters (only visible cells).
- **Where it lives:** `excel_handler.py:239-423` (`read_selection`), triggered from `find_tab.py:481-531` (`_on_grab_from_excel`)
- **Input → Output:** Excel selection → list of `NumberItem` objects (with row/column metadata) + warnings
- **User interaction:** Select cells in Excel, switch to "Excel Selection" mode, click "Grab from Excel"
- **Concrete trace:** User selects A1:A5 in Excel with A3 hidden by filter. `selection.Count` > 1, so proceed to `SpecialCells(12)` to get only visible cells. This returns areas for A1:A2 and A4:A5 (A3 is hidden). For each area, `area.Value` returns a tuple of tuples. For area A1:A2 with values (100, 200): `raw_values = ((100,), (200,))`, detected as single column, converted to `[[100], [200]]`. Each value becomes a `NumberItem` with row/column metadata. Result: 4 items (rows 1,2,4,5), skipping hidden row 3.

### F4: Combination Finding (Subset Sum Solver)

- **What it does:** Finds all combinations of loaded numbers whose sum falls within `target ± buffer`. Searches from smallest to largest combination size. Uses smart bounds to skip impossible sizes.
- **Where it lives:** `solver.py:16-183` (`SubsetSumSolver`), `solver.py:248-306` (`compute_smart_bounds`), triggered from `find_tab.py:567-655` (`_on_find_combinations`)
- **Input → Output:** List of `NumberItem` + target + buffer + size constraints → stream of `Combination` objects
- **User interaction:** Set target sum, buffer, min/max sizes, max results; click "Find Combinations"
- **Concrete trace:** Items: [10, 20, 30, 40, 50], target: 50, buffer: 1. Lower=49, upper=51. Smart bounds: sorted desc [50,40,30,20,10], cumsum at k=1: 50 >= 49, so smart_min=1. Sorted asc [10,20,30,40,50], cumsum at k=3: 60 > 51, so smart_max=2. Skipped sizes: [3,4,5]. Search only sizes 1-2. Size 1: check each — 50 (sum=50, within 49-51 ✓). Size 2: check all pairs — (10,40)=50 ✓, (20,30)=50 ✓. Yields 3 combinations.

### F5: Exact vs. Approximate Results Split

- **What it does:** Separates found combinations into "exact" (difference < 0.001) and "approximate" (within buffer but not exact).
- **Where it lives:** `find_tab.py:657-673` (`_on_combo_batch`)
- **Input → Output:** Combination → routed to exact_list or approx_list based on `abs(combo.difference) < 0.001`
- **User interaction:** Automatic — results appear in two separate list widgets

### F6: Size-Grouped Headers with Counts

- **What it does:** Groups results by combination size with visual headers like "═══ 3 Numbers (5 found) ═══".
- **Where it lives:** `find_tab.py:946-990` (`_add_combo_to_list`)
- **Input → Output:** Combination → header created/updated + numbered combo added to list widget
- **User interaction:** Automatic — headers appear as non-selectable grey bars in the results lists

### F7: Combination Selection and Highlighting

- **What it does:** When user clicks a combination in either results list, highlights the corresponding numbers in the source list with orange background, bold text, and "▶" marker prefix.
- **Where it lives:** `find_tab.py:692-788` (`_on_exact_selected`, `_on_approx_selected`, `_handle_combo_selection`, `_highlight_combination`)
- **Input → Output:** Click on combo → source list items highlighted
- **User interaction:** Click any combination in exact or approximate list
- **Concrete trace:** User clicks a 2-number combo [10, 30]. `_handle_combo_selection` sets `selected_combo`, enables Finalize button, updates info panel. `_highlight_combination` iterates source list: for items with matching indices, sets `QColor(255, 165, 0)` background, bold font, prepends "▶ " to text. For other items, restores default white background and normal font. For finalized items, keeps their assigned color.

### F8: Selected Combination Info Panel

- **What it does:** Displays sum, difference from target, item count, and individual values of the currently selected combination.
- **Where it lives:** `find_tab.py:89-177` (`SelectedComboInfoPanel`)
- **Input → Output:** Combination → formatted display of sum, difference (green if exact, yellow if approximate), count, and bullet-pointed values
- **User interaction:** Automatic — updates when any combination is selected

### F9: Finalization with Color Assignment

- **What it does:** Locks in a selected combination, assigns a unique color from a palette of 20 soft colors, marks items as finalized in the source list with "✓" prefix, highlights corresponding cells in Excel, and removes invalid combinations from the results.
- **Where it lives:** `find_tab.py:813-856` (`_on_finalize`), `utils.py:40-73` (`ColorManager`)
- **Input → Output:** Selected combo → items marked as finalized, color assigned, Excel cells highlighted, invalid combos removed
- **User interaction:** Click "Finalize Selected" button
- **Concrete trace:** User finalizes combo [10, 30] (items at indices 0, 2). `ColorManager.get_next_color()` returns `((173, 216, 230), "Light Blue")`. Items at indices 0 and 2 get `is_finalized=True`, `finalized_color=(173, 216, 230)`. Source list updates: items show "✓ #1: 10.00" with light blue background. Excel cells at corresponding rows get light blue fill. `_remove_invalid_combinations` removes any combo containing items 0 or 2. Signal emitted to SummaryTab.

> **BUG (CRITICAL):** The `_remove_invalid_combinations` → `_filter_combination_list` method has a **widget-row/list-index mismatch bug** when size-group headers are present. See Section 7 for full details.

### F10: Excel Workbook and Sheet Selection

- **What it does:** Lets user connect to a running Excel instance, select a specific workbook and sheet.
- **Where it lives:** `settings_tab.py:39-322` (`SettingsTab`), `excel_handler.py:112-198` (workbook/sheet methods)
- **Input → Output:** User selects workbook name → sheets populate → user selects sheet → sheet activated in Excel
- **User interaction:** Settings tab → Click "Connect to Excel" → Select workbook from dropdown → Select sheet from dropdown

### F11: Summary Tab (Finalized Combinations View)

- **What it does:** Displays all finalized combinations as colored cards showing combo number, color name, sum, difference, item count, values, and row references.
- **Where it lives:** `ui/summary_tab.py:1-197` (`SummaryTab`, `SummaryCard`)
- **Input → Output:** `FinalizedCombination` signal → new card added to scrollable list, total sum updated
- **User interaction:** Switch to Summary tab to view all finalized combos

### F12: Stop Search

- **What it does:** Allows user to stop a long-running search mid-progress.
- **Where it lives:** `find_tab.py:940-944` (`_on_stop_clicked`), `solver.py:91-93` (`SubsetSumSolver.stop()`)
- **Input → Output:** Button click → `_stop_requested` flag set → solver loop exits
- **User interaction:** Click "Stop Search" button (appears during active search, replacing the Find button)

### F13: Progress Tracking

- **What it does:** Shows real-time progress during search: which combination size is being checked and how many combinations have been tested.
- **Where it lives:** `find_tab.py:933-938` (`_on_progress`), `solver.py:181` (callback every 1000 iterations)
- **Input → Output:** Progress callback → label text update
- **User interaction:** Automatic — label visible during search

### F14: Smart Combination Size Bounds

- **What it does:** Before searching, calculates which combination sizes are mathematically impossible using cumulative sum analysis, and skips them entirely. Shows user which sizes were skipped and displays live "viable min/max" hints on the spinbox controls.
- **Where it lives:** `solver.py:248-306` (`compute_smart_bounds`), `find_tab.py:992-1053` (`_update_bounds_hints`)
- **Input → Output:** Item values + target range → tightened min/max sizes + skipped sizes list
- **User interaction:** Automatic pruning during search; live hint labels next to Min/Max spinboxes

### F15: Auto-Switch to Excel Mode

- **What it does:** When user connects to Excel in Settings tab, the Find tab automatically switches input mode to "Excel Selection".
- **Where it lives:** `settings_tab.py:44` (signal), `main_window.py:89-91` (connection), `find_tab.py:1055-1057` (`switch_to_excel_mode`)
- **Input → Output:** Excel connection → mode combo set to index 2
- **User interaction:** Automatic on Excel connection

### F16: Close-Event Save Handling

- **What it does:** When closing the app, prompts to save the Excel file (if connected and there are finalized combos), then confirms close.
- **Where it lives:** `main_window.py:121-167` (`closeEvent`)
- **Input → Output:** Close event → save prompt → optional save → confirm close → cleanup
- **User interaction:** Close the application window

### F17: Excel Connection Monitor

- **What it does:** Timer that checks every 5 seconds if Excel is still running. If Excel closes, shows a warning.
- **Where it lives:** `main_window.py:106-119` (`_start_excel_monitor`, `_check_excel`), `settings_tab.py:308-322` (`check_excel_closed`)
- **Input → Output:** Timer tick → COM ping → warning if disconnected
- **User interaction:** Automatic background monitoring

### F18: Batch UI Updates

- **What it does:** Instead of updating the UI for every single combination found, batches results (up to 20 or every 0.1 seconds) and disables list widget updates during batch processing to prevent UI freezing.
- **Where it lives:** `find_tab.py:42-87` (`SolverThread` with `BATCH_SIZE=20`, `BATCH_INTERVAL=0.1`), `find_tab.py:657-673` (`_on_combo_batch`)
- **User interaction:** Automatic — smoother UI during search

---

## SECTION 4: DATA FLOW

```
╔═══════════════════════════════════════════════════════════════════════╗
║                        DATA ENTRY POINTS                             ║
╠═══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  [Manual Text Input]          [Excel Selection]                       ║
║       │                            │                                  ║
║       ▼                            ▼                                  ║
║  parse_numbers_*()           ExcelHandler.read_selection()            ║
║  (utils.py)                  (excel_handler.py)                      ║
║       │                            │                                  ║
║       └──────────┬─────────────────┘                                  ║
║                  ▼                                                    ║
║         List[NumberItem]                                              ║
║         + List[errors/warnings]                                       ║
║                  │                                                    ║
║                  ▼                                                    ║
║      FindTab._set_items()                                            ║
║      (stores in self.items)                                          ║
║      (populates source_list widget)                                  ║
║                                                                       ║
╠═══════════════════════════════════════════════════════════════════════╣
║                        SEARCH PIPELINE                               ║
╠═══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  validate_parameters() ← target, buffer, min/max size, max results   ║
║         │                                                             ║
║         ▼                                                             ║
║  quick_check_possible() → False? → "No combinations possible"        ║
║         │ True                                                        ║
║         ▼                                                             ║
║  SubsetSumSolver.__init__()                                          ║
║    ├── Filters out finalized items                                    ║
║    ├── compute_smart_bounds() → skips impossible sizes                ║
║    └── no_solution? → "All sizes ruled out"                           ║
║         │                                                             ║
║         ▼                                                             ║
║  SolverThread.run()  [Background QThread]                            ║
║    └── find_combinations_generator()                                  ║
║         ├── itertools.combinations(items, size)                       ║
║         ├── sum(values) within lower_bound..upper_bound?              ║
║         │     Yes → yield Combination                                 ║
║         └── Every 1000 iter → progress callback                       ║
║                  │                                                    ║
║    Batching (20 results or 0.1s interval)                            ║
║                  │                                                    ║
║                  ▼                                                    ║
║  results_batch signal → _on_combo_batch()                            ║
║    ├── abs(difference) < 0.001 → exact_list + exact_combinations     ║
║    └── else → approx_list + approx_combinations                      ║
║         (with size-group headers via _add_combo_to_list)              ║
║                                                                       ║
╠═══════════════════════════════════════════════════════════════════════╣
║                     SELECTION & FINALIZATION                         ║
╠═══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  User clicks combo in list                                            ║
║         │                                                             ║
║         ▼                                                             ║
║  _handle_combo_selection()                                           ║
║    ├── Update info panel (sum, diff, values)                          ║
║    ├── _highlight_combination() (orange + bold + "▶" in source list)  ║
║    └── Enable Finalize button                                         ║
║         │                                                             ║
║  User clicks "Finalize Selected"                                      ║
║         │                                                             ║
║         ▼                                                             ║
║  _on_finalize()                                                      ║
║    ├── ColorManager.get_next_color() → unique color                   ║
║    ├── Mark items: is_finalized=True, finalized_color=color           ║
║    ├── Create FinalizedCombination record                             ║
║    ├── _update_source_colors() → "✓" prefix, colored background       ║
║    ├── ExcelHandler.highlight_items() → color cells in Excel          ║
║    ├── combination_finalized signal → SummaryTab.add_finalized()     ║
║    └── _remove_invalid_combinations() → remove combos using           ║
║         finalized items from both lists [BUGGY - see Section 7]       ║
║                                                                       ║
╠═══════════════════════════════════════════════════════════════════════╣
║                        DATA EXIT POINTS                              ║
╠═══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  [Screen Display]              [Excel Highlighting]                   ║
║  - Source list with colors     - Cell background colors               ║
║  - Exact/Approx result lists   - Applied via ExcelHandler             ║
║  - Info panel details          - Prompted to save on close            ║
║  - Summary tab cards                                                  ║
║                                                                       ║
╚═══════════════════════════════════════════════════════════════════════╝
```

---

## SECTION 5: BUSINESS RULES AND DOMAIN LOGIC

### Core Algorithm Rules

| Rule | Location | Details |
|------|----------|---------|
| Buffer is absolute, not percentage | `solver.py:49` | `self.buffer = round(abs(buffer), 2)` — always positive |
| Target range = target ± buffer | `solver.py:56-57` | `lower_bound = target - buffer`, `upper_bound = target + buffer` |
| Valid sum check | `solver.py:95-98` | `lower_bound <= round(total, 2) <= upper_bound` |
| Exact match threshold | `find_tab.py:663` | `abs(combo.difference) < 0.001` — any difference under 0.001 is "exact" |
| Min combo size floor | `solver.py:50` | `max(1, min_size)` — cannot be less than 1 |
| Max combo size ceiling | `solver.py:51` | `min(max_size, len(items))` — cannot exceed available items |
| Finalized items excluded from search | `solver.py:47` | `self.items = [item for item in items if not item.is_finalized]` |
| Values rounded to 2 decimal places | `models.py:50-52` | `NumberItem.__post_init__` rounds `self.value = round(self.value, 2)` |
| Search order: smallest combos first | `solver.py:111,162` | `for size in range(self.min_size, self.max_size + 1)` |
| Results sorted by (size, closeness) | `models.py:152-154` | `sort(key=lambda c: (c.size, abs(c.difference)))` (only in dead `find_combinations` method) |
| Progress reported every 1000 iterations | `solver.py:138,181` | `if total_checked % 1000 == 0` |

### Smart Bounds Rules

| Rule | Location | Details |
|------|----------|---------|
| Smart min: top-k sum must reach lower_bound | `solver.py:280-287` | Sort descending, cumulative sum; first k where sum ≥ lower_bound |
| Smart max: bottom-k sum must not exceed upper_bound | `solver.py:293-299` | Sort ascending, cumulative sum; first k where sum > upper_bound, then max = k-1 |
| No solution detection | `solver.py:304` | `smart_min > smart_max` means no valid size exists |

### Input Parsing Rules

| Rule | Location | Details |
|------|----------|---------|
| Line-separated: commas removed as thousand separators | `utils.py:98` | `line.replace(',', '').replace(' ', '')` |
| Comma-separated: spaces removed but commas are delimiters | `utils.py:138` | `part.replace(' ', '')` — commas are split characters, not removed from values |
| Empty lines/parts skipped | `utils.py:93-94, 133-134` | `if not line: continue` |
| Index is sequential among valid items (not line number) | `utils.py:102` | `index=len(items)` — skips invalid lines in numbering |

### Excel Reading Rules

| Rule | Location | Details |
|------|----------|---------|
| Single cell: bypass SpecialCells | `excel_handler.py:265-295` | `if selection.Count == 1` — process directly to avoid hang |
| Filter-aware: xlCellTypeVisible (12) | `excel_handler.py:299` | `selection.SpecialCells(12)` — only visible cells |
| String values: strip commas, try float | `excel_handler.py:274-278, 357-361` | `float(value.replace(',', ''))` |
| Multi-column warning | `excel_handler.py:415-418` | `if len(columns_used) > 1` — warns but still reads |
| Bulk read with cell-by-cell fallback | `excel_handler.py:306-412` | Try `area.Value` first; on any exception, fall back to iterating cells |
| Excel color format: RGB (not BGR) | `excel_handler.py:451` | `excel_color = r + (g * 256) + (b * 256 * 256)` — this is actually correct RGB→long conversion for Excel |

### Finalization Rules

| Rule | Location | Details |
|------|----------|---------|
| Colors cycle through 20 predefined colors | `utils.py:16-37, 53-63` | `HIGHLIGHT_COLORS[index % 20]` — wraps around |
| Finalized items get "✓" prefix | `find_tab.py:924` | `list_item.setText(f"✓ {original_text}")` |
| Selected items get "▶" prefix | `find_tab.py:776` | `list_item.setText(f"▶ {current_text}")` |
| Combos containing any finalized item are removed | `find_tab.py:858-894` | Set intersection check: `combo_indices & finalized_indices` |

### Validation Rules

| Rule | Location | Details |
|------|----------|---------|
| Target must be valid float | `utils.py:177-180` | `float(target)` with ValueError catch |
| Buffer must be non-negative float | `utils.py:183-189` | `buffer_val < 0` → error |
| Min/Max size must be positive integer | `utils.py:192-210` | `int(...)`, must be ≥ 1 |
| Min size ≤ Max size | `utils.py:223-225` | Cross-validation check |
| Max results must be positive integer | `utils.py:213-220` | `int(...)`, must be ≥ 1 |

### UI Constants

| Constant | Value | Location | Purpose |
|----------|-------|----------|---------|
| Default max results | 25 | `find_tab.py:319` | Changed from 100 in commit 0ce191e |
| Default min combo size | 1 | `find_tab.py:295` | Allows single-item matches |
| Default max combo size | 10 | `find_tab.py:306` | Reasonable upper bound |
| Default buffer | 0 | `find_tab.py:288` | Exact matches only by default |
| Excel monitor interval | 5000ms | `main_window.py:110` | Check Excel connection every 5 seconds |
| Batch size | 20 | `find_tab.py:49` | Max combos per UI update batch |
| Batch interval | 0.1s | `find_tab.py:50` | Max time between UI update batches |
| Selection highlight color | (255, 165, 0) | `find_tab.py:37` | Orange |
| Finalized text color | (150, 150, 150) | `find_tab.py:38` | Grey |
| Disabled combo color | (240, 240, 240) | `find_tab.py:39` | Light grey |
| Window minimum size | 1000×700 | `main_window.py:32` | Minimum window dimensions |
| Splitter initial sizes | [280, 380, 320] | `find_tab.py:228` | Left, middle, right panel widths |
| Highlight colors count | 20 | `utils.py:16-37` | Number of unique finalization colors |
| SpinBox range for sizes | 1-100 | `find_tab.py:294,305` | Min/max spinbox bounds |
| SpinBox range for results | 1-10000 | `find_tab.py:318` | Max results spinbox bound |

---

## SECTION 6: EDGE CASES AND DEFENSIVE CODE

### Handled Edge Cases

| Edge Case | Handling | Location | Rating |
|-----------|----------|----------|--------|
| Single cell Excel selection | Direct processing, bypasses SpecialCells to avoid hang | `excel_handler.py:265-295` | ROBUST |
| Filtered Excel data | SpecialCells(12) for visible cells only | `excel_handler.py:299` | ROBUST |
| Bulk read failure | Falls back to cell-by-cell reading per area | `excel_handler.py:380-412` | ROBUST |
| Multiple COM value formats | Handles: single value, tuple, tuple of tuples, non-tuple | `excel_handler.py:318-341` | ROBUST |
| Excel not running | Graceful error message, returns empty | `excel_handler.py:53-67` | ROBUST |
| pywin32 not installed | `try/except ImportError` with `EXCEL_AVAILABLE` flag | `excel_handler.py:16-22` | ROBUST |
| Excel closed unexpectedly | Timer-based monitor with warning dialog | `main_window.py:106-119` | MODERATE |
| Empty input text | Status message "Please enter some numbers" | `find_tab.py:463-464` | MODERATE |
| All items finalized | Status message "All numbers are finalized" | `find_tab.py:576-578` | MODERATE |
| No valid combinations possible | `quick_check_possible` returns early | `solver.py:204-245` | MODERATE |
| All sizes mathematically impossible | Smart bounds `no_solution` flag | `find_tab.py:623-630` | ROBUST |
| Invalid parameter values | `validate_parameters` returns error list | `utils.py:153-230` | MODERATE |
| Negative buffer input | Converted to absolute value in solver | `solver.py:49` | MODERATE |
| Header row clicked in results | Check `item.data(Qt.UserRole) is None`, clear selection | `find_tab.py:699,719` | ROBUST |
| String values in Excel cells | Try to strip commas and parse as float | `excel_handler.py:357-361` | MODERATE |
| COM thread initialization | `pythoncom.CoInitialize()` called in connect | `excel_handler.py:58` | MODERATE |
| Multiple column selection | Warning message but still processes all | `excel_handler.py:415-418` | MODERATE |
| Close with unsaved Excel highlights | Save/Discard/Cancel dialog | `main_window.py:127-147` | ROBUST |
| Close with finalized combos | Confirmation dialog | `main_window.py:150-161` | MODERATE |
| Negative numbers in data | Algorithm handles correctly; smart bounds less effective but correct | `solver.py:47-98` | MODERATE |

### Error Handling Patterns

| Pattern | Locations | Rating |
|---------|-----------|--------|
| Bare `except:` (catches all, silences) | `excel_handler.py:78,109,134,161,219,237,294,300,379,411,456,500` | FRAGILE — hides errors silently |
| `except com_error:` (specific) | `excel_handler.py:65` | ROBUST |
| `except (TypeError, ValueError):` | `excel_handler.py:290,376,409` | MODERATE |
| `except ValueError:` | `utils.py:108,148,179-220` | MODERATE |
| `except Exception as e:` with user message | `excel_handler.py:422` | MODERATE |

**Overall Rating: MODERATE** — Good handling of Excel COM edge cases (the hardest part), but excessive use of bare `except:` clauses throughout `excel_handler.py` that silently swallow errors, making debugging difficult.

---

## SECTION 7: EDGE CASES NOT HANDLED (GAPS AND VULNERABILITIES)

### CRITICAL: Combination Removal Bug After Finalization

**Location:** `find_tab.py:876-894` (`_filter_combination_list`)

**The Bug:** When removing invalid combinations after finalization, the code uses indices from `combo_list` (which contains only combinations) as row indices for `list_widget.takeItem()` (which contains both headers AND combinations). Because the list widget has header items interspersed among combination items, these indices don't match.

**Concrete Trace:**
Suppose the list widget contains:
```
Row 0: Header "═══ 2 Numbers (3 found) ═══"
Row 1: combo_0 (items with indices {0, 1})
Row 2: combo_1 (items with indices {2, 3})
Row 3: combo_2 (items with indices {0, 4})
Row 4: Header "═══ 3 Numbers (2 found) ═══"
Row 5: combo_3 (items with indices {1, 2, 5})
Row 6: combo_4 (items with indices {3, 4, 6})
```
And `combo_list = [combo_0, combo_1, combo_2, combo_3, combo_4]`.

If the user finalizes items with indices {0, 1}, the code identifies combos at `combo_list` indices [0, 2, 3] for removal. Removing in reverse:
- `list_widget.takeItem(3)` removes **Row 3 (combo_2)** but `combo_list.pop(3)` removes **combo_3** — mismatch!
- `list_widget.takeItem(2)` removes **combo_1** (now at row 2) — which is valid and shouldn't be removed!
- `list_widget.takeItem(0)` removes **the Header** — destroys the UI structure!

**Impact:** After finalization: (1) invalid combinations remain visible, (2) valid combinations disappear, (3) headers get deleted, (4) the widget display becomes desynchronized from internal data. This bug is always triggered when headers are present (i.e., results contain more than one combination of a given size).

**Why it hasn't been catastrophic in practice:** Users likely start a new search after finalization (which clears everything), masking the corrupted state. Also, finalization of smaller result sets (where fewer headers exist) may produce less visible corruption.

### Missing Validations

| Gap | Risk | Severity |
|-----|------|----------|
| No validation that target is non-zero | Target of 0 with buffer of 0 finds all single-zero items; not harmful but surprising | LOW |
| `max_size_input.setMaximum(len(items))` resets on each load | If user loads fewer items than before, the max size shrinks; subsequent loads of more items restore it; not a bug, just surprising UX | LOW |
| No check for extremely large search spaces | User can set max_size=100 on 100 items = C(100,50) ≈ 10^29 combinations; the app will effectively hang forever | MEDIUM |
| Thousand-separator ambiguity in comma mode | "1,000" → split to "1" and "000" → 1.0 and 0.0 instead of 1000.0 | MEDIUM |
| No column letter validation in Excel | If the column letter calculation produces an invalid Excel address, the COM call fails silently | LOW |

### Silent Logic Bugs

| Bug | Location | Details |
|-----|----------|---------|
| Header count drift | `find_tab.py:946-990` | After `_filter_combination_list` removes items, the header "count found" text is not updated. E.g., header says "(5 found)" but only 3 remain visible after finalization. The header row index in `exact_headers`/`approx_headers` also becomes stale after row removal. |
| `_get_combo_avg_position` exists but unused | `find_tab.py:92-96` | Static method defined but never called — remnant of abandoned sorting feature. Not harmful, just dead code. |

### Scenario-Based Failure Analysis

| Scenario | Impact |
|----------|--------|
| **Working directory is different from dev directory** | No impact — the app doesn't use file paths or relative imports beyond the project structure. Uses COM automation which is path-independent. |
| **Run by a different OS user** | No impact — no user-specific paths, no registry access, no file permissions issues. COM access to Excel is per-user but transparent. |
| **Files on a network drive that disconnects** | Minimal impact — the app doesn't write to disk during normal operation. Excel saving depends on the Excel process, not this app. Only risk: Python module loading fails if drive disconnects before launch. |
| **Run for the second time on the same data** | No issues — each run is stateless. Previous finalization colors remain in Excel from the last run (they are cell formatting changes), but the app doesn't read or conflict with them. |
| **Two instances run simultaneously** | Both would try `win32.GetActiveObject("Excel.Application")` and get the same Excel instance. Both could read/write to the same cells simultaneously. No locking mechanism. Result: potential overlapping highlights and confusing behavior. | MEDIUM risk. |
| **Different OS (Mac/Linux)** | Excel COM features fail gracefully (`EXCEL_AVAILABLE = False`). Manual input still works. Surprisingly, the code handles this via the try/except import guard. However, the app has never been tested on non-Windows. |
| **Python 3.13+ (as indicated by .pyc filenames)** | Currently running on Python 3.13 (per `.pyc` files). `QDesktopWidget` used in `main_window.py:43` is deprecated since Qt 5.11 and may be removed in PyQt6. No immediate issue but a migration concern. |
| **PyQt6 migration** | Would require changing: `exec_()` → `exec()`, `QDesktopWidget` → `QScreen`, possibly signal/slot syntax changes. Not backwards compatible. |

### Performance Concerns

| Concern | Details |
|---------|---------|
| Combinatorial explosion | 200 items, max_size=10: C(200,10) ≈ 2.2×10^16 combinations to check. Even with smart bounds, large searches are impractical. The stop button is the only mitigation. |
| COM calls per cell for highlighting | `highlight_items` makes one COM call per cell (via `highlight_cell`). For large finalized combos, this could be slow. Batch range highlighting would be faster. |
| Source list re-rendering | `_update_source_colors` iterates ALL source items and updates styling. For 200+ items, this triggers 200+ widget repaints. |

---

## SECTION 8: DESIGN DECISIONS (INFERRED)

### Technology Choices

| Decision | Reasoning | Confidence |
|----------|-----------|------------|
| Python + PyQt5 over Electron/web | Simpler stack for a single-user desktop tool; good Excel COM integration via pywin32; AI-assisted development is easier in Python | HIGH — stated in HANDOVER.md |
| pywin32 over openpyxl | Need to interact with a RUNNING Excel instance (not just read files). COM automation is the only way to read live selections and apply highlights in real-time. openpyxl works on files, not live instances. | HIGH — this is a hard constraint |
| `itertools.combinations` over recursive/DP approach | Negative numbers make pruned backtracking and DP approaches ineffective. `itertools.combinations` is simple, correct, and handles negatives naturally. | HIGH — explicitly documented in HANDOVER.md rejected approaches |
| QThread over multiprocessing | Simpler threading model; solver runs in background thread, communicates via Qt signals. Multiprocessing was discussed but rejected by user (concern about CPU usage). | HIGH — stated in HANDOVER.md |
| Singleton ExcelHandler | Single Excel connection shared across all tabs. Avoids multiple COM connections which could conflict. | MEDIUM — reasonable pattern for single-user app |
| Text marker ("▶"/"✓") over pure background color | `setBackground()` on QListWidgetItem didn't reliably work for selection highlighting. Text change forces Qt to redraw the item. Background color works for finalized items because the text also changes (adding "✓"). | HIGH — documented as Bug 5 in HANDOVER.md, workaround confirmed |
| Real-time combo display over collect-then-sort | Storing all results for end-of-search sorting caused memory access violations and crashes on large datasets. Real-time display via generator/signals is safer. | HIGH — documented as Bug 6 in HANDOVER.md |
| Batch UI updates (20 items / 0.1s) | One-at-a-time updates caused "Not Responding" on large result sets. Batching with `setUpdatesEnabled(False/True)` prevents excessive repaints. | HIGH — implemented in final commit after discussion |
| Absolute buffer over percentage | User preference for simplicity. Percentage-based tolerance would add complexity (relative to what base?) without clear benefit for the use case. | HIGH — stated in HANDOVER.md |
| No undo | User explicitly declined. Will handle corrections manually or via Excel. | HIGH — stated in HANDOVER.md |
| No session persistence | User explicitly declined as "too complicated". | HIGH — stated in HANDOVER.md |

### Pattern Consistency Analysis

| Pattern | Files Following | Files Breaking | Notes |
|---------|----------------|----------------|-------|
| Graceful win32com import | `excel_handler.py` ✓ | — | Only file that imports win32com; consistently uses try/except with `EXCEL_AVAILABLE` flag |
| Bare `except:` for COM calls | `excel_handler.py` — all COM methods | — | Consistent within this file but FRAGILE as a pattern; should catch specific exceptions |
| Signal/slot for cross-tab communication | `find_tab.py → summary_tab.py` ✓, `settings_tab.py → find_tab.py` ✓ | — | Consistent use of pyqtSignal for tab-to-tab communication |
| QBrush wrapper for setBackground | `_clear_highlights()` ✓, `_update_source_colors()` ✓ | `_highlight_combination()` uses `QColor` directly (no QBrush) | **INCONSISTENCY** — `_highlight_combination` (line 770-771) uses `QColor(255,165,0)` and `QColor(0,0,0)` directly, while `_clear_highlights` (line 802-811) wraps in `QBrush()`. The finalized styling in `_highlight_combination` also mixes: line 763 uses `QColor` directly, while `_clear_highlights` line 802 uses `QBrush(QColor(...))`. Both approaches work in PyQt5, but the inconsistency likely traces back to the Bug 5 debugging attempts. |
| Color references | `styles.py` defines COLORS dict | `find_tab.py` hardcodes `"#343a40"`, `"#ff6b6b"`, `"#51cf66"`, `"#e67700"` etc. | **INCONSISTENCY** — These hex values match the COLORS dict entries but are hardcoded rather than referencing `get_color()` (which is itself dead code). The styles module's color system is not used outside the stylesheet. |

### Re-export Pattern

`ui/__init__.py` re-exports `MainWindow`, `FindTab`, `SummaryTab`, `SettingsTab`, and `apply_styles` via `__all__`. This means any new UI component would need to be added to both the import line and `__all__`. Currently, only `MainWindow` and `apply_styles` are imported from `ui` package externally (by `main.py`). The other exports (`FindTab`, `SummaryTab`, `SettingsTab`) are used only within the `ui` package via relative imports.

---

## SECTION 9: DEPENDENCIES AND ENVIRONMENT

### External Libraries

| Package | Version | Purpose | Used By |
|---------|---------|---------|---------|
| PyQt5 | >=5.15.0 | Desktop GUI framework (widgets, layouts, signals/slots, threading) | All UI files, main.py |
| pywin32 | >=304 | Windows COM automation for Excel (win32com, pywintypes, pythoncom) | excel_handler.py |

### Standard Library Dependencies

| Module | Used By | Purpose |
|--------|---------|---------|
| `sys` | main.py | Command-line args, exit |
| `time` | find_tab.py | Search timing, batch intervals |
| `itertools.combinations` | solver.py | Core algorithm |
| `math.comb` | solver.py | Combination count estimation |
| `dataclasses` | models.py | Data class decorators |
| `typing` | All files | Type hints |
| `enum.Enum` | models.py | InputMode, ItemSource enums |
| `pythoncom` | excel_handler.py | COM thread initialization |

### Python Version

- **Required:** Python 3.8+ (uses dataclasses, f-strings, walrus operator is NOT used)
- **Currently running on:** Python 3.13 (per `.cpython-313.pyc` files)

### OS-Specific Dependencies

| Dependency | Required? | Fallback |
|------------|-----------|----------|
| Windows | For Excel COM features | Manual input still works on Mac/Linux |
| Excel (running instance) | For Excel features | Graceful "Not Available" message |
| pywin32 (Windows only) | For Excel features | `EXCEL_AVAILABLE = False`, features disabled |

### Setup on Fresh Machine

```bash
# 1. Install Python 3.8+
# 2. Install dependencies
pip install PyQt5 pywin32

# 3. (Windows only) Run COM registration for pywin32
python -c "import win32com"

# 4. Run the app
python main.py
```

> **Missing from requirements.txt:** `pythoncom` is imported directly in `excel_handler.py:11` but is part of the `pywin32` package, not a separate dependency. This is correct but could confuse someone reading the imports.

---

## SECTION 10: UI/INTERFACE DOCUMENTATION

### Window Structure

The application is a single window (`MainWindow`, 1000×700 minimum) with three tabs:

```
┌─────────────────────────────────────────────────────────────────┐
│ CombiMatch - Invoice Reconciliation                        [─□×]│
├─────────────────────────────────────────────────────────────────┤
│ [🔍 Find Combinations] [📋 Summary] [⚙️ Settings]              │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│              (Tab Content Area)                                 │
│                                                                 │
├─────────────────────────────────────────────────────────────────┤
│ Ready                                                           │
└─────────────────────────────────────────────────────────────────┘
```

### Tab 1: Find Combinations (Main Workspace)

Three-panel horizontal splitter layout:

```
┌──────────────┬──────────────────────┬─────────────────┐
│ LEFT PANEL   │ MIDDLE PANEL         │ RIGHT PANEL     │
│ (280px)      │ (380px)              │ (320px)         │
│              │                      │                 │
│ ┌──────────┐ │ ┌──────────────────┐ │ ┌─────────────┐│
│ │Input Mode│ │ │ Exact Matches    │ │ │  Selected   ││
│ │[Dropdown]│ │ │ 0 exact matches  │ │ │ Combination ││
│ └──────────┘ │ │                  │ │ │             ││
│ ┌──────────┐ │ │ ═══ 2 Numbers ══│ │ │ Sum: -      ││
│ │Numbers   │ │ │ 1. 10, 40 = 50  │ │ │ Diff: -     ││
│ │Input     │ │ │ 2. 20, 30 = 50  │ │ │ Items: -    ││
│ │          │ │ │                  │ │ │ Values: -   ││
│ │(TextEdit)│ │ └──────────────────┘ │ └─────────────┘│
│ │          │ │ ┌──────────────────┐ │ ┌─────────────┐│
│ └──────────┘ │ │ Approx Matches   │ │ │Source Nums  ││
│ [Load Nums]  │ │ 0 approximate    │ │ │0 loaded     ││
│              │ │                  │ │ │             ││
│ ┌──────────┐ │ │ ═══ 3 Numbers ══│ │ │ #1: 100.50  ││
│ │Search    │ │ │ 1. 10,20,21=51  │ │ │ #2: 200.00  ││
│ │Parameters│ │ │    (+1.00)       │ │ │ ▶ #3: 50.00 ││
│ │          │ │ │                  │ │ │ ✓ #4: 30.00 ││
│ │Target:___│ │ └──────────────────┘ │ │             ││
│ │Buffer:_0_│ │                      │ │             ││
│ │Min: [1]  │ │ [✓ Finalize Selected]│ └─────────────┘│
│ │Max: [10] │ │                      │                 │
│ │Results:25│ │                      │                 │
│ └──────────┘ │                      │                 │
│              │                      │                 │
│ [🔍 Find    ]│                      │                 │
│ [⏹ Stop    ] │                      │                 │
│ Progress...  │                      │                 │
│ Status msg   │                      │                 │
└──────────────┴──────────────────────┴─────────────────┘
```

**Buttons:**
- **Load Numbers:** Parses text input into NumberItems
- **📊 Grab from Excel:** (visible only in Excel mode) Reads Excel selection
- **🔍 Find Combinations:** Starts search with current parameters
- **⏹ Stop Search:** (replaces Find during search) Stops the solver
- **✓ Finalize Selected:** (enabled when a combo is selected) Locks in the selection

**Input Fields:**
- **Input Mode dropdown:** Line Separated / Comma Separated / Excel Selection
- **Numbers Input (TextEdit):** Free-form number entry (hidden in Excel mode)
- **Target Sum:** The value to match
- **Buffer (±):** Tolerance range (default: 0)
- **Min Numbers (SpinBox):** Minimum combo size (default: 1, range: 1-100)
- **Max Numbers (SpinBox):** Maximum combo size (default: 10, range: 1-100, auto-adjusted to item count)
- **Max Results (SpinBox):** Result limit (default: 25, range: 1-10000)

**Dynamic Elements:**
- Orange hint labels next to Min/Max spinboxes showing viable bounds
- Red "No solution possible" labels when no size is viable
- Progress label during search: "Checking 3-number combinations... (5,000 checked)"
- Status label: green for success, red for errors

### Tab 2: Summary

Scrollable list of colored cards, one per finalized combination:

```
┌─────────────────────────────────────┐
│ Finalized Combinations              │
│ 3 combinations                      │
│ Total Sum: 1,500.00                 │
│                                     │
│ ┌─────────────────────────────────┐ │
│ │ [Light Blue Background]         │ │
│ │ #1 - Light Blue                 │ │
│ │ Sum: 500.00 (+0.00)             │ │
│ │ Items: 2                        │ │
│ │ Values: 200.00, 300.00          │ │
│ │ Rows: R5, R12                   │ │
│ └─────────────────────────────────┘ │
│ ┌─────────────────────────────────┐ │
│ │ [Light Green Background]        │ │
│ │ #2 - Light Green                │ │
│ │ Sum: 1,000.00 (+0.00)           │ │
│ │ Items: 3                        │ │
│ │ Values: 100.00, 400.00, 500.00  │ │
│ └─────────────────────────────────┘ │
│                                     │
└─────────────────────────────────────┘
```

### Tab 3: Settings

```
┌──────────────────────────────────────┐
│ ┌──────────────────────────────────┐ │
│ │ Excel Connection                 │ │
│ │ Status: Connected (green)        │ │
│ │ Active: Book1.xlsx (Sheet1)      │ │
│ │                                  │ │
│ │ Workbook: [Book1.xlsx     ▼] [↻] │ │
│ │ Sheet:    [Sheet1         ▼]     │ │
│ │                                  │ │
│ │ [Connect to Excel] [Disconnect]  │ │
│ └──────────────────────────────────┘ │
│ ┌──────────────────────────────────┐ │
│ │ Highlight Colors                 │ │
│ │ Colors cycle through palette:    │ │
│ │ [■][■][■][■][■][■][■][■][■][■]  │ │
│ │ [■][■][■][■][■][■][■][■][■][■]  │ │
│ └──────────────────────────────────┘ │
│ ┌──────────────────────────────────┐ │
│ │ About                            │ │
│ │ CombiMatch - Invoice Recon Tool  │ │
│ │ Features: ...                    │ │
│ │ Tip: Start by loading numbers... │ │
│ └──────────────────────────────────┘ │
└──────────────────────────────────────┘
```

### Theme

Soft blue-grey palette defined in `styles.py`:
- Background: `#f5f6f8` (very light grey)
- Surface: `#ffffff` (white)
- Primary: `#5c7cfa` (medium blue)
- Text: `#343a40` (dark grey)
- Font: Segoe UI, 10pt
- Rounded corners (4px) on inputs, buttons, tabs
- No harsh borders — subtle `#dee2e6` borders

### Typical User Workflow

1. Open the app
2. (Optional) Go to Settings → Connect to Excel → Select workbook/sheet
3. Load numbers: paste into text area OR grab from Excel
4. Enter target sum and buffer
5. Click "Find Combinations"
6. Review results — click combos to preview in info panel
7. Click "Finalize Selected" to lock in matches
8. Repeat steps 4-7 for additional targets
9. Switch to Summary tab to review all finalized
10. Close app → prompted to save Excel with highlights

---

## SECTION 11: RECONSTRUCTION BLUEPRINT

### 1. Foundation

```
1. Create Python virtual environment (3.8+)
2. Install PyQt5 and pywin32
3. Create project structure:
   combimatch/
   ├── main.py
   ├── models.py
   ├── solver.py
   ├── excel_handler.py
   ├── utils.py
   ├── requirements.txt
   └── ui/
       ├── __init__.py
       ├── main_window.py
       ├── find_tab.py
       ├── summary_tab.py
       ├── settings_tab.py
       └── styles.py
```

### 2. Build Order (Features depend on each other in this order)

```
Layer 1 (No dependencies):
  ├── models.py          — Data classes first
  ├── utils.py           — Parsing, colors, validation
  └── styles.py          — Theme/stylesheet

Layer 2 (Depends on Layer 1):
  ├── solver.py          — Algorithm (uses models)
  └── excel_handler.py   — Excel COM (uses models)

Layer 3 (Depends on Layers 1-2):
  ├── find_tab.py        — Main workspace (uses everything)
  ├── summary_tab.py     — Summary view (uses models, utils)
  └── settings_tab.py    — Settings (uses excel_handler, utils)

Layer 4 (Depends on Layer 3):
  ├── main_window.py     — Tab container (uses all tabs)
  └── main.py            — Entry point (uses main_window, styles)
```

### 3. Feature Priority

| Priority | Feature | Complexity | Notes |
|----------|---------|------------|-------|
| **Essential** | F1, F2: Manual input parsing | SIMPLE | Foundation for all searches |
| **Essential** | F4: Subset sum solver | MODERATE | Core algorithm; test with small sets first |
| **Essential** | F5: Exact/approx split | SIMPLE | Just a threshold check |
| **Essential** | F9: Finalization with colors | MODERATE | Critical for workflow |
| **Essential** | F7: Selection highlighting | SIMPLE | Visual feedback |
| **Essential** | F12: Stop search | SIMPLE | Safety valve for long searches |
| **High** | F3: Excel reading | COMPLEX | COM automation with many edge cases |
| **High** | F14: Smart bounds | MODERATE | Major performance improvement |
| **High** | F18: Batch UI updates | MODERATE | Prevents UI freezing |
| **Medium** | F6: Size-grouped headers | MODERATE | Visual organization |
| **Medium** | F8: Info panel | SIMPLE | Nice-to-have detail view |
| **Medium** | F10: Workbook/sheet selection | MODERATE | Excel management |
| **Medium** | F11: Summary tab | SIMPLE | Overview of finalized combos |
| **Low** | F13: Progress tracking | SIMPLE | Progress feedback |
| **Low** | F15: Auto-switch to Excel mode | SIMPLE | UX convenience |
| **Low** | F16: Close-event save handling | SIMPLE | Data protection |
| **Low** | F17: Excel connection monitor | SIMPLE | Robustness |

### 4. Known Improvements for Rebuild

> Per the constraint in the prompt: I have checked Section 8 before making these recommendations. None of these contradict established design decisions.

| Improvement | Addresses | Approach |
|-------------|-----------|----------|
| **Fix combo removal bug** | Section 7 Critical Bug | Instead of using combo_list index as widget row, iterate the widget, check each item's `data(Qt.UserRole)`, and remove items that overlap with finalized indices. This avoids the header-offset problem entirely. |
| **Replace bare `except:` with specific exceptions** | Section 6 FRAGILE rating | Use `except com_error:` for COM calls, `except (AttributeError, TypeError):` for property access. Log unexpected exceptions for debugging. |
| **Consistent QBrush/QColor usage** | Section 8 pattern inconsistency | Standardize on QBrush wrapping everywhere (it's the more explicit approach). |
| **Use COLORS dict from styles.py** | Section 8 color hardcoding | Import and use `COLORS` dict instead of hardcoded hex strings in find_tab.py. |
| **Add search space warning** | Section 7 performance | Before starting search, calculate `estimate_combinations()` and warn user if the search space exceeds a threshold (e.g., 10 million). |
| **Remove dead code** | Section 2 dead code inventory | Delete `format_number`, `_column_number`, `clear_highlight`, `find_combinations` (non-generator), `CombinationResult`, `_get_combo_avg_position`, `check_excel_connection`, `get_color`, and the orphan `placeholder` attribute. |
| **Fix thousand-separator handling in comma mode** | Section 7 validation gap | Detect and handle patterns like "1,000" in comma-separated mode. Consider using a smarter split that distinguishes delimiter commas from thousand-separator commas. |
| **Update header counts after removal** | Section 7 header count drift | After `_filter_combination_list`, update the header text to reflect the remaining count, or remove empty headers. |
| **Replace QDesktopWidget** | Section 7 deprecation | Use `QScreen` and `QGuiApplication.primaryScreen()` instead of the deprecated `QDesktopWidget`. |

### 5. Estimated Complexity (AI-Assisted Rebuild)

| Component | Complexity | Estimate |
|-----------|------------|----------|
| models.py | SIMPLE | < 30 min |
| utils.py | SIMPLE | < 30 min |
| styles.py | SIMPLE | < 30 min |
| solver.py (with smart bounds) | MODERATE | 1-2 hours |
| excel_handler.py (COM edge cases) | COMPLEX | 2-4 hours |
| find_tab.py (main workspace) | COMPLEX | 3-5 hours |
| summary_tab.py | SIMPLE | < 1 hour |
| settings_tab.py | MODERATE | 1-2 hours |
| main_window.py | SIMPLE | < 30 min |
| main.py | SIMPLE | < 15 min |
| **Total** | | **~10-15 hours AI-assisted** |

### 6. Suggested Claude Code Skills

| Skill | What It Would Encode |
|-------|---------------------|
| `pyqt-list-widget-manager` | Pattern for managing QListWidget with headers, data roles, safe removal, and batch updates. Avoids the header-index-mismatch class of bugs. |
| `win32com-excel-reader` | Excel COM automation pattern: filter-aware reading, bulk value extraction, format detection (single/row/col/block), single-cell special handling, fallback strategies. |
| `subset-sum-solver` | Configurable combination finder with smart bounds, generator-based yielding, progress callbacks, and stop mechanism. Reusable for any "find items that sum to target" problem. |
| `pyqt-batch-thread` | QThread pattern with batched signal emission, configurable batch size and interval, progress reporting, and stop capability. |

---

## SECTION 12: CHANGELOG ARCHAEOLOGY (FROM GIT HISTORY)

All 8 commits are on the `main` branch (no feature branches exist).

### 12A. Complete Development Timeline

#### Phase 1: Initial Implementation (2025-12-02)

**Commit 1: `031fe0d` — 2025-12-02 21:44:04 +0530 (main)**
*"Initial Commit"*

- **What changed:** Created entire project from scratch — 14 files, 2,927 lines. All core files, UI files, README, gitignore, requirements.
- **Why:** This was the result of a complete AI-assisted development conversation. The project was built in a single session and committed as a whole.
- **What it reveals:**
  - Initial max_results default was **100** (later changed to 25)
  - `SolverThread` emitted single combos (`result_found` signal per combo, no batching)
  - No stop button, no progress indicator
  - No split results (single combined list)
  - No info panel
  - No sheet selection
  - No single-cell handling
  - No bulk reading (cell-by-cell iteration)
  - Selection highlight attempted with `QColor("transparent")` and `QColor` directly (no QBrush)
  - `_handle_selection` took (combo_list, row) — used row as index into combo_list

**Commit 2: `c1e0cf3` — 2025-12-02 22:02:26 +0530 (main)**
*"Enhance README and UI: Add features for split results view, selected combination info panel, and Excel sheet selection."*

- **What changed:** 451 additions, 68 deletions across 4 files
  - `README.md`: Added split results, info panel, sheet selection documentation
  - `excel_handler.py`: Added `get_sheets()`, `select_sheet()`, `get_active_sheet_name()` methods (+63 lines)
  - `ui/find_tab.py`: Major rework — added `SelectedComboInfoPanel` class, split exact/approx lists, added `_on_exact_selected`/`_on_approx_selected` handlers, mutual selection clearing between lists (+370/-68 lines)
  - `ui/settings_tab.py`: Added sheet dropdown, `_on_workbook_changed` (populates sheets), `_on_sheet_changed` (+58 lines)
- **Why:** These were features discussed in the original conversation but not in the initial commit. Likely implemented right after the first commit.
- **What it reveals:** The initial commit was a snapshot mid-conversation; the split view, info panel, and sheet selection were implemented immediately after.

#### Phase 2: Excel Performance Fixes (2025-12-03, morning)

**Commit 3: `80b4d19` — 2025-12-03 07:49:40 +0530 (main)**
*"Optimize read_selection method for bulk reading of Excel cells"*

- **What changed:** 56 additions, 34 deletions in `excel_handler.py` only
  - Replaced cell-by-cell iteration (`for cell in area`) with bulk `area.Value` reading
  - Added format detection: single cell = raw value, multi-cell = tuple of tuples
  - Preserved row/column tracking using area start position + offsets
- **Why:** Cell-by-cell COM calls froze the UI when reading filtered Excel data (Bug 1 in HANDOVER.md)
- **What it reveals:** Real-world testing with filtered data immediately exposed the performance bottleneck. This is the most impactful performance fix in the project.

**Commit 4: `9d83f5b` — 2025-12-03 08:21:09 +0530 (main)**
*"Enhance ExcelHandler: Improve bulk reading... Add fallback... Add progress tracking and stop functionality"*

- **What changed:** 143 additions, 58 deletions across 2 files
  - `excel_handler.py`: Added more robust format detection (single row, single column cases), wrapped bulk read in try/except with cell-by-cell fallback
  - `ui/find_tab.py`: Added Stop button (hidden by default), progress label, `_on_stop_clicked`, `_on_progress`, changed Find button behavior (hide/show instead of enable/disable)
- **Why:** Bug 2 (COM error with different range formats) and Bug 4 (search button stuck) from HANDOVER.md
- **What it reveals:** The original bulk read didn't handle all COM value formats. The iterative refinement shows real-world testing with various Excel selection patterns (single column, single row, block, non-contiguous).

**Commit 5: `28c5d76` — 2025-12-03 08:51:14 +0530 (main)**
*"Add direct handling for single cell selection and improve background/foreground color handling with QBrush"*

- **What changed:** 50 additions, 19 deletions across 2 files
  - `excel_handler.py`: Added single-cell detection (`selection.Count == 1`) with direct processing, bypassing SpecialCells
  - `ui/find_tab.py`: Changed all `QColor` → `QBrush(QColor(...))` for background/foreground, changed `QColor("transparent")` → `QBrush()` for clearing
- **Why:** Bug 3 (SpecialCells hangs on single cell) and Bug 5 (orange highlight not showing) from HANDOVER.md
- **What it reveals:** SpecialCells(12) has an undocumented behavior where it hangs on single-cell selections. The QBrush change was an attempt to fix the highlight issue (it didn't fully work — the "▶" marker workaround came later).

#### Phase 3: Refactoring and Headers (2025-12-03, afternoon/evening)

**Commit 6: `7825168` — 2025-12-03 16:22:45 +0530 (main)**
*"Refactor FindTab: Extract combination handling logic into separate methods"*

- **What changed:** 105 additions, 49 deletions in `find_tab.py` only
  - Added `_get_combo_avg_position()` static method (never called — remnant of abandoned sorting)
  - Extracted `_add_combo_to_list()` method with size headers
  - Changed `_handle_selection(combo_list, row)` to `_handle_combo_selection(combo)` — now receives combo directly from item data, not by indexing into combo_list
  - Added "▶ " text marker for selection highlighting (the workaround for Bug 5)
  - Changed clearing to remove "▶ " prefix
  - Mixed QColor/QBrush usage — reverted some QBrush back to QColor for selection highlighting
- **Why:** Major refactoring to add size-grouped headers required changing how selection works (can't use list index as combo_list index anymore). Also implemented the "▶" marker workaround.
- **What it reveals:**
  - The `_get_combo_avg_position` method was added for a position-based sorting feature that was immediately abandoned (Bug 6 in HANDOVER.md — memory crash)
  - The selection handling was fundamentally changed: before, `_handle_selection` used row index to look up combos; after, `_handle_combo_selection` receives the combo directly from `Qt.UserRole` data. This fixed the header-induced offset problem for SELECTION, but the same class of bug was introduced in `_filter_combination_list` (which still uses combo_list indices as widget row indices).

**Commit 7: `0ce191e` — 2025-12-03 21:37:21 +0530 (main)**
*"Introduce exact and approximate headers management for combination lists"*

- **What changed:** 22 additions, 10 deletions in `find_tab.py` only
  - Added `exact_headers` and `approx_headers` dicts to track header row positions
  - Changed header to include count: `"═══ 2 Numbers (1 found) ═══"`
  - Added logic to update header count when more combos of same size are found
  - Changed default max_results from 100 to **25**
  - Added headers dict clearing when starting new search
- **Why:** User requested counts in headers ("(5 found)") and lower default max results
- **What it reveals:** The header tracking system was added as an afterthought. The stored row indices become stale after any row insertion/deletion, which contributes to the removal bug (Section 7).

#### Phase 4: Optimization and UX Polish (2026-02-20)

**Commit 8: `3910f52` — 2026-02-20 09:41:14 +0530 (main)**
*"Add smart combination size bounds, batch UI updates, and UX improvements"*
*Co-Authored-By: Claude Opus 4.6*

- **What changed:** 280 additions, 40 deletions across 5 files. This is the most recent commit, made ~2.5 months after the initial development burst.
  - `solver.py`: Added `compute_smart_bounds()` function (+64 lines), integrated into `SubsetSumSolver.__init__` (+20 lines), `bounds_info` dict
  - `ui/find_tab.py`: Major changes:
    - `SolverThread` changed from per-combo signals to **batch emission** (20 items or 0.1s interval)
    - Added `_on_combo_batch` replacing `_on_combo_found`
    - Added `setUpdatesEnabled(False/True)` around batch processing
    - Added `_update_bounds_hints()` with live viable min/max labels
    - Added `switch_to_excel_mode()` for auto-switching
    - Added total sum display in source count label
    - Added auto-adjust max size spinbox to item count
    - Added elapsed time display in search results
    - Added "Try increasing the buffer" hint on no results
    - Added available count/total after finalization
    - Added `time` import and `_search_start_time` tracking
  - `ui/main_window.py`: Connected `excel_connected` signal to `switch_to_excel_mode`
  - `ui/settings_tab.py`: Added `excel_connected = pyqtSignal()`, emitted on connection
  - `.gitignore`: Added `.claude/` to ignore list
- **Why:** This was a new development session (note the 2.5 month gap). The focus was on performance optimization (smart bounds, batching) and UX polish (hints, totals, timing, auto-switch).
- **What it reveals:**
  - The Co-Authored-By tag confirms this was an AI-assisted session with Claude
  - Despite HANDOVER.md saying batch updates were "decided not to implement", they WERE implemented here — suggesting the user changed their mind in the new session
  - The smart bounds feature is the most algorithmically sophisticated addition, showing the project has grown from a simple utility to a more optimized tool
  - The `.claude/settings.local.json` contains permissions for a `csolver/combisolver.exe` that doesn't exist in the repo, suggesting experimentation with a C-compiled solver (possibly abandoned)

### 12B. File Evolution Map

| File | Created | Evolution |
|------|---------|-----------|
| `main.py` | Commit 1 | **Never changed.** Stable entry point. |
| `models.py` | Commit 1 | **Never changed.** Data classes established once and stable. |
| `solver.py` | Commit 1 | Changed in Commit 8: Added `compute_smart_bounds()` and integration into solver init. Import changed to add `Tuple`. |
| `excel_handler.py` | Commit 1 | Changed in Commits 2, 3, 4, 5. Most edited file proportionally. Evolution: initial cell-by-cell → bulk read → robust format handling → fallback strategy → single-cell detection. Each change addressed a real-world bug. |
| `utils.py` | Commit 1 | **Never changed.** Parsing and color utilities established once. |
| `ui/__init__.py` | Commit 1 | **Never changed.** |
| `ui/find_tab.py` | Commit 1 | Changed in Commits 2, 4, 5, 6, 7, 8. **Most edited file overall (6 commits).** Grew from 578 lines to ~1073 lines. Evolution: single list → split lists → headers → header tracking → batch updates + smart bounds UI. |
| `ui/main_window.py` | Commit 1 | Changed in Commit 8 only: added auto-switch signal connection. |
| `ui/settings_tab.py` | Commit 1 | Changed in Commits 2, 8. Added sheet selection (Commit 2) and excel_connected signal (Commit 8). |
| `ui/styles.py` | Commit 1 | **Never changed.** |
| `ui/summary_tab.py` | Commit 1 | **Never changed.** |
| `README.md` | Commit 1 | Changed in Commit 2 only: added split results, info panel, sheet selection docs. |
| `.gitignore` | Commit 1 | Changed in Commit 8: added `.claude/` |
| `requirements.txt` | Commit 1 | **Never changed.** |

### 12C. Logic and Rule Evolution

| Rule/Logic | First Appeared | Changes | Interpretation |
|------------|---------------|---------|----------------|
| Buffer tolerance (absolute ±) | Commit 1 | Never changed | Foundational from day one |
| Exact threshold (< 0.001) | Commit 2 | Never changed | Established with split view |
| Max results default | Commit 1 (100) → Commit 7 (25) | Reduced at user request | 100 was too many results for practical use |
| Bulk value reading | Commit 3 (basic) → Commit 4 (robust) | Added format detection and fallback | Real-world testing exposed COM quirks |
| Single cell handling | Commit 5 | Never changed | Discovered through production use (SpecialCells hang) |
| Size headers | Commit 6 (no count) → Commit 7 (with count) | Added "(N found)" count | User feedback on header usefulness |
| Smart bounds | Commit 8 | New feature | Performance optimization for large datasets |
| Batch UI updates | Commit 8 | New feature (reversed earlier "no" decision) | UI freezing was more annoying than expected |
| Progress callback frequency | Commit 1 (1000 iter) | Never changed | Established early, adequate throughout |
| Color palette (20 colors) | Commit 1 | Never changed | 20 soft colors sufficient for typical use |

### 12D. Edge Cases Discovered in Production

| Commit | Bug | What Broke | How Fixed | Unhandled Similar Cases |
|--------|-----|------------|-----------|------------------------|
| 3 (80b4d19) | UI freeze on filtered Excel | Cell-by-cell COM calls on filtered data | Bulk `area.Value` reading | Non-contiguous selections with thousands of areas |
| 4 (9d83f5b) | COM error on different formats | `area.Value` returns inconsistent types | Format detection + cell-by-cell fallback | Very large areas that exceed COM memory limits |
| 4 (9d83f5b) | Search button stuck | No way to stop long searches | Added Stop button + progress | Thread doesn't terminate — it only checks a flag periodically |
| 5 (28c5d76) | App hang on single cell | SpecialCells(12) hangs on Count=1 | Direct processing for single cell | Other SpecialCells edge cases (entire row/column selected) |
| 5 (28c5d76) | Orange highlight not showing | QColor background not rendering | Changed to QBrush (partially worked) | Root cause never fully identified |
| 6 (7825168) | Memory crash on sort | Storing all combos for sorting | Reverted to real-time display | No position-based sorting is available (tradeoff accepted) |

### 12E. Abandoned Approaches and Reversals

| Approach | Evidence | Outcome |
|----------|----------|---------|
| **Position-based sorting** | `_get_combo_avg_position` added in Commit 6, never called | Abandoned after memory crash. Dead method remains in codebase. |
| **QColor for background highlighting** | Commit 1 used `QColor` directly, Commit 5 changed to `QBrush`, Commit 6 reverted selection back to `QColor` | Mixed approach remains. The "▶" text marker was the real fix. |
| **"No batch updates" decision** | HANDOVER.md says "User decided not to implement" | Reversed in Commit 8 — batch updates WERE implemented. The user apparently changed their mind. |
| **C-compiled solver** | `.claude/settings.local.json` has permissions for `csolver/combisolver.exe` | No trace in repo. Either experimented and abandoned in an uncommitted session, or planned for future. |
| **Result counting via `total_checked`** | Commit 1: `total_checked` variable in `SolverThread.run()`, passed to `finished_signal` | `total_checked` was never actually computed — always emitted as 0. The variable was removed from the run method in Commit 4 but the signal still accepts two ints. |
| **Combined results list** | Commit 1 had a single `combinations` list and `result_list` widget | Split into exact/approx in Commit 2. |

### 12F. Dependency Evolution

| Change | Commit | Details |
|--------|--------|---------|
| Initial dependencies | Commit 1 | PyQt5>=5.15.0, pywin32>=304 |
| **No changes since** | — | Dependencies have been stable throughout all 8 commits |

**Dependencies used but not in requirements.txt:** None — `pythoncom` is part of pywin32. All standard library modules don't need listing.

**Note from `.claude/settings.local.json`:** Permissions reference `gcc` and `combisolver.exe`, suggesting experimentation with a C solver. If this were implemented, it would add a C compiler as a build dependency. Currently no C code exists in the repo.

---

## Self-Consistency Review (Completed)

1. **Cross-section contradictions:** None found. Section 8 notes batch updates were initially rejected but later implemented (Commit 8). Section 11 does not recommend reverting this — the implementation is correct and beneficial. Section 3's concrete traces are consistent with Section 5's business rules.

2. **Features verified:** All 18 features have been traced with concrete inputs. The combination removal bug (F9 → Section 7) was discovered during the concrete trace of finalization.

3. **Shared state verified:**
   - `NumberItem` objects are shared by reference between `self.items`, `combo_list` entries, and `source_list` item data roles. Mutation of `is_finalized` and `finalized_color` propagates correctly.
   - `ExcelHandler` singleton is accessed consistently via `get_excel_handler()` across `find_tab.py`, `settings_tab.py`, and `main_window.py`.
   - `exact_headers`/`approx_headers` dicts store row indices that become stale after row removal — documented in Section 7.

4. **Pattern consistency verified:** QBrush/QColor inconsistency documented in Section 8. Color hardcoding inconsistency documented in Section 8. Bare except pattern documented in Section 6.

5. **Corrections:** [Corrected during self-review — HANDOVER.md states "Batch UI updates: User decided not to implement after discussion" but Section 12A reveals they WERE implemented in Commit 8 (2026-02-20). The HANDOVER.md was written before this final commit and is therefore stale on this point.]

[Corrected during self-review — FILE_INDEX.md states find_tab.py is "~650 lines" but it is actually ~1073 lines in the current version. The FILE_INDEX.md was written before Commits 7 and 8 which added significant code.]

[Corrected during self-review — HANDOVER.md lists "Sorting by average position: Was planned, caused crash, user said 'let this be'" as not implemented, but the dead method `_get_combo_avg_position` still exists in the codebase as a remnant.]
