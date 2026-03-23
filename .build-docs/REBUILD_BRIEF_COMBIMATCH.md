# Rebuild Brief: COMBIMATCH

# Date: 2026-03-22

# Author: Adv. Sidhant Goyal

---

### Top 3 Priorities
1. **C Solver (.dll)** — Accurate, bug-free solver for 5,000+ invoices. Speed is the bonus; correctness is non-negotiable.
2. **Seed Numbers (Must Include / Partial Lock-In)** — Pin known invoices, solver finds what completes the target.
3. **Bug Fixes** — Combo removal bug, stop button that actually stops, orange highlight, smart bounds with warnings. Rock-solid foundation.

### Non-Negotiables (beyond the Top 3)
- **Excel filtered cell reading** — SpecialCells(12), single-cell bypass, bulk read with fallback. This is the backbone of daily operations. If this breaks, nothing else matters.

### Zero-Defect Requirement
**The solver — specifically the C (.dll) solver.** The core algorithm that finds combinations must produce correct results every single time. A bug here means wrong reconciliation, wrong ITC claims, wrong numbers presented to clients. Unacceptable.

---

## 1. WHAT THIS TOOL DOES (in my own words)

CombiMatch is a general-purpose subset-sum solver — given a list of numbers and a target amount, it finds which combinations of numbers add up to that target (exactly or approximately). It is not locked to one specific use case. The most common use in our office is GSTR-2B to GSTR-3B reconciliation: when I know the total ITC claimed in 3B but need to identify exactly which 2B invoices were taken to arrive at that figure. But it's equally useful for any scenario where you have a list of numbers (could be 5,000 invoices or more) and need to figure out which ones make up a given sum. The approximate match feature is critical — it's a lifesaver when exact matches don't exist. Both the number of clients needing this and the volume of data are growing exponentially, so **handling large datasets efficiently is a non-negotiable priority for the rebuild** — the original was built for smaller sets and the rebuild must scale well beyond that.

---

## 2. WHO USES THIS AND HOW

- **Primary users:** Articles and interns at the office. They are non-coders — they run from Excel formulas, let alone Python. The tool must behave like a regular desktop application with zero technical setup required.
- **Technical comfort:** Absolute minimum. The app is wrapped with `pythonw` and launched via a desktop shortcut — users never see a terminal. This must remain the case in the rebuild.
- **Frequency of use:** Daily. This is a workhorse tool, not an occasional utility.
- **Environment:** Office runs an nComputing setup — multiple nodes sharing one PC's CPU and RAM, with a separate file server storing project files. The builder also uses a personal laptop. Only one instance of CombiMatch runs at a time to avoid choking the shared machine, but the app can use moderate resources since other nodes mainly run Excel and Chrome.
- **Primary workflow (Excel path):** User opens an Excel sheet with the data → selects cells (linear rows/columns or filtered rows/columns) → grabs numbers into CombiMatch → enters target sum, buffer, and size range → runs the search → reviews combinations → finalizes the best match → cells get highlighted in Excel → repeats with remaining numbers for the next target → continues until the reconciliation is complete or desired result is reached.

---

## 3. FEATURE DECISIONS

| Feature | Audit Ref | Decision | Notes |
| --- | --- | --- | --- |
| Manual Number Input (Line-Separated) | F1 | KEEP AS-IS | Office culture uses accounting format in Excel. Copy-pasting brings commas — thousand-separator parsing is essential. Every input mode has its use case. |
| Manual Number Input (Comma-Separated) | F2 | KEEP BUT IMPROVE | Used rarely but worth keeping. **Change delimiter from comma (,) to semicolon (;)** to eliminate the thousand-separator ambiguity bug. "1,000" will always parse as one thousand, semicolons separate values. |
| Excel Selection Reading | F3 | KEEP AS-IS | Running perfectly after 5 commits of hardening. Filtered cell reading (SpecialCells), single-cell bypass, bulk read with fallback — all critical. **ACTION ITEM (post-rebuild):** Extract this pattern into a reusable SKILL.md for other apps. |
| Combination Finding (Subset Sum Solver) | F4 | KEEP BUT IMPROVE | **Two-layer solver architecture:** (1) C-compiled solver (.dll) as primary — with smarter pruning/branch-and-bound algorithms for 5,000+ invoice datasets. (2) Python `itertools` solver as silent fallback if .dll is missing or compilation fails. **Auto-compilation:** App checks for solver.dll at startup → if missing and compiler available, auto-builds from solver.c → if no compiler, falls back to Python silently. The .dll lives on the file server; article nodes never need a compiler. C source stays in the repo for maintainability via Claude Code. |
| Exact vs. Approximate Results Split | F5 | KEEP AS-IS | Two separate lists preferred. Approximate match is a lifesaver — not secondary to exact. |
| Size-Grouped Headers with Counts | F6 | KEEP BUT IMPROVE | Grouping by size is very helpful — sometimes need more numbers, sometimes less. **Two improvements:** (1) Fix stale header count bug — count must update after finalization removes combos. (2) Sort approximate matches by closeness to target (ascending absolute difference) within each size group, irrespective of positive/negative sign. |
| Combination Selection and Highlighting | F7 | KEEP BUT IMPROVE | **[AUDIT CORRECTION: Audit stated orange highlight "partially worked" with QBrush. Per builder's experience, the orange background never actually worked — only bold text and "▶" marker are visible.]** Three improvements: (1) Actually fix the orange background highlight so it renders. (2) Keep bold + "▶" marker. (3) Add scroll bar position markers (like Microsoft Edge's PDF search minimap) on the source numbers list, so matching numbers are visible even in long lists of 500+ numbers without scrolling. |
| Selected Combination Info Panel | F8 | KEEP AS-IS | Shows sum, difference, count, and values. Works fine for now. Open to revisit if new ideas come up during or after rebuild. |
| Finalization with Color Assignment | F9 | KEEP BUT IMPROVE | Core workflow — color assignment, "✓" marker, Excel cell highlighting, removing finalized numbers from results — all stay. **Critical bug fix required:** The combo-removal logic after finalization is broken (widget-row/combo-list index mismatch due to headers). Must be fixed so iterative finalization works reliably. For GSTR-2B reco, each match has its own conditions so user restarts fresh. But for other use cases, the iterative flow (finalize → search remaining → finalize again) must work seamlessly. |
| Excel Workbook and Sheet Selection | F10 | KEEP BUT IMPROVE | Settings tab flow is smooth. **Two improvements:** (1) Bigger display boxes for workbook and sheet name dropdowns — names get cut off currently. (2) Auto-switch Find tab to "Excel Selection" mode when workbook is selected for the first time or when refresh is hit. If already in Excel Selection mode, no action needed. **[AUDIT CORRECTION: Audit stated F15 (Auto-Switch to Excel Mode) works on Excel connection. Per builder's experience, this auto-switch is not functioning in practice.]** |
| Summary Tab (Finalized Combinations View) | F11 | KEEP AS-IS | Good-to-have feature to see all the workings from start to finish. Colored cards with sum, difference, count, values, and row references — all useful for review. |
| Stop Search | F12 | KEEP BUT IMPROVE | Stop button exists but doesn't stop promptly — solver only checks the stop flag every 1000 iterations, which can feel unresponsive on large datasets. **Fix:** Make stop checks much more frequent (every 50-100 iterations). For the C solver (.dll), build in a proper interrupt mechanism so stop means stop immediately. |
| Progress Tracking | F13 | KEEP BUT IMPROVE | Progress label is good — keeps the user engaged and shows the computer is working. **Improvement:** Progress callbacks must work with both C solver (.dll) and Python fallback. The C solver must periodically report progress back to the Python UI so the app doesn't look frozen during large searches. |
| Smart Combination Size Bounds | F14 | KEEP BUT IMPROVE | Smart bounds hints are showing but not actually working in current build. **Improvements:** (1) Smart bounds must actually function — calculate and skip impossible sizes. (2) If user sets bounds outside viable range, show a warning dialog with two options: "Change Bounds" or "Proceed Anyway" — don't silently proceed or block them. (3) Increase max size limit beyond 100 — C solver (.dll) can handle larger sizes for 5,000+ invoice datasets. |
| Auto-Switch to Excel Mode | F15 | COVERED UNDER F10 | Not working in current build. Fix merged into F10's improvement — auto-switch to Excel Selection mode on workbook selection or refresh. |
| Close-Event Save Handling | F16 | KEEP AS-IS | Save prompt on close is necessary, not annoying. Keep the Excel save prompt and close confirmation as-is. |
| Excel Connection Monitor | F17 | KEEP BUT IMPROVE | Background monitor is useful — catches accidental closures. **Improvement:** When Excel disconnects, the app should remember which workbook was open and offer a dialog: "Excel was closed. Would you like to reopen [filename]?" If yes, app relaunches Excel with that file and reconnects automatically. |
| Batch UI Updates | F18 | KEEP BUT IMPROVE | Current batching (20 items / 0.1s) helps but UI still doesn't feel smooth. **Improvements:** (1) Smoother result streaming — list shouldn't visibly "jump" during updates. (2) App must stay fully responsive during search — scrolling, clicking, resizing without lag. (3) Better visual feedback (proper progress bar instead of just text label). With C solver handling computation and better UI update management, the search experience should feel smooth. |

---

## 4. NEW FEATURES TO ADD

Features that didn't exist in the original but you want in the rebuild:

| # | New Feature | What it should do | Why I need it | Priority |
| --- | --- | --- | --- | --- |
| N1 | Session Save/Restore | Save progress (loaded numbers, finalized combinations, search state) so if the app closes — crash, restart, lunch break — user can resume where they left off. | Articles do long reconciliation sessions. A crash or restart shouldn't mean starting over. | MUST HAVE |
| N2 | Seed Numbers (Must Include / Partial Lock-In) | Before searching, user pins certain numbers as "must include." Solver finds what remaining numbers complete the target. Remaining target = original target minus sum of pinned numbers (including negatives). Buffer applies on remaining target. | In GSTR-2B reco, you often already know some invoices are definitely part of the match. Currently must include them manually in your head. | MUST HAVE |
| N3 | Quick Reset / Re-grab | Clear current loaded data and grab fresh from Excel without restarting the app. A "Clear All Numbers" button. | If user grabs wrong column or wrong sheet, they currently have to restart the entire app. | MUST HAVE |
| N4 | ~~Reverse Mode (Possible Totals)~~ | MOVED TO PARKING LOT — technically infeasible at scale (2^5000 subsets for 5,000 invoices). Needs scoping before implementation. | — | PARKED |
| N5 | Custom Labels for Finalized Combinations | On finalization, popup asks for an optional label (e.g., "Dec Payment", "TDS Adj"). Label appears on Summary Tab cards. In Excel, label is written as a Shift+F2 note on the first cell of the combination, along with the total sum with buffer difference and cross-sheet/workbook references if applicable. | With 6+ finalized combinations, colors alone aren't enough to remember which match was for which purpose. Notes in Excel provide full traceability. | MUST HAVE |
| N6 | Unmatched Numbers View | After matching is done, show/highlight all numbers NOT part of any finalized combination. For Excel source: offer a "Mark Unmatched" button that highlights remaining cells in a single distinct color (e.g., light grey). Match colors always override unmatched color on subsequent runs. For Line/Comma source: include in the export. | Unmatched leftovers are just as important as matches — that's what you chase the client about. | MUST HAVE |
| N7 | Undo Last Finalization | Strictly last-in-first-out only. Undo releases those numbers back into the available pool, removes the color from Excel cells, and restores the combination to results. No picking and choosing — only the most recent finalization can be undone. | Articles sometimes finalize the wrong combination by mistake. Currently requires full restart. | MUST HAVE |
| N8 | Multi-Workbook/Sheet Grab | Settings tab shows workbooks with checkboxes. Each checked workbook expands to show its sheets with checkboxes. User goes to each checked sheet in Excel, makes their cell selection (Excel remembers selections per sheet). Clicks "Grab from Excel" once — app loops through all checked workbook/sheet combinations, reads each remembered selection, loads everything into one combined list. Each number tagged with source (workbook + sheet + cell). Finalization highlights cells across all relevant workbooks/sheets with same color and cross-reference notes. | Real reconciliations span multiple workbooks and sheets (e.g., Apr-Jun in one file, Jul-Sep in another). Currently limited to one sheet at a time. | MUST HAVE |
| N9 | Sort Source List | Sort loaded numbers by value — ascending or descending — to spot outliers and mentally narrow down before searching. | With 2,000+ numbers, hard to get a sense of the data without sorting. | NICE TO HAVE |
| N10 | Reconciliation Summary Stats | Display matched total, unmatched total, percentage reconciled, count of matched vs unmatched. | Gives a quick overview of how complete the reconciliation is. Absorbed into the report export as well. | MUST HAVE |
| N11 | One-Click Reconciliation Report Export | Full Excel export serving as a complete audit trail / working paper: (1) Source info — which workbooks, sheets, cell ranges were grabbed. (2) Run log — each search performed with target, buffer, size range, results count. (3) Actions taken — each finalization with combination selected, label given, cells highlighted across which sheets/workbooks. (4) Matched summary — all finalized combinations with colors, labels, totals, differences, cell references with full workbook/sheet paths. (5) Unmatched list — all numbers not part of any combination, with source references. (6) Summary stats — total matched, unmatched, percentage reconciled, number of runs, combinations finalized. Must be detailed enough that someone reading it a year later can reconstruct exactly what happened. "Extra info is no harm, no info is." | Need a proper working paper for client files, record-keeping, and traceability. The colored Excel alone isn't enough. | MUST HAVE |
| N12 | Multi-Target Batch Mode | Enter multiple target amounts upfront. App searches for matches for each target sequentially against the same number list, presenting results grouped by target. Seed numbers (#N2) can be used within this. | Multiple payments against the same invoice list is common. Less clicking, faster workflow. | NICE TO HAVE |

---

## 5. BUSINESS RULES — KEEP / CHANGE / ADD

Reference audit Section 5.

### Rules to KEEP exactly as-is:

- **Buffer is absolute, not percentage-based** — (solver.py:49). Absolute buffer is the correct approach.
- **Exact match threshold < 0.001** — (find_tab.py:663). Safety net for floating-point rounding. Paisa-level differences are irrelevant in practice ("not a big deal in Supreme Court either"). Keep as-is.
- **Values rounded to 2 decimal places on load** — (models.py:50-52). All real-world data is 2 decimal places. Keep as-is.
- **Finalized items excluded from subsequent searches** — (solver.py:47). Once matched, a number shouldn't appear in another combination. Core workflow logic.
- **Min combo size floor = 1, max combo size ceiling = item count** — (solver.py:50-51). Keep as-is.
- **Progress reported every 1000 iterations** — (solver.py:138,181). Keep but ensure it works with C solver too (covered under F13).
- **Colors cycle through 20 predefined colors, wrapping around** — (utils.py:16-37, 53-63). 20 combinations is more than enough for any real session.
- **Finalized items get "✓" prefix, selected items get "▶" prefix** — (find_tab.py:924, 776). Keep as-is.
- **Combos containing any finalized item are removed from results** — (find_tab.py:858-894). Keep logic, but fix the removal bug (covered under F9).
- **Line-separated mode: commas stripped as thousand separators, empty lines skipped** — (utils.py:93-98). Essential for accounting-format numbers. Keep as-is.
- **Single cell Excel selection: bypass SpecialCells to avoid hang** — (excel_handler.py:265-295). Keep as-is.
- **Filter-aware reading: SpecialCells(12) for visible cells only** — (excel_handler.py:299). Critical feature. Keep as-is.
- **String values in Excel: strip commas and parse as float** — (excel_handler.py:357-361). Keep as-is.
- **Multi-column selection: warn but still process** — (excel_handler.py:415-418). Keep the warning.
- **Bulk read with cell-by-cell fallback** — (excel_handler.py:306-412). Battle-tested pattern. Keep as-is.
- **Target must be valid number, buffer must be non-negative, min ≤ max, max results must be positive** — (utils.py:153-230). All standard validation. Keep as-is.
- **Max results default = 25** — (find_tab.py:319). Practical limit. Keep as-is.

### Rules to CHANGE:

- **Old rule:** Comma-separated mode uses comma (,) as delimiter — (utils.py:114-150)
  **New rule:** Use semicolon (;) as delimiter instead. Commas always treated as thousand separators in all modes.
  **Why:** Eliminates the "1,000" → "1" and "000" ambiguity bug.

- **Old rule:** Search order is always smallest combination size first — (solver.py:111,162)
  **New rule:** Search order becomes a user-selectable option: "Smallest first" or "Largest first."
  **Why:** Depends on the case — sometimes small matches are needed, sometimes user wants to fill to max possible.

- **Old rule:** Target and buffer fields reject pasted values with thousand-separator commas (e.g., "1,00,000")
  **New rule:** Target and buffer input fields must strip commas before parsing, just like the number input does. Pasting "1,00,000" should work as 100000.
  **Why:** Users copy-paste target amounts from Excel in accounting format. Current behaviour blocks them unnecessarily.

### Rules to ADD (not in original):

- **Match color always overrides unmatched color** — When a previously "unmatched" (grey) cell gets matched in a subsequent run, the match color replaces the grey. (Supports N6: Unmatched Numbers View)
- **Cross-reference notes on finalized cells** — Shift+F2 notes on highlighted Excel cells must include: label, total sum with buffer difference, and cross-sheet/workbook references if applicable. (Supports N5 and N8)
- **Accumulative grab** — "Grab from Excel" adds to existing list, doesn't replace. Each number tagged with source workbook + sheet + cell. "Clear All Numbers" resets. (Supports N8: Multi-Workbook/Sheet Grab)
- **Session auto-save** — App periodically saves state so progress survives crashes. (Supports N1: Session Save/Restore)
- **Undo is strictly last-in-first-out** — Only the most recent finalization can be undone. No picking and choosing. (Supports N7: Undo Last Finalization)

---

## 6. EDGE CASES — LESSONS FROM THE AUDIT

Reference audit Sections 6 and 7.

### Must handle (were handled in original — keep):

- **Single cell Excel selection bypass** — SpecialCells(12) hangs on single cell. Happens when articles forget to select a range. Bypass must stay.
- **Filtered Excel data (SpecialCells visible cells only)** — Working perfectly after 5 commits of hardening. Critical for daily filtered-sheet workflow.
- **Bulk read failure → cell-by-cell fallback** — Robust fallback pattern. Keep.
- **Multiple COM value formats** — Handles single value, tuple, tuple of tuples, non-tuple. Keep.
- **Excel not running / pywin32 not installed** — Graceful error messages. Keep.
- **Excel closed unexpectedly** — Monitor with warning. Enhanced in rebuild with auto-reopen offer (F17).
- **Empty input text** — Status message. Keep.
- **All items finalized** — Status message. Keep.
- **No valid combinations possible** — Quick check returns early. Keep.
- **All sizes mathematically impossible** — Smart bounds no_solution flag. Keep.
- **Invalid parameter values** — Validation with error list. Keep.
- **Negative buffer input** — Converted to absolute. Keep.
- **Header row clicked in results** — Checks UserRole data, clears selection. Keep.
- **String values in Excel cells** — Strip commas, parse as float. Keep.
- **Multi-column selection** — Warning but still processes. Keep the warning.
- **Close with unsaved Excel highlights** — Save/Discard/Cancel dialog. Keep.
- **Negative numbers in data** — Algorithm handles correctly. Keep. Important for credit notes in GST work.

### Must handle (were NOT handled — gaps that need fixing):

- **Extremely large search spaces** — User can set impossible search parameters (max_size=100 on 5,000 items) and app hangs, requiring Task Manager kill. **Fix:** Warn before starting with estimated search size (covered under F14), AND ensure stop button actually works immediately (covered under F12). No more Task Manager kills.
- **Combination removal bug after finalization** — Widget-row/combo-list index mismatch due to headers. **Fix:** Covered under F9.
- **Header count drift after finalization** — Header says "(5 found)" but fewer remain. **Fix:** Update header counts after removal, or remove empty headers.
- **Two instances accessing same Excel sheets** — **Fix:** Enforce single instance only. If someone tries to open a second instance, show "CombiMatch is already running" and block it. Avoids resource choking on nComputing setup and eliminates all multi-instance conflicts.
- **Target/buffer fields rejecting comma-separated values** — **Fix:** Covered under Section 5 rule change (strip commas before parsing).
- **Thousand-separator ambiguity in comma mode** — **Fix:** Covered under F2 (semicolon delimiter).
- **QDesktopWidget deprecation** — **Fix:** Use modern replacement (QScreen) to future-proof. Technical decision, not user-facing.

### Not worth handling (over-engineering for my use case):

- **Mac/Linux compatibility** — Windows only. "Windows is king." No cross-platform overhead.
- **Dirty/malformed data (spaces in numbers, text mixed with numbers)** — Office culture enforces clean data. App can assume reasonably clean input.
- **PyQt6 migration** — Not needed now. Rebuild on PyQt5, future-proof where easy (QScreen instead of QDesktopWidget) but don't target PyQt6.

---

## 7. DESIGN PREFERENCES

- **Framework:** PyQt5. Same as all other tools in the office. Great Windows support, Excel COM compatibility, large library of components. No reason to switch.
- **Interface style:** Needs a redesign — the old three-panel layout won't hold 12 new features. Design decision delegated to the builder ("you're the designer, design it"). Must be clean, intuitive for non-coders (articles and interns). Key UI improvement: **collapsible/accordion-style size groups** in the results lists — click a "2 Numbers (5 found)" header to expand/collapse that group. Enables fast navigation when large datasets generate hundreds of combinations.
- **Output format:** Excel (.xlsx) for all exports — reconciliation report, unmatched list, everything. No PDF needed.
- **File handling:** Accumulative Excel grab (multi-workbook/sheet via checkboxes). Desktop shortcut launch via pythonw — no terminal ever visible.
- **Error display:** **Pop-ups where attention is demanded** (e.g., warnings before large searches, finalization confirmations, Excel disconnection). **Inline messages where user can proceed without intervention** (e.g., parsing errors, progress updates, status messages). Simple rule: interrupt only when action is required.
- **Windows only:** No Mac/Linux support. No cross-platform overhead.
- **Theme:** Current soft blue-grey palette is fine as a starting point. Redesign as needed for the new feature set.

---

## 8. WHAT I LEARNED (from reading the audit)

The builder did not read the audit in detail — the audit was used as a reference by the interviewer (AI) to drive focused questions during the brief-filling process. The real lessons come from months of daily use:

1. **Plan before building.** The original CombiMatch was born from a single AI conversation — functional but unplanned. This rebuild is the opposite: a full brief, every feature debated, every rule validated, before a single line of code is written. That IS the lesson.

2. **The tool outgrew its original scope.** What started as a simple "find matching numbers" utility is now a full reconciliation workstation — multi-workbook, session persistence, audit trail reports, seed numbers. The rebuild must be architected for this larger scope from day one, not patched onto a small-tool foundation.

3. **The guiding principle for the rebuild: "No bugs, no assumptions, no fault."** If the developer (Claude Code) has to choose between two approaches, the one that is more explicit, more tested, and less assumption-dependent wins. Every feature must work correctly. Silent failures, swallowed errors, and "it probably works" are not acceptable.

4. **Users are non-coders.** Every design decision must pass the article/intern test: "Would someone who runs from Excel formulas be able to use this without asking for help?" If the answer is no, redesign it.

5. **Scale is the new reality.** The tool was built for small datasets. The practice is growing exponentially — more clients, more data, up to 5,000+ invoices. The C solver, smarter algorithms, and smooth UI are not nice-to-haves — they're the reason for the rebuild.

---

## 9. PARKING LOT

| Idea | When it came up | Priority | Status |
| --- | --- | --- | --- |
| Extract Excel COM filtered-cell-reading pattern into reusable SKILL.md | Section 3 (F3 discussion) | Post-rebuild | Planned |
| Reverse Mode (Possible Totals) — show achievable target sums from a list of numbers. Needs scoping: infeasible at 5,000 invoices without constraints (range limit or max combo size). | Section 4 (N4 discussion) | Future | Parked — needs feasibility scoping |

---
