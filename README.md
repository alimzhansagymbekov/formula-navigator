# Formula Navigator — Excel Add-in

An Arixcel-class formula explorer and model auditor for Excel. Every formula component shows its **calculated value**, the **branch actually taken** by IF/CHOOSE is marked, **lookups resolve to their target cells**, precedents and dependents trace across sheets, SUMIFS-style formulas explain **exactly which records matched**, and a one-click audit flags hardcodes, broken references, inconsistent formulas and other model risks. Built on Office.js (TypeScript + React); runs inside Excel's sandbox.

> **Status:** version 16. Built and used on desktop Excel for Microsoft 365 on Windows. Mac and Excel on the web are untested. Some features need ExcelApi 1.9+ (any current 365 build); a banner tells you if your Excel is missing something.

---

## What it does

### Explore — the selected formula as a live tree
- Every component row shows its **computed value** (`= 282,048`); cell references show their **displayed text** in the cell's own number format (`12.5%`, `1,234`, dates). Values sit in **one right-aligned column**.
- **Argument captions on the argument's own row** (`then: B3`, `criteria range: C2:C7`) for 96 functions; variadic ones (SUMIFS, COUNTIFS, LET, SWITCH…) are captioned at any length.
- The **branch actually taken** by IF / IFS / CHOOSE / SWITCH / IFERROR / IFNA is marked in green; untaken *results* are dimmed (the condition never is). A **← source** chip shows where the value came from: `← B3` when the branch is a cell (followed through nested IFs and lookups), `← else` when it is an expression (click to focus it).
- **INDEX, VLOOKUP, HLOOKUP, XLOOKUP, OFFSET, INDIRECT, MATCH are evaluated** — a green **→ Sheet!B14** chip (same column as ← source) jumps to the exact cell the lookup landed on. Long addresses truncate with "…"; hover shows the full one.
- For **SUMIF / SUMIFS / COUNTIF(S) / AVERAGEIF(S) / FILTER** and **SUMPRODUCT logical filters**, the tree shows **which records matched**, each criterion's own hit count, and the first matches as links — plus a **match walk**: `4 of 48 records matched  ◂ 2 / 4 ▸  D22`. Arrow onto it, then → / ← step through every match (up to 2,000) and Excel jumps to each; ◂ ▸ work with the mouse. For SUMIFS & co the walk lands on the **summed value**, not the criterion cell. Full-column ranges (`A:A`) are clamped to the used range.
- **Formula line highlight:** the part of the formula the focused row stands for is highlighted in the formula line above the tabs (which now has a full-width row and scrolls to the highlight).
- **Follow** (on by default, remembered per computer): arrowing onto a reference moves Excel there (~0.1 s debounce) while the explored cell stays the anchor. Off: browse only — Enter or click jumps. A click is always exactly one jump.
- **Values: as shown / exact** (next to Follow; also drives Trace and the floating window): exact reveals digits the number format hides, in grey — `1,249`**`.6372`**, `12.5`**`372`**`%`, `(1,234)` → `-1,234`**`.48`**. Dates and text are untouched; hover always shows both (`shown: 1,250 · exact: 1,249.6372`). Exact values are read in the same Excel round trip as the displayed text.
- The **▸ on any reference unfolds that cell's own formula in place** (cycle-safe, multi-level).
- **Defined names and table references are navigable** (V16): click / Enter / Follow goes where the name really points — a cell (▸ unfolds its formula) or a range (selected).
- **Multi-cell trace:** select a range and the tree builds precedents/dependents for every cell at once.
- A **yellow "← Back" banner** appears when you've navigated away; click it, reach it with ↑, or press **Backspace / Ctrl+Backspace** to return to the explored cell.

### Trace — precedents / dependents across sheets
- Multi-level tree with values, range-membership chips and cycle flags; refreshes when you edit. Follow and Values: exact work here too.
- A reference to **one cell** is always listed, even when that cell is empty — marked `(empty)` (also for a name pointing at one cell and each sheet of `Jan:Mar!B5`).
- **A range is one row** (V16): `Data!D2:D500 · 499 cells`, folded; → lists its filled cells (first 25, then `… N more`); Enter selects the range. `INDEX(Data!D2:D500, MATCH(A1, Data!C2:C500, 0))` shows 3 rows instead of 51. Full columns show their filled count; names show `Rates (Data!D2:D500)`.
- **Found by lookup** (V16): for the traced formula itself, the cells its INDEX / VLOOKUP / HLOOKUP / XLOOKUP / OFFSET / INDIRECT actually returned are listed first (`→ D15 · found by INDEX`) — including INDIRECT/OFFSET sources a static trace can't see. Shown only when the add-in's evaluator reproduces the value Excel holds for the formula; otherwise nothing is shown. Lookups in an IF branch not taken are skipped.
- An **impact line** sums the whole-model blast radius ("feeds 1,284 formulas across 9 sheets").
- A **Tree | Graph toggle** switches to an interactive dependency graph.

### Map — sheet scan + workbook audit
- **Scan active sheet** classifies every cell (input / formula / output / cross-sheet / label) and flags model risks — see *Audit checks* below.
- **Audit workbook** adds whole-model checks: circular reference chains, formulas reading hidden sheets, unused named ranges.
- All findings appear in **one keyboard-navigable list** (↑/↓, Excel follows; one jump per click). Filter by status or severity. Repeated inconsistent cells in a row collapse into one entry (×N).
- Each finding carries a **review status** (Open → Checked → Question → Error) saved inside the workbook. Statuses survive model edits — a finding is identified by what it is and where it is, not by its description text.
- **Unused named ranges** can be deleted: the panel asks for confirmation, keeps a backup of the definition, and offers **Restore**.
- **Export to sheet** writes the same list — category, location, the cell's formula, detail and review status (colour-coded) — to a "Formula Navigator Audit" worksheet.

### Find — pins and search
- Pinned cells (★, saved inside the workbook), plus search by reference (`Model!B7`), function (`XLOOKUP`) or named range, with dependency counts.

### Floating windows (Ctrl+Shift+C / D)
Compact, keyboard-first precedent/dependent windows over the grid. One window is reused; drag it to a corner and it stays there for the session. Range, `… N more` and found rows work the same as in the Trace tab. The footer shows **the formula that links the focused cell to its parent, with the linking reference highlighted** (`in E2: =IF(B2>1000,[B3],B4)`), plus the focused cell's own formula or value. With Values: exact the header shows an `exact` badge.

### Live updates — when is Rescan needed?
Edits are picked up as you type: formulas, values, pastes (very large pastes rescan that sheet automatically), sheets added or deleted. **Numbers stay current:** after every edit (~0.25 s) and every recalculation — F9, live data feeds, volatile functions — (at most once a second) the pane re-reads everything it shows: the explored formula, unfolded formulas, the Trace tab, and the open floating window (in place, view kept). Excel sends no event for cells that merely recalculate, which is why V15 and earlier could show stale numbers downstream of an edit.

Press **Rescan** only after changing defined names or tables, renaming a sheet, or when the button turns orange (sheets added/removed in a model with 3-D references).

### Crash safety net
Each tab, the formula line and the floating window are isolated: a rendering error shows a small **"This view hit an error — Reload"** card instead of a blank pane, other tabs keep working, and the error text is listed under **⚙ → Diagnostics**. The window's Esc / Enter / Backspace keep working even then.

---

## Audit checks

| Check | Severity | Notes |
|---|---|---|
| Broken reference | Risk | `#REF!` or an unresolvable name/sheet |
| Formula error | Risk | `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NUM!`, `#SPILL!`, `#CALC!`. **Not** `#N/A` (a normal "lookup found nothing"). |
| Hardcode in calc area | Risk | a typed number surrounded by formulas — likely a plug |
| Inconsistent formula | Risk | breaks its row/column pattern |
| Hardcode inside formula | Warning | e.g. `=A1*1.05`, `=A1*(1+5%)`, `=A1+1000`. **Not** structural function arguments (`ROUND(x,2)`, `VLOOKUP(…,3,0)`, `INDEX(…,2,3)`, `EOMONTH(x,3)` …) and **not** unit factors (`×12`, `/1000`, `/365`, `×100` …). |
| Empty reference | Warning | a formula points at an empty cell — unless it tests for that (`IF(A2="",…)`, `ISBLANK(A2)`, `A2=0`) or is wrapped in `IFERROR` |
| Division-by-zero risk | Warning | every `…/B7` in the formula is checked; text in quotes is ignored; `IFERROR` or an `IF` that tests the divisor counts as guarded |
| Number stored as text | Warning | only when a real number sits next to it (a lone "2024" label is fine) |
| Volatile function / no dependents | Info | `OFFSET`, `INDIRECT`, `NOW` …; formulas nothing else uses |
| Circular chain / hidden-sheet input / unused name | Risk / Risk / Info | whole-workbook audit |

**Data plugins are skipped** for value and hardcode checks — FactSet `FDS*`, Capital IQ `CIQ*`, Bloomberg `BDP/BDH/BDS/BEQS/BSRCH/BQL`, `RTD`, Refinitiv/LSEG `TR/RHistory/RSearch/RtGet`, also when written as `_xll.BDP(…)`. Their blanks, errors and literal arguments aren't model bugs.

**Names counted as used:** a name referenced by a cell formula, called as a LAMBDA (`=MyFn(A1)`), or used inside another name's definition. Usage in charts, data validation, conditional formatting and VBA is not visible to the add-in — that is why deletion asks first and keeps a backup.

---

## Keyboard

**Floating windows** (Arixcel default + one bonus key):

| Key | Action |
|---|---|
| Ctrl+Shift+C / D | open Precedents / Dependents window |
| ↑ / ↓ | move between rows |
| → / ← | expand / collapse (← on a collapsed row jumps to its parent) |
| Enter / Esc | close the window, stay on the current cell |
| Ctrl+Backspace | close and return to the explored cell |
| Backspace | same as Ctrl+Backspace |

Window hint: `↑↓ move · →/← open/close · Enter/Esc close · Ctrl+Backspace back` / `Drag to a corner — it stays · Ctrl+Shift+C/D retarget`.

**Explore panel:**

| Key | Action |
|---|---|
| Ctrl+Shift+E | open the pane / re-anchor on the selected cell |
| ↑ / ↓ | move between tree rows (↑ from the first row reaches the "← Back" banner); with Follow on, Excel follows |
| → / ← | unfold / collapse a component; on the SUMIFS match-walk row: next / previous match |
| Enter | jump Excel to the focused row's cell |
| Backspace / Ctrl+Backspace | return to the explored cell (not while editing a formula) |
| Esc | hide the pane |

**Map list:** ↑ / ↓ move through findings (Excel follows when Follow is on), Enter jumps.

> Keyboard works once focus is in the pane — click a row first. Focus still in the Excel grid can't reach the pane (an Office limitation). The add-in's three global shortcuts don't collide with the companion formatting macros (Ctrl+Shift+Q/T/R/B/A/Z/J/M).

---

## Install — no administrator rights needed

The add-in is served from **https://alimzhansagymbekov.github.io/formula-navigator/** (a static GitHub Pages site). Nothing runs in the background on your PC.

1. Close Excel.
2. From the `installer` folder, double-click **`Install-FormulaNavigator.cmd`** (keep `manifest.xml` next to it). It copies the manifest to `%LOCALAPPDATA%\FormulaNavigator` and registers it under your own user registry key (`HKCU`) — no admin password.
3. Open Excel → **Home → Formula Navigator**. If the button isn't there: **Home → Add-ins → My Add-ins → Developer Add-ins → Formula Navigator**.

Uninstall: close Excel, run **`Uninstall-FormulaNavigator.cmd`**.

If your company blocks this kind of per-user install by policy, ask IT to deploy `manifest.xml` centrally (Microsoft 365 admin center → Integrated apps). The older shared-folder method (Trusted Add-in Catalog) still works but needs Windows folder sharing, which asks for an admin password — don't use both methods at once.

## Is this safe?

Everything is readable source code — no compiled binaries, no telemetry, no account. The add-in runs inside Excel's sandbox, sees only the workbook you have open, and makes no network calls of its own. It never changes your formatting on its own; it writes to a workbook only when you explicitly act: Apply in the formula editor, Delete/Restore a name, Unhide all sheets, or Export to sheet.

---

## Build from source

```bash
cd formula-explorer
npm install
npm run build       # outputs to dist/
npm run typecheck   # tsc --noEmit
npm test            # 324 tests: engine, audit, keyboard rules, exact values, spans, DOM + React (jsdom)
npm run test:browser  # optional: the real pane + window in Chromium over a fake workbook (needs Playwright)
```

`dist/` is what GitHub Pages serves. To update the hosted add-in, upload the changed `dist/` files to the repository root and restart Excel; users get the update on their next Excel start.

## Project layout

- `src/core/` — engine: parser, dependency graph, evaluator, scanner, audit, findings model (no Excel, unit-tested).
- `src/shared/key-logic.ts` — every keyboard/click decision, unit-tested.
- `src/shared/` (V15) — exact-value formatting, formula highlight spans, argument captions, match walk, the crash safety net (`ErrorBoundary`) — shared by the pane and the window, unit-tested.
- `test/browser/` — Playwright harness that runs the production UI code against a fake workbook.
- `src/taskpane/` — the pane UI (React): Explore / Trace / Map / Find, plus `app-controller.ts` (state + Excel orchestration).
- `src/dialog/` — the floating precedent/dependent windows.
- `src/core/excel/excel-adapter.ts` — the only layer that talks to Excel.
- `../installer/` — no-admin install/uninstall scripts.

## Limitations (Office.js platform)

- Cross-**workbook** tracing isn't possible — Office.js exposes only the open workbook.
- One floating window at a time; it can't be positioned programmatically or made transparent.
- Name usage in charts, data validation, conditional formatting and VBA isn't visible.
- The pane receives keyboard input only when focus is inside it.

## Version history (short)

- **V16** — Trace / window: one row per range (folded, filled count, → lists cells, Enter selects the range); verified "found by lookup" rows for the traced formula (incl. INDIRECT/OFFSET); Explore: defined names and table refs navigable, names no longer squeezed by their definition chip, each explored cell starts a fresh tree (open/closed no longer leaks between formulas).
- **V15.2** — the explored cell no longer jumps away after the pane's own navigation: a late or doubled Excel selection event that reports exactly the cell the pane just moved to (within 3 s) is recognised as that jump's echo, not a user click (before, the "← Back" banner vanished and the tree switched cells — most often across sheets, with Follow on, or when Excel was busy); an explicit jump cancels a queued Follow jump; a busy Excel no longer blanks the selection; selection events are settled into one read; version shown in ⚙ → Diagnostics.
- **V15.1** — empty single-cell precedents listed as `(empty)`; numbers re-read after every edit and recalculation (they used to go stale downstream of an edit); the open floating window updates its numbers live; Rescan no longer duplicates change listeners; sheets added later are listened to immediately.
- **V15** — crash safety net (per tab + window, errors in Diagnostics); Follow restored in Explore/Trace (default on, remembered); formula highlight (pane line + window footer with the linking reference); right-aligned value column with → lookup / ← source chips; Values: as shown / exact everywhere; inline argument captions for 96 functions; IF conditions no longer dimmed; two-line window hint; SUMIFS match walk (to the summed values); browser test harness.
- **V14** — no grid painting at all (removed the "you are here" fill); safe name deletion (usage in other names and LAMBDA calls, confirmation, Restore); stable review statuses (auto-migrated); export matches the pane; far less audit noise (structural arguments, unit factors, guarded formulas, quoted text, plugins incl. `_xll.`); `#SPILL!`/`#CALC!`; one jump per click; Diagnostics/Validate shown in the pane; keyboard rules unit-tested; 0 npm vulnerabilities; no-admin installer.
- **V13** — `#N/A` not flagged; context-aware number-as-text.
- **V12** — empty reference, formula error, number-as-text, division-by-zero checks.
- **V8–V11** — keyboard: Back banner, Backspace/Ctrl+Backspace, autoscroll, focus after Back; Arixcel key map.
- **V4–V7** — review statuses, export, multi-cell trace, unified findings list, delete unused names.
