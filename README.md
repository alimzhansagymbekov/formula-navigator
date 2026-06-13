[README-FormulaNavigator.md](https://github.com/user-attachments/files/28914811/README-FormulaNavigator.md)
# formula-navigator
# Formula Navigator — Excel Add-in

An Arixcel-class formula explorer for Excel. Every formula component shows its **calculated value**, the **branch actually taken** by IF/CHOOSE is marked, **lookups resolve to their target cells**, precedents and dependents trace across sheets, and SUMIFS-style formulas explain **exactly which records matched**. Built on Office.js (TypeScript + React); runs inside Excel's sandbox.

> **Status:** version 11. Desktop Excel on Microsoft 365 (Windows/Mac). Some features need ExcelApi 1.9+ (any current 365 build); the add-in tells you in a banner if your host is missing something.

---

## What it does

### Explore — the selected formula as a live tree
- Every component row shows its **computed value** (`= 282,048`); cell references show their **displayed text** with the cell's own number format (`12.5%`, `1,234`, dates).
- The **branch actually taken** by IF / IFS / CHOOSE / SWITCH / IFERROR is marked in green; untaken branches are dimmed.
- **INDEX, VLOOKUP, HLOOKUP, XLOOKUP, OFFSET, INDIRECT, MATCH are evaluated** — a green **→ Sheet!B14** chip jumps to the exact cell the lookup landed on.
- For **SUMIF / SUMIFS / COUNTIF(S) / AVERAGEIF(S) / FILTER** and **SUMPRODUCT logical filters** like `=SUMPRODUCT(--(A2:A100="West"),B2:B100)`, the tree shows **which records matched** ("4 of 48 records matched"), explains **each criterion with its own hit count**, and lists the first matched cells as clickable jump links. Full-column ranges (`A:A`) are clamped to the used range automatically.
- The **▸ on any reference unfolds that cell's own formula in place** (cycle-safe, multi-level), so a whole calculation chain reads as one tree.
- **Multi-cell trace:** select a range and the tree builds precedents/dependents for every cell at once.
- A **yellow "← Back" banner** appears when you've navigated away; click it, or press **Backspace / Ctrl+Backspace**, to return to the explored cell. Arrow keys navigate the tree; **Follow** (toggle) makes Excel's selection move with you.

### Trace — precedents / dependents across sheets
- Multi-level tree with values, range-membership chips, and cycle flags; auto-refreshes when you edit.
- An **impact line** sums the whole-model blast radius ("feeds 1,284 formulas across 9 sheets").
- A **Tree | Graph toggle** switches to an interactive dependency graph (levels flow left-to-right, nodes carry sheet color and live values, branches collapse, click a node to jump).

### Map — sheet classification + workbook audit
- Classifies the active sheet (input / formula / output / cross-sheet / label) and flags **inconsistent formulas, hardcoded plugs, broken references**.
- **Audit workbook** runs whole-model checks in one pass: **circular reference chains, formulas reading hidden sheets, unused named ranges**.
- All findings appear in **one keyboard-navigable list** — arrow ↑/↓ through them and Excel follows; filter by status or category; each finding carries a **review status** (Open → Checked → Question → Error) you cycle by clicking its pill.
- **Unused named ranges** can be deleted from the workbook with a per-row Delete button. (Excel's internal `_xlfn.*` names are hidden — they aren't real unused names.)
- **Export to sheet** writes every finding — with its review status (color-coded) and the cell's formula — to a "Formula Navigator Audit" worksheet.

### Find — pins and search
- Pinned cells (★, saved inside the workbook), plus search by reference (`Model!B7`), function (`XLOOKUP`), or **named range**, with dependency counts.

### Floating windows (Ctrl+Shift+C / D)
Compact, keyboard-first precedent/dependent windows that open over the grid. They reuse a single window (drag it to a corner and it stays there for the session).

---

## Keyboard

**Floating windows** (Arixcel-aligned):

| Key | Action |
|---|---|
| Ctrl+Shift+C / D | open Precedents / Dependents window |
| ↑ / ↓ | move between rows |
| → / ← | expand / collapse (← on a collapsed row jumps to its parent) |
| Enter / Esc | close the window, stay on the current cell |
| Ctrl+Backspace | close and return to the explored cell |
| Backspace | same as Ctrl+Backspace |

**Explore panel:**

| Key | Action |
|---|---|
| Ctrl+Shift+E | open the pane / re-anchor on the selected cell |
| ↑ / ↓ | move between tree rows (↑ from the first row reaches the "← Back" banner) |
| → / ← | unfold / collapse a component |
| Enter | jump Excel to the focused row's cell |
| Backspace / Ctrl+Backspace | return to the explored cell |
| Esc | hide the pane |

> Keyboard works once focus is in the pane — click a tree row first. Focus that's still in the Excel grid can't reach the pane (an Office limitation).

---

## Install (hosted on GitHub Pages — nothing runs on your PC)

The add-in is served from **https://alimzhansagymbekov.github.io/formula-navigator/** — a static page on GitHub Pages. No local server, no certificates, no background process; Excel loads it like a store add-in.

One-time setup per machine (~3 minutes; see `INSTALL-INSTRUCTIONS.md` for click-by-click steps):

1. Put `manifest.xml` into a **shared folder** (Windows network sharing enabled).
2. Add that folder as a **Trusted Add-in Catalog** (File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs), tick *Show in Menu*, restart Excel.
3. Home → Add-ins → **More Add-ins** → **SHARED FOLDER** tab → Formula Navigator → **Add**.

To remove it: delete `manifest.xml` from the shared folder and clear Office's add-in cache (`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`).

## Is this safe?

Everything is plain, readable source code — no compiled binaries, no telemetry; your workbooks never leave your machine. GitHub Pages only delivers the add-in's static files; the add-in runs inside Excel's sandbox with access only to the workbook you have open, and makes no network calls of its own.

---

## Build from source

```bash
cd formula-explorer
npm install
npm run build       # outputs to dist/
npm run typecheck   # tsc --noEmit
npm test            # node test harness (88 tests)
```

The built `dist/` is what GitHub Pages serves. To update the hosted add-in, upload the changed files in `dist/` to the repo (root) and restart Excel.

## Project layout

- `src/core/` — engine: parser, dependency graph, evaluator, scanner, audit, findings model (no Excel, unit-tested).
- `src/taskpane/` — the pane UI (React): Explore / Trace / Map / Find, plus `app-controller.ts` (state + Excel orchestration).
- `src/dialog/` — the floating precedent/dependent windows.
- `src/core/excel/` — the Office.js adapter (the only layer that talks to Excel).

## Limitations (Office.js platform)

- Cross-**workbook** tracing isn't possible — Office.js exposes only the open workbook.
- Floating windows can't be positioned programmatically (Office centers them) or made transparent on all builds; drag one to a corner and it persists for the session.
- The pane receives keyboard input only when focus is inside it (click a row first).
