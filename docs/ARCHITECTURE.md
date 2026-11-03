# Excel 100x Pro — Architecture

## 1. Goal
One add-in that makes employees faster in Excel: clean data, reshape it, format it, audit it, and spin up business templates in one or two clicks. Same package serves enterprises (central deploy), companies (shared catalog) and small businesses (sideload).

## 2. Technology choice
| Option | Verdict |
|---|---|
| VBA `.xlam` | Needs Excel to compile `vbaProject.bin`; macro-security prompts; Windows only. Rejected. |
| VSTO / COM | Needs .NET + admin install per PC. Rejected. |
| **Office.js web add-in** | Pure HTML/JS, no install of binaries, works on Windows/Mac/Web, central deployment via Microsoft 365 admin. **Chosen.** |

## 3. Layers
```
┌────────────────────────── Excel ribbon ──────────────────────────┐
│ Tab "Pro 100x"  →  8 groups  →  buttons / dropdown menus         │  (generated manifest.xml)
└───────────────┬───────────────────────────────┬──────────────────┘
        ExecuteFunction                    ShowTaskpane
                │                               │
        commands.html                    taskpane.html
        (headless, runs tool)            (Command Palette · Insights ·
                │                         Split/Merge · Formulas · Settings)
                └──────────────┬────────────────┘
                         src/tools.js   ← single registry: id, label, group, icon, run()
                         src/lib.js     ← helpers (selection, toast, number→words…)
```
**Single source of truth:** `src/tools.js`. The ribbon manifest, icons, command handlers and the palette search are all *generated from it*. Add a tool once → it appears on the ribbon, in the palette, and works.

## 4. Ribbon map
| Group | Purpose | Tools |
|---|---|---|
| Clean | Fix messy data | Trim, Remove junk chars, Case ▾, Text→Number, Fill blanks, Delete blank rows, Remove duplicates, Magic Clean |
| Transform | Reshape | Split / Merge columns (pane), Transpose, Sort ▾, Serial numbers, GUIDs, Amount in words (₹ Lakh/Crore + intl) |
| Format | Make it beautiful | Beautify ▾ (4 themes), Zebra, Autofit, Freeze header, Number formats ▾, Clear format, Table, Gridlines |
| Formulas | Formula power | Formulas→values, IFERROR wrap, Highlight formulas, Formula library (pane) |
| Sheets | Workbook control | Index sheet, Sort sheets, Unhide all, Protect/Unprotect all, Rainbow tabs, Sheets from list, Split by column, Merge sheets |
| Analyze | Insight | Highlight dupes/blanks/errors, Outliers, Heatmap, Data bars, Compare columns, Quick pivot, Quick chart ▾, Select special ▾, Stats |
| Business | Ready templates | Invoice, Expenses, Budget, Inventory, Attendance, Project tracker, KPI dashboard |
| Pro | Control centre | Command palette, Insights pane, Hide/Show values, Workbook audit, Settings |

## 5. Task pane
Modern single-page app (no framework, ~1 CSS file): tabs **Search** (fuzzy command palette over every tool), **Insights** (live stats of the selection), **Tools** (split/merge forms), **Formulas** (searchable library, one-click insert), **Settings** (accent colour, dark mode — saved in localStorage).

## 6. Build & deploy
`node build/build.js [--url https://host]` → `dist/` (static site + `manifest.xml` with that base URL + PNG icons) → `Excel-100x-Pro.zip`.
- **Small business / single PC:** `START-LOCAL.cmd` (https localhost + trusted dev cert + shared-folder sideload).
- **Company:** host `dist/` on any HTTPS static host, share manifest via network catalog.
- **Enterprise:** Microsoft 365 Admin Center → Integrated apps → upload `manifest.xml` → assign to users/groups.

## 7. Quality rules
Every tool: works on the current selection (or used range if single cell), batches one `context.sync()` where possible, shows a toast with the result, and never needs a network. No telemetry, no data leaves the workbook.
