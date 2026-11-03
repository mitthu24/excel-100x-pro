# ⚡ Excel 100x Pro

Two editions of the same idea: one ribbon tab that does in a click what normally takes minutes.

| | **Desktop edition** `Excel-100x-Pro.xlam` | **Web edition** (Office.js) |
|---|---|---|
| Tools | **202**, all in one **100x Pro Tools** tab | 82, in a "100x Pro" tab + task pane |
| Works on | Excel for Windows (2016+) | Windows, Mac, Excel on the web |
| Install | copy the file, tick one box — see [INSTALL-XLAM.md](INSTALL-XLAM.md) | hosted, add via Microsoft 365 admin |
| Needs | nothing, works offline | the hosted URL below |

Live web edition: <https://excel-100x-pro.vercel.app/> · Architecture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)

Select **one cell** and most tools work on the whole sheet; select a range for precise control.
**Ctrl+Z** undoes any tool.

---

## Desktop edition — the 100x Pro Tools tab

**Clean** — Trim Spaces · Magic Clean · Case ▾ · Fix ▾ (text→number, junk characters, fill blanks, unmerge & fill) · Remove ▾ (duplicates, blank rows, blank columns, hyperlinks)

**Transform** — Split Column · Merge Columns · Transpose · Unpivot Columns · Group and Sum · Sort ▾ (A–Z, Z–A, shuffle) · Insert ▾ (serial numbers, GUIDs, date, timestamp, drop-down) · Amount in Words ▾ (₹ Lakh/Crore, $ Million)

**Text Tools** — Extract ▾ (emails, numbers, phone numbers, remove digits) · Add ▾ (prefix, suffix, pad zeros, reverse) · Find & Replace in selection · Round Values · Text to Date

**Dates & Calc** — Date Column ▾ (weekday, month, quarter, financial year, age) · Add Column ▾ (% of total, running total, rank, % change)

**Format** — Beautify ▾ (5 themes) · Zebra Rows · Auto-fit · Freeze Header · Numbers ▾ (₹, Indian 12,34,567 grouping, $, percent, date, text) · More ▾ (table, gridlines, borders, wrap, clear)

**Formulas** — Formulas→Values · Wrap IFERROR · Highlight Formulas · References ▾ (absolute, relative, names from headers) · Quick Formulas ▾ (XLOOKUP, VLOOKUP, INDEX+MATCH, SUMIFS, COUNTIFS, IFS, TEXTJOIN, NETWORKDAYS, PMT)

**Sheets** — Index Sheet · Sort Sheets · Compile Sheets · Compare Two Sheets · Lock Formulas Only · Sheets from List · Split by Column · Copy Sheet N Times · Visibility ▾ (unhide, rainbow tabs, protect/unprotect all) · **Folder ▾ — Compile Workbooks from Folder · Folder Files to Sheets · List Files in Folder · Save Each Sheet as Workbook**

**Print & Export** — Sheet to PDF · Payslips from Payroll (one PDF per employee) · Email Sheet as PDF · Set Print Area · A4 Page Setup · Fit to 1 Page · Repeat Header Rows · Header & Footer · PDF ▾ (selection, workbook, each sheet)

**Highlight** — duplicates, blanks, errors, outliers, top N, bottom N, above average · Heat Map · Data Bars · Icon Sets · Compare Columns

**Analyze** — Statistics Sheet · Frequency Table · Pareto/ABC · Correlation Matrix · Forecast · Quick Pivot · Live Insights

**Charts** — column, line, bar, pie, area, scatter · KPI Dashboard

**Select** — blanks, formulas, constants, errors, visible cells

**Documents** — GST Invoice · Quotation · Purchase Order · Salary Slip · Vouchers ▾ (payment receipt, delivery challan) · Letters ▾ (offer letter, experience letter, certificate)

**Templates (28)** — Everyday (expenses, budget, calendar, to-do, timesheet, minutes, attendance, project tracker) · Sales & Marketing (CRM pipeline, sales tracker, commission, customers, content calendar, campaign ROI) · HR & Admin (directory, leave, payroll, recruitment) · Finance (invoice register, cash flow, P&L, receivables ageing, loan EMI schedule) · Operations (inventory, purchases, vendors, production, BOM, shipments, maintenance) · Industry (gradebook, patient register, property listing, retail sales, menu costing, site log)

**Pro Tools** — Workbook Audit · Data Entry Form · Backup Copy · Go to Sheet · Zoom All Sheets · Privacy ▾ (hide/show values, mask text, remove personal info) · Workbook ▾ (break links, all formulas to values, delete empty sheets) · Web edition · About

---

## Install

**Desktop (.xlam):** see [INSTALL-XLAM.md](INSTALL-XLAM.md) — copy the file, then File → Options → Add-ins → Excel Add-ins → Go → Browse.

**Web edition, whole company:** Microsoft 365 Admin Center → Settings → Integrated apps → Upload custom apps → Office Add-in → upload `dist/manifest.xml`, assign to users.

**Web edition, one user:** Excel on the web → Insert → Add-ins → Manage My Add-ins → Upload My Add-in → `dist/manifest.xml`.

**Web edition, local trial:** `START-LOCAL.cmd` (needs Node.js; rebuild with `node build/build.js` first so the manifest points at localhost).

## Build from source

```bash
node build/build.js --icons --url https://your-host     # web edition -> dist/ + zip
powershell -ExecutionPolicy Bypass -File xlam/make-xlam.ps1   # desktop edition -> Excel-100x-Pro.xlam
```

The `.xlam` build drives your installed Excel, so it needs **File → Options → Trust Center →
Trust Center Settings → Macro Settings → Trust access to the VBA project object model** ticked.
Turn it back off afterwards. Every tool is defined once in `xlam/xlam-tools.js`; the ribbon, the
icons and the VBA dispatcher are generated from it.

## Notes and limits

- The two editions are independent: the desktop one has more tools, the web one runs anywhere.
- Page-setup and PDF tools need a default printer installed (even "Microsoft Print to PDF").
- Beautify clears existing conditional formats; Compare and Outliers colour cells. Ctrl+Z undoes them.
- "All formulas to values" and "Delete empty sheets" ask first and cannot be undone with Ctrl+Z.
- Everything runs locally. No data leaves the workbook, no telemetry, no sign-in.
