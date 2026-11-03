# ⚡ Excel 100x Pro

82 one-click tools in a custom **100x Pro** ribbon tab + a modern task pane (command search, live insights, split/merge, formula library, themes). Built as an Office.js add-in: works on Excel for Windows, Mac and the web. Everything runs locally - no data leaves the workbook.

Architecture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)

## Ribbon

| Group | Tools |
|---|---|
| **Clean** | Trim Spaces · Magic Clean · Case ▾ · Fix ▾ (text→number, junk chars, fill blanks) · Remove ▾ (duplicates, blank rows) |
| **Transform** | Split Column · Merge Columns · Transpose · Sort ▾ · Insert ▾ (serials, GUIDs, date, timestamp) · Amount in Words ▾ (₹ Lakh/Crore, $ Million) |
| **Format** | Beautify ▾ (5 themes) · Zebra Rows · Auto-fit · Freeze Header · Numbers ▾ (₹ $ % date text) · More ▾ (table, gridlines, clear) |
| **Formulas** | Formulas→Values · Wrap IFERROR · Highlight Formulas · Formula Library (42 formulas) |
| **Sheets** | Index Sheet · Sort Sheets · Visibility ▾ (unhide, rainbow tabs, protect/unprotect all) · Sheets from List · Split by Column · Merge Sheets |
| **Analyze** | Highlight ▾ (duplicates, blanks, errors, outliers) · Heat Map · Data Bars · Compare Columns · Quick Pivot · Chart ▾ · Select ▾ |
| **Business** | Invoice (GST) · Templates ▾ (Expenses, Budget, Inventory, Attendance, Project Tracker) · KPI Dashboard |
| **Pro Tools** | Command Search · Live Insights · Privacy ▾ (hide/show values) · Workbook Audit · Settings |

Tip: select **one cell** and most tools act on the whole sheet; select a range for precise control. Ctrl+Z undoes any tool.

## Install - pick one

### A. One PC / small business (5 minutes)
1. Install [Node.js](https://nodejs.org) (LTS).
2. Unzip, double-click **`START-LOCAL.cmd`**, confirm the certificate prompt.
3. Excel → *Home → Add-ins → Developer Add-ins → Excel 100x Pro*. Keep the black window open while working.

### B. Whole company / enterprise (recommended)
Host the add-in once on any HTTPS static host (Azure Static Web Apps, Cloudflare Pages, GitHub Pages, your IIS/nginx):
```bash
node build/build.js --url https://addins.yourcompany.com/excel100x
```
Upload the contents of `dist/` to that URL, then:
- **Microsoft 365 Admin Center → Settings → Integrated apps → Upload custom apps → Office Add-in → upload `dist/manifest.xml`** → choose users/groups. The tab appears for them within hours. (Best for enterprises: central control, auto-updates - just re-upload files.)
- Or share `manifest.xml` via a trusted network-share catalog (File → Options → Trust Center → Trusted Add-in Catalogs).

### C. Excel on the web
*Insert → Add-ins → Manage My Add-ins → Upload My Add-in* → pick `dist/manifest.xml` (needs a hosted URL as in B).

## Customise
All tools live in `src/tools.js` (one entry = ribbon button + search entry). Add a tool, then:
```bash
node build/build.js --icons --url <your url>   # regenerates icons (needs Python + Pillow), manifest and zip
```

## Notes
- Requires ExcelApi 1.9+ (Excel 2019 / Microsoft 365).
- Formula-library entries marked (365) need Microsoft 365.
- Some tools overwrite formatting (Beautify clears existing conditional formats; Compare/Outliers colour cells) - use Ctrl+Z.
