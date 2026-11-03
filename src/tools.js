/* Excel 100x Pro - single source of truth for every tool.
   tool: {id, g:group, menu?, label, tip, e:emoji, k:search keywords, pane?, run?(ctx,args)->message} */
const GROUPS = [
  { id: 'clean', label: 'Clean', e: '🧹' },
  { id: 'transform', label: 'Transform', e: '🔄' },
  { id: 'format', label: 'Format', e: '🎨' },
  { id: 'formulas', label: 'Formulas', e: '🧮' },
  { id: 'sheets', label: 'Sheets', e: '📑' },
  { id: 'analyze', label: 'Analyze', e: '📊' },
  { id: 'business', label: 'Business', e: '💼' },
  { id: 'pro', label: 'Pro Tools', e: '🚀' },
];
const MENUS = { // menu key -> emoji
  Case: '🔠', Fix: '🩹', Remove: '🗑️', Sort: '↕️', Insert: '➕', Words: '🗣️', Beautify: '🪄', Numbers: '💱',
  More: '⚙️', Visibility: '👁️', Highlight: '🖍️', Chart: '📈', Select: '🎯', Templates: '📋', Privacy: '🔒',
};
const TOOLS = [];
const add = o => TOOLS.push(o);

const THEMES = {
  ocean: ['#0F4C81', '#EAF2FB', 'Ocean Blue'], emerald: ['#0B6E4F', '#E6F4EE', 'Emerald'],
  sunset: ['#C2410C', '#FFF1E6', 'Sunset'], slate: ['#1F2937', '#F1F3F5', 'Slate'], violet: ['#5B21B6', '#F3EEFF', 'Violet'],
};
const NUMFMT = {
  inr: ['Rupees (₹)', '[$₹-4009]#,##0.00', '₹'], usd: ['Dollars ($)', '$#,##0.00', '💵'],
  num: ['Number 1,234', '#,##0', '🔢'], pct: ['Percent', '0.0%', '％'], date: ['Date', 'dd-mmm-yyyy', '📅'], txt: ['Text', '@', '🔤'],
};
const NUMFMT_EMOJI = { inr: '💰', usd: '💵', num: '🔢', pct: '📉', date: '📅', txt: '🔤' };

/* ---------- helpers local to tools ---------- */
const strFn = f => async ctx => {
  const { r } = await P.target(ctx);
  const n = await P.mapCells(ctx, r, v => (typeof v === 'string' ? f(v) : v));
  return n + ' cell(s) updated';
};
const TRIM = s => s.replace(/[ ​]/g, ' ').replace(/\s+/g, ' ').trim();
const JUNK = s => s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F​-‏﻿]/g, '');
const TONUM = s => {
  const t = s.replace(/[,\s₹$€£]/g, '');
  if (!/^[-+]?(\d+\.?\d*|\.\d+)%?$/.test(t)) return s;
  return t.endsWith('%') ? parseFloat(t) / 100 : parseFloat(t);
};
const colKeyIdx = async ctx => { // key column for sort / split: active cell column inside the target
  const t = await P.target(ctx);
  return { ...t, key: t.whole ? t.sel.columnIndex - t.r.columnIndex : 0 };
};

/* ================= CLEAN ================= */
add({ id: 'trim', g: 'clean', label: 'Trim Spaces', e: '✂️', k: 'spaces whitespace nbsp', tip: 'Remove leading, trailing and double spaces (incl. non-breaking) in the selection, or the whole sheet if one cell is selected.', run: strFn(TRIM) });
add({ id: 'magic', g: 'clean', label: 'Magic Clean', e: '✨', k: 'all in one fix', tip: 'One click: trim spaces, strip hidden characters and convert numbers stored as text.', run: strFn(s => TONUM(JUNK(TRIM(s)))) });
[['upper', 'UPPERCASE', '🔠', s => s.toUpperCase()], ['lower', 'lowercase', '🔡', s => s.toLowerCase()],
['proper', 'Proper Case', '🔤', s => s.toLowerCase().replace(/(^|[\s\-(])(\p{L})/gu, (_, a, b) => a + b.toUpperCase())],
['sentence', 'Sentence case', '📝', s => s.toLowerCase().replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (_, a, b) => a + b.toUpperCase())],
].forEach(([id, label, e, f]) => add({ id, g: 'clean', menu: 'Case', label, e, k: 'text case', tip: `Convert text to ${label}.`, run: strFn(f) }));
add({ id: 'textnum', g: 'clean', menu: 'Fix', label: 'Text → Number', e: '🔢', k: 'convert numbers stored as text', tip: 'Convert numbers stored as text (also 1,200 / ₹500 / 12%) into real numbers.', run: strFn(TONUM) });
add({ id: 'junk', g: 'clean', menu: 'Fix', label: 'Remove Junk Chars', e: '🧽', k: 'non printable hidden characters clean', tip: 'Strip non-printable and zero-width characters that break lookups.', run: strFn(JUNK) });
add({
  id: 'fillblank', g: 'clean', menu: 'Fix', label: 'Fill Blanks Down', e: '⬇️', k: 'empty cells copy above', tip: 'Fill empty cells with the value above them (great for grouped reports).',
  run: async ctx => {
    const { r } = await P.target(ctx);
    const last = {};
    const n = await P.mapCells(ctx, r, (v, i, j) => {
      if (v === '' && last[j] !== undefined) return last[j];
      if (v !== '') last[j] = v;
      return v;
    });
    return n + ' blank(s) filled';
  },
});
add({
  id: 'dedupe', g: 'clean', menu: 'Remove', label: 'Remove Duplicates', e: '👯', k: 'duplicate unique', tip: 'Delete duplicate rows (first row is treated as header).',
  run: async ctx => {
    const { r } = await P.target(ctx);
    const res = r.removeDuplicates(Array.from({ length: r.columnCount }, (_, i) => i), true);
    await ctx.sync();
    const o = res.value || res;
    return `${o.removed} duplicate row(s) removed, ${o.uniqueRemaining} kept`;
  },
});
add({
  id: 'delblank', g: 'clean', menu: 'Remove', label: 'Delete Blank Rows', e: '🚮', k: 'empty rows', tip: 'Delete rows that are completely empty inside the selection / used range.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    r.load('values');
    await ctx.sync();
    let n = 0;
    for (let i = r.values.length - 1; i >= 0; i--) {
      if (r.values[i].every(v => v === '')) { r.getRow(i).getEntireRow().delete('Up'); n++; }
    }
    await ctx.sync();
    return n + ' blank row(s) deleted';
  },
});

/* ================= TRANSFORM ================= */
add({
  id: 'split', g: 'transform', label: 'Split Column', e: '✂️', pane: 'tools', k: 'text to columns delimiter comma', tip: 'Split text in the selected column into several columns by a delimiter.',
  run: async (ctx, { delim }) => {
    const { sel } = await P.target(ctx, false);
    const src = sel.getColumn(0);
    src.load('values');
    await ctx.sync();
    const rows = src.values.map(x => String(x[0]).split(delim));
    const max = Math.max(...rows.map(x => x.length));
    if (max < 2) return 'Delimiter not found';
    const ws = sel.worksheet;
    ws.getRangeByIndexes(0, sel.columnIndex + 1, 1, max - 1).getEntireColumn().insert('Right');
    sel.getColumn(0).getOffsetRange(0, 1).getResizedRange(0, max - 2).values = rows.map(x => Array.from({ length: max - 1 }, (_, i) => (x[i + 1] ?? '').trim()));
    src.values = rows.map(x => [x[0].trim()]);
    await ctx.sync();
    return `Split into ${max} columns`;
  },
});
add({
  id: 'merge', g: 'transform', label: 'Merge Columns', e: '🔗', pane: 'tools', k: 'join combine concatenate', tip: 'Join the selected columns into one new column with a separator.',
  run: async (ctx, { delim }) => {
    const { sel } = await P.target(ctx, false);
    if (sel.columnCount < 2) return 'Select 2 or more columns';
    sel.load('values');
    await ctx.sync();
    sel.worksheet.getRangeByIndexes(0, sel.columnIndex + sel.columnCount, 1, 1).getEntireColumn().insert('Right');
    sel.getLastColumn().getOffsetRange(0, 1).values = sel.values.map(r => [r.filter(v => v !== '').join(delim)]);
    await ctx.sync();
    return 'Merged into a new column';
  },
});
add({
  id: 'transpose', g: 'transform', label: 'Transpose', e: '🔀', k: 'rows to columns flip rotate', tip: 'Flip rows and columns of the selection into a new sheet.',
  run: async ctx => {
    const { r } = await P.target(ctx, false);
    const ws = await P.newSheet(ctx, 'Transposed');
    ws.getRange('A1').copyFrom(r, 'All', false, true);
    ws.activate();
    await ctx.sync();
    return 'Transposed into a new sheet';
  },
});
[['sortasc', 'Sort A → Z', '🔼', true], ['sortdesc', 'Sort Z → A', '🔽', false]].forEach(([id, label, e, asc]) =>
  add({
    id, g: 'transform', menu: 'Sort', label, e, k: 'order', tip: 'Sort by the column of the active cell (first row is the header).',
    run: async ctx => {
      const { r, key, whole } = await colKeyIdx(ctx);
      r.sort.apply([{ key, ascending: asc }], false, whole);
      await ctx.sync();
      return 'Sorted';
    },
  }));
add({
  id: 'serial', g: 'transform', menu: 'Insert', label: 'Serial Numbers', e: '🔢', k: 'sr no sequence 1 2 3', tip: 'Fill 1, 2, 3 ... down the first selected column.',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    sel.getColumn(0).values = Array.from({ length: sel.rowCount }, (_, i) => [i + 1]);
    await ctx.sync();
    return 'Serial numbers added';
  },
});
add({
  id: 'guid', g: 'transform', menu: 'Insert', label: 'Unique IDs (GUID)', e: '🆔', k: 'uuid random', tip: 'Fill the selection with unique GUIDs.',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    sel.values = Array.from({ length: sel.rowCount }, () => Array.from({ length: sel.columnCount }, () => crypto.randomUUID()));
    await ctx.sync();
    return 'IDs generated';
  },
});
add({
  id: 'today', g: 'transform', menu: 'Insert', label: 'Today\'s Date', e: '📅', k: 'date stamp', tip: 'Insert today\'s date (static value) in the active cell.',
  run: async ctx => {
    const c = ctx.workbook.getSelectedRange().getCell(0, 0);
    c.values = [[Math.floor(P.serial())]];
    c.numberFormat = [['dd-mmm-yyyy']];
    await ctx.sync();
    return 'Date inserted';
  },
});
add({
  id: 'stamp', g: 'transform', menu: 'Insert', label: 'Timestamp', e: '⏱️', k: 'date time now', tip: 'Insert the current date & time (static value).',
  run: async ctx => {
    const c = ctx.workbook.getSelectedRange().getCell(0, 0);
    c.values = [[P.serial()]];
    c.numberFormat = [['dd-mmm-yyyy hh:mm']];
    await ctx.sync();
    return 'Timestamp inserted';
  },
});
[['wordsinr', 'Amount in Words (₹)', '🪙', 'in'], ['wordsusd', 'Amount in Words ($)', '💵', 'intl']].forEach(([id, label, e, sys]) =>
  add({
    id, g: 'transform', menu: 'Words', label, e, k: 'number to text lakh crore invoice cheque', tip: 'Write each selected number in words in the next column (Lakh/Crore for ₹, Million for $).',
    run: async ctx => {
      const { sel } = await P.target(ctx, false);
      sel.load('values');
      await ctx.sync();
      sel.getLastColumn().getOffsetRange(0, 1).values = sel.values.map(r => [typeof r[r.length - 1] === 'number' ? P.words(r[r.length - 1], sys) : '']);
      await ctx.sync();
      return 'Words added in the next column';
    },
  }));

/* ================= FORMAT ================= */
Object.entries(THEMES).forEach(([id, [dark, light, name]]) => add({
  id: 'bt_' + id, g: 'format', menu: 'Beautify', label: name, e: '🪄', k: 'style table theme pretty colour', tip: `Instantly style the data as a clean ${name} table: header, zebra rows, borders, auto-fit. Replaces existing conditional formats.`,
  run: async ctx => {
    const { r } = await P.target(ctx);
    r.conditionalFormats.clearAll();
    r.format.verticalAlignment = 'Center';
    ['InsideHorizontal', 'EdgeBottom', 'EdgeTop', 'EdgeLeft', 'EdgeRight'].forEach(b => {
      const x = r.format.borders.getItem(b);
      x.style = 'Continuous';
      x.color = '#D5DBE5';
    });
    const h = r.getRow(0);
    P.header(h, dark);
    h.format.rowHeight = 26;
    const first = r.rowIndex + 1;
    P.cf(r, `=AND(ROW()>${first},MOD(ROW()-${first},2)=0)`, light);
    r.format.autofitColumns();
    await ctx.sync();
    return name + ' style applied';
  },
}));
add({
  id: 'zebra', g: 'format', label: 'Zebra Rows', e: '🦓', k: 'banded alternate rows', tip: 'Alternate row shading that survives sorting and filtering.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    P.cf(r, `=MOD(ROW()-${r.rowIndex},2)=0`, '#F1F3F8');
    await ctx.sync();
    return 'Zebra rows added';
  },
});
add({
  id: 'autofit', g: 'format', label: 'Auto-fit', e: '↔️', k: 'column width row height', tip: 'Auto-fit all columns and rows to their content.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    r.format.autofitColumns();
    r.format.autofitRows();
    await ctx.sync();
    return 'Auto-fitted';
  },
});
add({
  id: 'freeze', g: 'format', label: 'Freeze Header', e: '🧊', k: 'lock top row panes', tip: 'Freeze / unfreeze the top row so headers stay visible.',
  run: async ctx => {
    const f = ctx.workbook.worksheets.getActiveWorksheet().freezePanes;
    const loc = f.getLocationOrNullObject();
    loc.load('address');
    await ctx.sync();
    if (loc.isNullObject) { f.freezeRows(1); await ctx.sync(); return 'Header row frozen'; }
    f.unfreeze();
    await ctx.sync();
    return 'Panes unfrozen';
  },
});
Object.entries(NUMFMT).forEach(([id, [label, fmt]]) => add({
  id: 'nf_' + id, g: 'format', menu: 'Numbers', label, e: NUMFMT_EMOJI[id], k: 'number format currency', tip: `Format the selection as ${label}.`,
  run: async ctx => {
    const { r } = await P.target(ctx, false);
    r.numberFormat = P.grid(r.rowCount, r.columnCount, fmt);
    await ctx.sync();
    return label + ' format applied';
  },
}));
add({
  id: 'table', g: 'format', menu: 'More', label: 'Make Excel Table', e: '🗂️', k: 'filter ListObject', tip: 'Convert the data into a filterable Excel table.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    const t = r.worksheet.tables.add(r, true);
    t.style = 'TableStyleMedium2';
    await ctx.sync();
    return 'Table created';
  },
});
add({
  id: 'grid', g: 'format', menu: 'More', label: 'Toggle Gridlines', e: '📐', k: 'show hide grid', tip: 'Show or hide gridlines on this sheet.',
  run: async ctx => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load('showGridlines');
    await ctx.sync();
    ws.showGridlines = !ws.showGridlines;
    await ctx.sync();
    return ws.showGridlines ? 'Gridlines on' : 'Gridlines off';
  },
});
add({
  id: 'clearfmt', g: 'format', menu: 'More', label: 'Clear Formatting', e: '🧼', k: 'reset style', tip: 'Remove all formatting but keep the data.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    r.clear('Formats');
    await ctx.sync();
    return 'Formatting cleared';
  },
});

/* ================= FORMULAS ================= */
add({
  id: 'tovalues', g: 'formulas', label: 'Formulas → Values', e: '🧱', k: 'paste values freeze', tip: 'Replace formulas in the selection with their current results.',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    sel.load('values');
    await ctx.sync();
    sel.values = sel.values;
    await ctx.sync();
    return 'Converted to values';
  },
});
add({
  id: 'iferror', g: 'formulas', label: 'Wrap IFERROR', e: '🛡️', k: 'error handling', tip: 'Wrap every formula in the selection with IFERROR(…,"").',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    sel.load('formulas');
    await ctx.sync();
    let n = 0;
    sel.formulas.forEach((row, i) => row.forEach((f, j) => {
      if (typeof f === 'string' && f[0] === '=' && !/^=IFERROR\(/i.test(f)) { sel.getCell(i, j).formulas = [[`=IFERROR(${f.slice(1)},"")`]]; n++; }
    }));
    await ctx.sync();
    return n + ' formula(s) wrapped';
  },
});
add({
  id: 'hlformula', g: 'formulas', label: 'Highlight Formulas', e: '🔦', k: 'find formulas colour', tip: 'Colour every formula cell so hard-coded numbers stand out.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    const f = r.getSpecialCellsOrNullObject('Formulas');
    f.load('address');
    await ctx.sync();
    if (f.isNullObject) return 'No formulas found';
    f.format.fill.color = '#FFF4C2';
    await ctx.sync();
    return 'Formulas highlighted';
  },
});
add({ id: 'formulas', g: 'formulas', label: 'Formula Library', e: '📚', pane: 'formulas', k: 'xlookup vlookup sumifs help', tip: 'Searchable library of 40+ essential formulas with one-click insert.' });

/* ================= SHEETS ================= */
add({
  id: 'toc', g: 'sheets', label: 'Index Sheet', e: '🧭', k: 'table of contents hyperlinks navigate', tip: 'Create a clickable index of every sheet in the workbook.',
  run: async ctx => {
    const old = ctx.workbook.worksheets.getItemOrNullObject('Index');
    await ctx.sync();
    if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const names = await P.sheetNames(ctx);
    const ws = ctx.workbook.worksheets.add('Index');
    ws.position = 0;
    ws.showGridlines = false;
    ws.getRange('B2').values = [['Workbook Index']];
    ws.getRange('B2').format.font.size = 22;
    ws.getRange('B2').format.font.bold = true;
    ws.getRange('B2').format.font.color = '#4F46E5';
    names.forEach((n, i) => {
      const c = ws.getRange('B' + (4 + i));
      c.hyperlink = { documentReference: `'${n}'!A1`, textToDisplay: `${i + 1}.  ${n}`, screenTip: 'Go to ' + n };
    });
    ws.getRange('B:B').format.columnWidth = 260;
    ws.activate();
    await ctx.sync();
    return 'Index created';
  },
});
add({
  id: 'sortsheets', g: 'sheets', label: 'Sort Sheets', e: '🔤', k: 'order tabs alphabetical', tip: 'Sort all sheet tabs A → Z.',
  run: async ctx => {
    const w = ctx.workbook.worksheets;
    w.load('items/name');
    await ctx.sync();
    [...w.items].sort((a, b) => a.name.localeCompare(b.name)).forEach((s, i) => (s.position = i));
    await ctx.sync();
    return 'Sheets sorted';
  },
});
add({
  id: 'unhide', g: 'sheets', menu: 'Visibility', label: 'Unhide All Sheets', e: '👀', k: 'show hidden', tip: 'Make every hidden sheet visible.',
  run: async ctx => {
    const w = ctx.workbook.worksheets;
    w.load('items/visibility');
    await ctx.sync();
    w.items.forEach(s => (s.visibility = 'Visible'));
    await ctx.sync();
    return 'All sheets visible';
  },
});
add({
  id: 'rainbow', g: 'sheets', menu: 'Visibility', label: 'Rainbow Tab Colours', e: '🌈', k: 'colour tabs', tip: 'Give every sheet tab its own colour.',
  run: async ctx => {
    const w = ctx.workbook.worksheets;
    w.load('items');
    await ctx.sync();
    const pal = ['#EF4444', '#F59E0B', '#10B981', '#3B82F6', '#8B5CF6', '#EC4899', '#14B8A6', '#F97316'];
    w.items.forEach((s, i) => (s.tabColor = pal[i % pal.length]));
    await ctx.sync();
    return 'Tabs coloured';
  },
});
add({
  id: 'protectall', g: 'sheets', menu: 'Visibility', label: 'Protect All Sheets', e: '🔐', k: 'lock', tip: 'Protect every sheet (no password).',
  run: async ctx => {
    const w = ctx.workbook.worksheets;
    w.load('items');
    await ctx.sync();
    w.items.forEach(s => s.protection.protect());
    await ctx.sync();
    return 'All sheets protected';
  },
});
add({
  id: 'unprotectall', g: 'sheets', menu: 'Visibility', label: 'Unprotect All Sheets', e: '🔓', k: 'unlock', tip: 'Remove protection from all sheets that have no password.',
  run: async ctx => {
    const w = ctx.workbook.worksheets;
    w.load('items');
    await ctx.sync();
    w.items.forEach(s => s.protection.unprotect());
    await ctx.sync();
    return 'Protection removed';
  },
});
add({
  id: 'fromlist', g: 'sheets', label: 'Sheets from List', e: '🗒️', k: 'create many sheets names months', tip: 'Create one new sheet for every cell selected (uses cell text as the name).',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    sel.load('values');
    const taken = await P.sheetNames(ctx);
    let n = 0;
    sel.values.flat().filter(v => String(v).trim()).forEach(v => { ctx.workbook.worksheets.add(P.safeName(v, taken)); n++; });
    await ctx.sync();
    return n + ' sheet(s) created';
  },
});
add({
  id: 'splitby', g: 'sheets', label: 'Split by Column', e: '🪓', k: 'separate sheets per value region department', tip: 'Split the data into one sheet per unique value of the active cell\'s column (row 1 = header).',
  run: async ctx => {
    const { r, key } = await colKeyIdx(ctx);
    r.load('values');
    await ctx.sync();
    const [head, ...rows] = r.values;
    const groups = new Map();
    rows.forEach(x => { const k = String(x[key]); (groups.get(k) || groups.set(k, []).get(k)).push(x); });
    if (groups.size > 100) throw new Error(`${groups.size} groups - pick a column with fewer than 100 unique values`);
    const taken = await P.sheetNames(ctx);
    for (const [k, data] of groups) {
      const ws = ctx.workbook.worksheets.add(P.safeName(k || 'Blank', taken));
      const out = ws.getRangeByIndexes(0, 0, data.length + 1, head.length);
      out.values = [head, ...data];
      P.header(out.getRow(0));
      out.format.autofitColumns();
    }
    await ctx.sync();
    return `Split into ${groups.size} sheets`;
  },
});
add({
  id: 'mergesheets', g: 'sheets', label: 'Merge Sheets', e: '🧩', k: 'combine consolidate append', tip: 'Stack the data of all sheets (same columns) into one "Merged" sheet with a Source column.',
  run: async ctx => {
    const w = ctx.workbook.worksheets;
    w.load('items/name');
    await ctx.sync();
    const used = w.items.filter(s => !/^(Merged|Index|Audit Report)/.test(s.name)).map(s => { const u = s.getUsedRangeOrNullObject(true); u.load('values'); return [s.name, u]; });
    await ctx.sync();
    let head = null; const out = [];
    used.forEach(([name, u]) => {
      if (u.isNullObject) return;
      const [h, ...rows] = u.values;
      head = head || h;
      rows.forEach(x => out.push([...x.slice(0, head.length), name]));
    });
    if (!out.length) return 'No data to merge';
    const ws = await P.newSheet(ctx, 'Merged');
    const rg = ws.getRangeByIndexes(0, 0, out.length + 1, head.length + 1);
    rg.values = [[...head, 'Source'], ...out.map(x => Array.from({ length: head.length + 1 }, (_, i) => x[i] ?? ''))];
    P.header(rg.getRow(0));
    rg.format.autofitColumns();
    ws.activate();
    await ctx.sync();
    return `${out.length} rows merged from ${used.length} sheets`;
  },
});

/* ================= ANALYZE ================= */
add({
  id: 'hldupes', g: 'analyze', menu: 'Highlight', label: 'Duplicates', e: '👯', k: 'find repeated', tip: 'Highlight duplicate values in the selection (live conditional format).',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    const a = sel.address, tl = P.stripSheet(a).split(':')[0];
    P.cf(sel, `=AND(LEN(${tl})>0,COUNTIF(${P.abs(a)},${tl})>1)`, '#FFC7CE', '#9C0006');
    await ctx.sync();
    return 'Duplicates highlighted';
  },
});
add({
  id: 'hlblank', g: 'analyze', menu: 'Highlight', label: 'Blanks', e: '⬜', k: 'empty missing', tip: 'Highlight empty cells inside the data.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    P.cf(r, `=LEN(TRIM(${P.stripSheet(r.address).split(':')[0]}))=0`, '#FFE8A3');
    await ctx.sync();
    return 'Blanks highlighted';
  },
});
add({
  id: 'hlerror', g: 'analyze', menu: 'Highlight', label: 'Errors', e: '❌', k: 'N/A REF DIV/0', tip: 'Highlight cells with #N/A, #REF!, #DIV/0! etc.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    P.cf(r, `=ISERROR(${P.stripSheet(r.address).split(':')[0]})`, '#FFC7CE', '#9C0006');
    await ctx.sync();
    return 'Errors highlighted';
  },
});
add({
  id: 'outliers', g: 'analyze', menu: 'Highlight', label: 'Outliers', e: '🚨', k: 'anomaly unusual statistics iqr', tip: 'Mark numbers far outside the normal range (1.5 × IQR rule).',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    sel.load('values');
    await ctx.sync();
    const nums = sel.values.flat().filter(v => typeof v === 'number').sort((a, b) => a - b);
    if (nums.length < 4) return 'Select at least 4 numbers';
    const q = p => nums[Math.floor((nums.length - 1) * p)];
    const iqr = q(0.75) - q(0.25), lo = q(0.25) - 1.5 * iqr, hi = q(0.75) + 1.5 * iqr;
    let n = 0;
    sel.values.forEach((row, i) => row.forEach((v, j) => {
      if (typeof v === 'number' && (v < lo || v > hi)) { const c = sel.getCell(i, j); c.format.fill.color = '#FFD6A5'; c.format.font.color = '#B91C1C'; n++; }
    }));
    await ctx.sync();
    return `${n} outlier(s) marked (outside ${+lo.toFixed(2)} – ${+hi.toFixed(2)})`;
  },
});
add({
  id: 'heatmap', g: 'analyze', label: 'Heat Map', e: '🌡️', k: 'colour scale', tip: 'Red-yellow-green colour scale on the selected numbers.',
  run: async ctx => {
    const { r } = await P.target(ctx, false);
    const c = r.conditionalFormats.add('ColorScale');
    c.colorScale.criteria = {
      minimum: { type: 'LowestValue', color: '#F8696B' },
      midpoint: { type: 'Percentile', formula: '50', color: '#FFEB84' },
      maximum: { type: 'HighestValue', color: '#63BE7B' },
    };
    await ctx.sync();
    return 'Heat map applied';
  },
});
add({
  id: 'databars', g: 'analyze', label: 'Data Bars', e: '📶', k: 'in cell bars', tip: 'In-cell bars that show relative size.',
  run: async ctx => {
    const { r } = await P.target(ctx, false);
    const c = r.conditionalFormats.add('DataBar');
    c.dataBar.positiveFormat.fillColor = '#6366F1';
    await ctx.sync();
    return 'Data bars added';
  },
});
add({
  id: 'compare', g: 'analyze', label: 'Compare Columns', e: '⚖️', k: 'vlookup match difference two lists reconcile', tip: 'Select two columns: values missing from the other list are coloured (orange = only in 1st, blue = only in 2nd).',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    if (sel.columnCount < 2) return 'Select two columns';
    const a = sel.getColumn(0), b = sel.getColumn(1);
    a.load('values'); b.load('values');
    await ctx.sync();
    const norm = v => String(v).trim().toLowerCase();
    const sa = new Set(a.values.map(x => norm(x[0]))), sb = new Set(b.values.map(x => norm(x[0])));
    let na = 0, nb = 0;
    a.values.forEach((x, i) => { if (x[0] !== '' && !sb.has(norm(x[0]))) { a.getCell(i, 0).format.fill.color = '#FFD6A5'; na++; } });
    b.values.forEach((x, i) => { if (x[0] !== '' && !sa.has(norm(x[0]))) { b.getCell(i, 0).format.fill.color = '#BFDBFE'; nb++; } });
    await ctx.sync();
    return `${na} only in column 1, ${nb} only in column 2`;
  },
});
add({
  id: 'pivot', g: 'analyze', label: 'Quick Pivot', e: '🧮', k: 'pivot table summary group sum', tip: 'Auto-build a PivotTable: first text column × sum of first numeric column.',
  run: async ctx => {
    const { r } = await P.target(ctx);
    r.load('values');
    await ctx.sync();
    const [h, ...rows] = r.values;
    if (!rows.length) return 'Need a header and data rows';
    const isNum = j => rows.slice(0, 20).filter(x => typeof x[j] === 'number').length > Math.min(rows.length, 20) / 2;
    const num = h.findIndex((_, j) => isNum(j)), cat = h.findIndex((_, j) => !isNum(j));
    if (num < 0 || cat < 0) return 'Need one text and one numeric column';
    if (new Set(h).size !== h.length || h.some(x => x === '')) return 'Headers must be unique and non-empty';
    const ws = await P.newSheet(ctx, 'Pivot');
    const pt = ws.pivotTables.add('Pivot' + Date.now() % 100000, r, ws.getRange('A3'));
    pt.rowHierarchies.add(pt.hierarchies.getItem(String(h[cat])));
    pt.dataHierarchies.add(pt.hierarchies.getItem(String(h[num])));
    ws.activate();
    await ctx.sync();
    return `Pivot: ${h[cat]} × ${h[num]}`;
  },
});
[['ChartCol', 'Column Chart', '📊', 'ColumnClustered'], ['ChartLine', 'Line Chart', '📈', 'Line'], ['ChartBar', 'Bar Chart', '📉', 'BarClustered'],
['ChartPie', 'Pie Chart', '🥧', 'Pie'], ['ChartArea', 'Area Chart', '🏔️', 'Area']].forEach(([id, label, e, type]) => add({
  id: id.toLowerCase(), g: 'analyze', menu: 'Chart', label, e, k: 'graph visualize', tip: `Insert a ${label} from the selection, placed beside the data.`,
  run: async ctx => {
    const { r } = await P.target(ctx);
    const ch = r.worksheet.charts.add(type, r, 'Auto');
    ch.title.text = label.replace(' Chart', '') + ' overview';
    const s = r.getLastColumn().getOffsetRange(0, 1).getCell(0, 0);
    ch.setPosition(s, s.getOffsetRange(15, 7));
    await ctx.sync();
    return label + ' inserted';
  },
}));
[['selblank', 'Select Blanks', 'Blanks', '⬜'], ['selformula', 'Select Formulas', 'Formulas', '🧮'], ['selconst', 'Select Constants', 'Constants', '🔒'], ['selerror', 'Select Errors', 'Formulas', '❌']]
  .forEach(([id, label, type, e]) => add({
    id, g: 'analyze', menu: 'Select', label, e, k: 'go to special', tip: label + ' in the selection / used range.',
    run: async ctx => {
      const { r } = await P.target(ctx);
      const x = id === 'selerror' ? r.getSpecialCellsOrNullObject(type, 'Errors') : r.getSpecialCellsOrNullObject(type);
      x.load('address');
      await ctx.sync();
      if (x.isNullObject) return 'Nothing found';
      x.select();
      await ctx.sync();
      return 'Selected: ' + x.address.replace(/[^,]*!/g, '').slice(0, 60);
    },
  }));

/* ================= BUSINESS TEMPLATES ================= */
const INR = '[$₹-4009]#,##0.00';
async function tpl(ctx, o) {
  const ws = await P.newSheet(ctx, o.name);
  const n = o.cols.length, N = o.n || 30, top = 5; // header on row 5
  ws.showGridlines = false;
  const t = ws.getRange('A1');
  t.values = [[o.title]];
  t.format.font.size = 22; t.format.font.bold = true; t.format.font.color = '#4F46E5';
  const s = ws.getRange('A2');
  s.values = [[o.sub]];
  s.format.font.italic = true; s.format.font.color = '#6B7280';
  const rows = Array.from({ length: N }, (_, i) => Array.from({ length: n }, (_, j) => (o.calc && o.calc[j] ? o.calc[j](top + 1 + i) : (o.rows[i] || [])[j] ?? '')));
  const rg = ws.getRangeByIndexes(top - 1, 0, N + 1, n);
  rg.getRow(0).values = [o.cols];
  ws.getRangeByIndexes(top, 0, N, n).formulas = rows;
  const tb = ws.tables.add(rg, true);
  tb.name = o.name.replace(/\W/g, '') + 'T' + (Date.now() % 10000);
  tb.style = 'TableStyleLight9';
  P.header(rg.getRow(0));
  rg.getRow(0).format.rowHeight = 28;
  o.w.forEach((w, j) => (ws.getRangeByIndexes(0, j, 1, 1).getEntireColumn().format.columnWidth = w));
  Object.entries(o.fmt || {}).forEach(([j, f]) => { const c = ws.getRangeByIndexes(top, +j, N, 1); c.numberFormat = P.grid(N, 1, f); });
  Object.entries(o.lists || {}).forEach(([j, l]) => (ws.getRangeByIndexes(top, +j, N, 1).dataValidation.rule = { list: { inCellDropDown: true, source: l } }));
  if (o.sum) {
    ws.getRange('A3').values = [['TOTAL ▸']];
    ws.getRange('A3').format.font.bold = true;
    o.sum.forEach(j => {
      const c = ws.getRangeByIndexes(2, j, 1, 1), L = String.fromCharCode(65 + j);
      c.formulas = [[`=SUBTOTAL(109,${L}${top + 1}:${L}${top + N})`]];
      c.numberFormat = [[o.fmt && o.fmt[j] || '#,##0']];
      c.format.font.bold = true; c.format.font.color = '#4F46E5'; c.format.font.size = 13;
    });
  }
  if (o.after) o.after(ws, top, N);
  ws.freezePanes.freezeRows(top);
  ws.activate();
  ws.getRange('A' + (top + 1)).select();
  await ctx.sync();
  return o.name + ' template ready';
}
const TEMPLATES = {
  expenses: {
    name: 'Expenses', title: '💳 Expense Tracker', sub: 'Log every expense. Pick Category / Mode / Status from the drop-downs.',
    cols: ['Date', 'Category', 'Description', 'Paid By', 'Mode', 'Amount', 'Status'], w: [90, 120, 220, 120, 90, 110, 100],
    rows: [[45000, 'Travel', 'Client visit cab', 'Rahul', 'UPI', 1250, 'Pending'], [45001, 'Office', 'Stationery', 'Priya', 'Cash', 640, 'Approved']],
    fmt: { 0: 'dd-mmm-yy', 5: INR }, sum: [5],
    lists: { 1: 'Travel,Food,Office,Software,Marketing,Utilities,Other', 4: 'Cash,UPI,Card,Bank', 6: 'Pending,Approved,Reimbursed,Rejected' },
  },
  budget: {
    name: 'Budget', title: '🎯 Monthly Budget', sub: 'Variance and % used calculate automatically.', n: 20,
    cols: ['Category', 'Budget', 'Actual', 'Variance', '% Used'], w: [180, 120, 120, 120, 110],
    rows: [['Salaries', 500000, 480000], ['Marketing', 80000, 95000], ['Rent', 60000, 60000]],
    calc: { 3: r => `=IF(B${r}="","",B${r}-C${r})`, 4: r => `=IF(B${r}="","",IFERROR(C${r}/B${r},0))` },
    fmt: { 1: INR, 2: INR, 3: INR, 4: '0%' }, sum: [1, 2, 3],
    after: (ws, top, N) => { const c = ws.getRangeByIndexes(top, 4, N, 1).conditionalFormats.add('DataBar'); c.dataBar.positiveFormat.fillColor = '#6366F1'; },
  },
  inventory: {
    name: 'Inventory', title: '📦 Inventory', sub: 'Status flips to REORDER when stock reaches the reorder level.',
    cols: ['SKU', 'Item', 'Category', 'Qty', 'Reorder Level', 'Unit Cost', 'Stock Value', 'Status'], w: [90, 200, 120, 80, 110, 110, 120, 100],
    rows: [['SKU-001', 'Plywood 18mm', 'Board', 120, 50, 1450], ['SKU-002', 'Laminate 1mm', 'Surface', 30, 40, 980]],
    calc: { 6: r => `=IF(D${r}="","",D${r}*F${r})`, 7: r => `=IF(D${r}="","",IF(D${r}<=E${r},"REORDER","OK"))` },
    fmt: { 5: INR, 6: INR }, sum: [3, 6],
    after: (ws, top, N) => P.cf(ws.getRangeByIndexes(top, 7, N, 1), `=H${top + 1}="REORDER"`, '#FFC7CE', '#9C0006'),
  },
  attendance: {
    name: 'Attendance', title: '🗓️ Attendance Sheet', sub: 'P = Present, A = Absent, L = Leave, H = Half day. Totals are automatic.', n: 40,
    cols: ['Employee', 'Dept', ...Array.from({ length: 31 }, (_, i) => String(i + 1)), 'Present', 'Absent', 'Leave'],
    w: [150, 100, ...Array(31).fill(32), 70, 70, 70], rows: [['Rahul', 'Sales'], ['Priya', 'HR']],
    calc: { 33: r => `=IF(A${r}="","",COUNTIF(C${r}:AG${r},"P")+COUNTIF(C${r}:AG${r},"H")/2)`, 34: r => `=IF(A${r}="","",COUNTIF(C${r}:AG${r},"A"))`, 35: r => `=IF(A${r}="","",COUNTIF(C${r}:AG${r},"L"))` },
    after: (ws, top, N) => {
      const d = ws.getRangeByIndexes(top, 2, N, 31);
      d.dataValidation.rule = { list: { inCellDropDown: true, source: 'P,A,L,H' } };
      d.format.horizontalAlignment = 'Center';
      P.cf(d, `=C${top + 1}="A"`, '#FFC7CE', '#9C0006');
      P.cf(d, `=C${top + 1}="P"`, '#D1FAE5', '#065F46');
      P.cf(d, `=C${top + 1}="L"`, '#FEF3C7', '#92400E');
    },
  },
  project: {
    name: 'Project Tracker', title: '🚀 Project Tracker', sub: 'Track owners, deadlines and progress.', n: 40,
    cols: ['Task', 'Owner', 'Start', 'Due', 'Priority', 'Status', '% Complete', 'Notes'], w: [220, 110, 90, 90, 90, 110, 110, 240],
    rows: [['Kick-off meeting', 'Rahul', 45000, 45003, 'High', 'Done', 1], ['Prepare proposal', 'Priya', 45002, 45010, 'Medium', 'In Progress', 0.4]],
    fmt: { 2: 'dd-mmm-yy', 3: 'dd-mmm-yy', 6: '0%' },
    lists: { 4: 'High,Medium,Low', 5: 'To Do,In Progress,Blocked,Done' },
    after: (ws, top, N) => {
      const c = ws.getRangeByIndexes(top, 6, N, 1).conditionalFormats.add('DataBar'); c.dataBar.positiveFormat.fillColor = '#10B981';
      P.cf(ws.getRangeByIndexes(top, 5, N, 1), `=F${top + 1}="Blocked"`, '#FFC7CE', '#9C0006');
      P.cf(ws.getRangeByIndexes(top, 5, N, 1), `=F${top + 1}="Done"`, '#D1FAE5', '#065F46');
    },
  },
};
add({
  id: 'invoice', g: 'business', label: 'Invoice', e: '🧾', k: 'bill gst tax', tip: 'Create a professional GST-ready invoice with automatic totals.',
  run: async ctx => {
    const ws = await P.newSheet(ctx, 'Invoice');
    ws.showGridlines = false;
    const d = new Date(), dt = `=DATE(${d.getFullYear()},${d.getMonth() + 1},${d.getDate()})`;
    ws.getRange('A1:F6').formulas = [
      ['Your Company Name', '', '', '', '', 'INVOICE'], ['Address line · GSTIN · Phone', '', '', '', '', ''], ['', '', '', '', '', ''],
      ['BILL TO', '', '', '', 'Invoice No', 'INV-0001'], ['Client name', '', '', '', 'Date', dt], ['Client address', '', '', '', 'Due Date', '=F5+15'],
    ];
    ws.getRange('A1').format.font.size = 20; ws.getRange('A1').format.font.bold = true; ws.getRange('A1').format.font.color = '#4F46E5';
    ws.getRange('F1').format.font.size = 26; ws.getRange('F1').format.font.bold = true; ws.getRange('F1').format.font.color = '#4F46E5';
    ws.getRange('F1:F6').format.horizontalAlignment = 'Right';
    ws.getRange('A4').format.font.bold = true; ws.getRange('A4').format.font.color = '#6B7280';
    ws.getRange('E4:E6').format.font.bold = true;
    ws.getRange('F5:F6').numberFormat = [['dd-mmm-yyyy'], ['dd-mmm-yyyy']];
    ws.getRange('A8:F8').values = [['Item', 'Description', 'Qty', 'Rate', 'GST %', 'Amount']];
    P.header(ws.getRange('A8:F8'));
    ws.getRange('A8:F8').format.rowHeight = 28;
    const rows = Array.from({ length: 10 }, (_, i) => { const r = 9 + i; return [i === 0 ? 'Product A' : i === 1 ? 'Service B' : '', '', i === 0 ? 2 : i === 1 ? 1 : '', i === 0 ? 1500 : i === 1 ? 5000 : '', i < 2 ? 0.18 : '', `=IF(C${r}="","",C${r}*D${r}*(1+E${r}))`]; });
    ws.getRange('A9:F18').formulas = rows;
    ws.getRange('D9:D18').numberFormat = P.grid(10, 1, INR);
    ws.getRange('F9:F18').numberFormat = P.grid(10, 1, INR);
    ws.getRange('E9:E18').numberFormat = P.grid(10, 1, '0%');
    ['InsideHorizontal', 'EdgeBottom'].forEach(b => { const x = ws.getRange('A9:F18').format.borders.getItem(b); x.style = 'Continuous'; x.color = '#E5E7EB'; });
    ws.getRange('E20:F22').formulas = [['Subtotal', '=SUMPRODUCT(C9:C18,D9:D18)'], ['GST', '=SUMPRODUCT(C9:C18,D9:D18,E9:E18)'], ['TOTAL', '=SUM(F9:F18)']];
    ws.getRange('F20:F22').numberFormat = P.grid(3, 1, INR);
    ws.getRange('E20:F22').format.font.bold = true;
    ws.getRange('E22:F22').format.fill.color = '#EEF2FF'; ws.getRange('E22:F22').format.font.color = '#4F46E5'; ws.getRange('E22:F22').format.font.size = 14;
    ws.getRange('A25').values = [['Terms: Payment due within 15 days. Thank you for your business!']];
    ws.getRange('A25').format.font.italic = true; ws.getRange('A25').format.font.color = '#6B7280';
    [150, 220, 70, 110, 70, 130].forEach((w, j) => (ws.getRangeByIndexes(0, j, 1, 1).getEntireColumn().format.columnWidth = w));
    ws.activate();
    await ctx.sync();
    return 'Invoice ready - edit the grey placeholders';
  },
});
[['expenses', '💳'], ['budget', '🎯'], ['inventory', '📦'], ['attendance', '🗓️'], ['project', '🚀']].forEach(([id, e]) => add({
  id: 'tpl_' + id, g: 'business', menu: 'Templates', label: TEMPLATES[id].name, e, k: 'template ready made sheet', tip: TEMPLATES[id].sub,
  run: ctx => tpl(ctx, TEMPLATES[id]),
}));
add({
  id: 'dashboard', g: 'business', label: 'KPI Dashboard', e: '🏆', k: 'report summary cards chart executive', tip: 'Turn the current sheet\'s data into a KPI dashboard: 4 cards + chart.',
  run: async ctx => {
    const src = ctx.workbook.worksheets.getActiveWorksheet();
    const u = src.getUsedRange(true);
    u.load('values,rowCount,rowIndex,columnIndex');
    src.load('name');
    await ctx.sync();
    const [h, ...rows] = u.values;
    if (rows.length < 2) return 'Need a header and data rows';
    const isNum = j => rows.slice(0, 20).filter(x => typeof x[j] === 'number').length > Math.min(rows.length, 20) / 2;
    const num = h.findIndex((_, j) => isNum(j));
    if (num < 0) return 'No numeric column found';
    let cat = h.findIndex((_, j) => !isNum(j)); if (cat < 0) cat = 0;
    const L = j => { let s = '', x = u.columnIndex + j; do { s = String.fromCharCode(65 + (x % 26)) + s; x = Math.floor(x / 26) - 1; } while (x >= 0); return s; };
    const r1 = u.rowIndex + 2, r2 = u.rowIndex + u.rowCount, q = `'${src.name}'!`;
    const col = `${q}$${L(num)}$${r1}:$${L(num)}$${r2}`;
    const ws = await P.newSheet(ctx, 'Dashboard');
    ws.showGridlines = false;
    ws.getRange('A1:K30').format.fill.color = '#F5F7FB';
    ws.getRange('B1').values = [['📊 Dashboard · ' + h[num]]];
    ws.getRange('B1').format.font.size = 22; ws.getRange('B1').format.font.bold = true; ws.getRange('B1').format.font.color = '#312E81';
    const cards = [['RECORDS', `=COUNT(${col})`, '#4F46E5', '#,##0'], ['TOTAL', `=SUM(${col})`, '#059669', '#,##0'], ['AVERAGE', `=AVERAGE(${col})`, '#D97706', '#,##0.0'], ['HIGHEST', `=MAX(${col})`, '#DC2626', '#,##0']];
    cards.forEach(([lab, f, color, nf], i) => {
      const c = ws.getRangeByIndexes(2, 1 + i * 2, 2, 1);
      c.formulas = [[lab], [f]];
      c.format.fill.color = '#FFFFFF';
      c.getRow(0).format.font.size = 9; c.getRow(0).format.font.bold = true; c.getRow(0).format.font.color = '#6B7280';
      c.getRow(1).format.font.size = 24; c.getRow(1).format.font.bold = true; c.getRow(1).format.font.color = color;
      c.getRow(1).numberFormat = [[nf]];
      c.format.horizontalAlignment = 'Left'; c.format.indentLevel = 1;
      const b = c.format.borders.getItem('EdgeTop'); b.style = 'Continuous'; b.weight = 'Thick'; b.color = color;
      ws.getRangeByIndexes(0, 1 + i * 2, 1, 1).getEntireColumn().format.columnWidth = 140;
      ws.getRangeByIndexes(0, 2 + i * 2, 1, 1).getEntireColumn().format.columnWidth = 18;
    });
    ws.getRange('A:A').format.columnWidth = 18;
    const m = Math.min(rows.length, 12) + 1;
    const ch = ws.charts.add('ColumnClustered', src.getRangeByIndexes(u.rowIndex, u.columnIndex + num, m, 1), 'Columns');
    ch.series.getItemAt(0).setXAxisValues(src.getRangeByIndexes(u.rowIndex + 1, u.columnIndex + cat, m - 1, 1));
    ch.title.text = `${h[num]} by ${h[cat]} (first ${m - 1})`;
    ch.legend.visible = false;
    ch.setPosition('B7', 'I24');
    ws.activate();
    await ctx.sync();
    return 'Dashboard created';
  },
});

/* ================= PRO TOOLS ================= */
add({ id: 'palette', g: 'pro', label: 'Command Search', e: '🔎', pane: 'search', k: 'find any tool command palette', tip: 'Type what you want to do - find and run any of the tools instantly.' });
add({ id: 'insights', g: 'pro', label: 'Live Insights', e: '💡', pane: 'insights', k: 'stats sum average count status bar', tip: 'Live count, sum, average, median, min/max and more for your selection.' });
add({
  id: 'hidevals', g: 'pro', menu: 'Privacy', label: 'Hide Values', e: '🙈', k: 'mask confidential salary screen share', tip: 'Hide the displayed values of the selection (data stays intact) - ideal for screen sharing.',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    sel.numberFormat = P.grid(sel.rowCount, sel.columnCount, ';;;');
    await ctx.sync();
    return 'Values hidden (use Show Values to restore)';
  },
});
add({
  id: 'showvals', g: 'pro', menu: 'Privacy', label: 'Show Values', e: '🐵', k: 'unmask', tip: 'Restore General number format on the selection.',
  run: async ctx => {
    const { sel } = await P.target(ctx, false);
    sel.numberFormat = P.grid(sel.rowCount, sel.columnCount, 'General');
    await ctx.sync();
    return 'Values visible';
  },
});
add({
  id: 'audit', g: 'pro', label: 'Workbook Audit', e: '🩺', k: 'check health errors formulas hidden report', tip: 'Report on every sheet: size, formulas, errors, hidden state, tables, charts.',
  run: async ctx => {
    const w = ctx.workbook.worksheets;
    w.load('items/name,items/visibility');
    await ctx.sync();
    const info = w.items.map(s => { const u = s.getUsedRangeOrNullObject(true); u.load('address,rowCount,columnCount'); s.tables.load('count'); s.charts.load('count'); return { s, u }; });
    await ctx.sync();
    info.forEach(x => { if (!x.u.isNullObject && x.u.rowCount * x.u.columnCount <= 500000) x.u.load('formulas,valueTypes'); });
    await ctx.sync();
    const out = info.map(({ s, u }) => {
      if (u.isNullObject) return [s.name, s.visibility, 'empty', 0, 0, 0, s.tables.count, s.charts.count];
      const big = u.rowCount * u.columnCount > 500000;
      const f = big ? 'n/a' : u.formulas.flat().filter(x => typeof x === 'string' && x[0] === '=').length;
      const e = big ? 'n/a' : u.valueTypes.flat().filter(x => x === 'Error').length;
      return [s.name, s.visibility, P.stripSheet(u.address), u.rowCount * u.columnCount, f, e, s.tables.count, s.charts.count];
    });
    const ws = await P.newSheet(ctx, 'Audit Report');
    ws.showGridlines = false;
    ws.getRange('A1').values = [['🩺 Workbook Audit']];
    ws.getRange('A1').format.font.size = 20; ws.getRange('A1').format.font.bold = true; ws.getRange('A1').format.font.color = '#4F46E5';
    const rg = ws.getRangeByIndexes(2, 0, out.length + 1, 8);
    rg.values = [['Sheet', 'Visibility', 'Used range', 'Cells', 'Formulas', 'Errors', 'Tables', 'Charts'], ...out];
    P.header(rg.getRow(0));
    P.cf(ws.getRangeByIndexes(3, 5, out.length, 1), '=AND(ISNUMBER(F4),F4>0)', '#FFC7CE', '#9C0006');
    P.cf(ws.getRangeByIndexes(3, 1, out.length, 1), '=B4<>"Visible"', '#FFE8A3');
    rg.format.autofitColumns();
    ws.activate();
    await ctx.sync();
    const errs = out.reduce((a, x) => a + (+x[5] || 0), 0);
    return `${out.length} sheets audited · ${errs} error cell(s)`;
  },
});
add({ id: 'settings', g: 'pro', label: 'Settings', e: '⚙️', pane: 'settings', k: 'theme dark accent about help', tip: 'Choose accent colour, dark mode and view shortcuts & help.' });

if (typeof module !== 'undefined') module.exports = { GROUPS, MENUS, TOOLS, THEMES };
