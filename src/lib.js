/* Excel 100x Pro - shared helpers (loaded by taskpane + commands) */
var P = {};

P.stripSheet = a => a.slice(a.lastIndexOf('!') + 1);
P.abs = a => P.stripSheet(a).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
P.grid = (r, c, v) => Array.from({ length: r }, () => Array(c).fill(v));
P.err = e => (e && e.code === 'ItemNotFound' ? 'Nothing found' : (e && e.message) || String(e));
P.serial = () => (Date.now() - new Date().getTimezoneOffset() * 60000) / 86400000 + 25569;

/** Selection; a single cell expands to the used range when `whole` is true. */
P.target = async (ctx, whole = true) => {
  const sel = ctx.workbook.getSelectedRange();
  sel.load('rowCount,columnCount,rowIndex,columnIndex,address');
  await ctx.sync();
  if (whole && sel.rowCount * sel.columnCount === 1) {
    const u = sel.worksheet.getUsedRangeOrNullObject(true);
    u.load('rowCount,columnCount,rowIndex,columnIndex,address');
    await ctx.sync();
    if (!u.isNullObject) return { r: u, sel, whole: true };
  }
  return { r: sel, sel, whole: false };
};

/** Apply fn(value,i,j) to every non-formula cell; write back only changed cells. */
P.mapCells = async (ctx, r, fn) => {
  r.load('values,formulas');
  await ctx.sync();
  let n = 0;
  r.values.forEach((row, i) => row.forEach((v, j) => {
    const f = r.formulas[i][j];
    if (typeof f === 'string' && f[0] === '=') return;
    const nv = fn(v, i, j);
    if (nv !== v) { r.getCell(i, j).values = [[nv]]; n++; }
  }));
  await ctx.sync();
  return n;
};

P.sheetNames = async ctx => {
  const w = ctx.workbook.worksheets;
  w.load('items/name');
  await ctx.sync();
  return w.items.map(s => s.name);
};

P.safeName = (base, taken) => {
  let b = String(base).replace(/[\\/?*[\]:]/g, '').trim().slice(0, 31) || 'Sheet';
  let n = b, i = 2;
  const lo = taken.map(t => t.toLowerCase());
  while (lo.includes(n.toLowerCase())) { const s = ` (${i++})`; n = b.slice(0, 31 - s.length) + s; }
  taken.push(n);
  return n;
};

P.newSheet = async (ctx, base) => {
  const name = P.safeName(base, await P.sheetNames(ctx));
  return ctx.workbook.worksheets.add(name);
};

P.cf = (r, formula, fill, font) => {
  const c = r.conditionalFormats.add('Custom');
  c.custom.rule.formula = formula;
  c.custom.format.fill.color = fill;
  if (font) c.custom.format.font.color = font;
};

P.header = (r, color = '#4F46E5') => {
  r.format.fill.color = color;
  r.format.font.color = '#FFFFFF';
  r.format.font.bold = true;
  r.format.verticalAlignment = 'Center';
};

/** Number -> words. sys: 'in' (Lakh/Crore, Rupees) or 'intl' (Million, Dollars). */
P.words = (num, sys) => {
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve',
    'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  const two = n => (n < 20 ? ones[n] : tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : ''));
  const three = n => (n > 99 ? ones[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' ' : '') : '') + (n % 100 ? two(n % 100) : '');
  const units = sys === 'in'
    ? [[1e7, 'Crore'], [1e5, 'Lakh'], [1e3, 'Thousand']]
    : [[1e12, 'Trillion'], [1e9, 'Billion'], [1e6, 'Million'], [1e3, 'Thousand']];
  const w = n => {
    if (n === 0) return 'Zero';
    const parts = [];
    for (const [u, name] of units) {
      const q = Math.floor(n / u);
      if (q) { parts.push(w(q) + ' ' + name); n %= u; }
    }
    if (n) parts.push(three(n));
    return parts.join(' ');
  };
  if (!isFinite(num) || Math.abs(num) >= 1e15) return '';
  const neg = num < 0, cents = Math.round(Math.abs(num) * 100);
  const whole = Math.floor(cents / 100), frac = cents % 100;
  const [big, small, tail] = sys === 'in' ? ['Rupees', 'Paise', ' Only'] : ['Dollars', 'Cents', ''];
  return (neg ? 'Minus ' : '') + big + ' ' + w(whole) + (frac ? ' and ' + w(frac) + ' ' + small : '') + tail;
};

/** Result message: in-pane snackbar, or a small auto-closing dialog from the ribbon. */
P.toast = (m, kind = 'ok') => {
  const el = document.getElementById('toast');
  if (el) {
    el.textContent = m;
    el.className = 'toast show ' + kind;
    clearTimeout(P._t);
    P._t = setTimeout(() => (el.className = 'toast'), 3500);
    return;
  }
  try {
    const base = location.href.replace(/[^/]*$/, '');
    Office.context.ui.displayDialogAsync(base + 'toast.html?k=' + kind + '&m=' + encodeURIComponent(m),
      { height: 12, width: 24, displayInIframe: false },
      r => r.status === 'succeeded' &&
        r.value.addEventHandler(Office.EventType.DialogMessageReceived, () => r.value.close()));
  } catch (e) { /* no UI available */ }
};

P.exec = async (tool, args) => {
  try {
    const msg = await Excel.run(ctx => tool.run(ctx, args));
    P.toast(msg || 'Done', 'ok');
    return true;
  } catch (e) {
    console.error(e);
    P.toast(P.err(e), 'err');
    return false;
  }
};
