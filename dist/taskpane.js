/* Excel 100x Pro - task pane */
const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const store = (() => { try { return JSON.parse(localStorage.p100 || '{}'); } catch (e) { return {}; } })();
const save = () => { try { localStorage.p100 = JSON.stringify(store); } catch (e) { /* storage blocked */ } };

const ACCENTS = ['#4F46E5', '#0EA5E9', '#10B981', '#F97316', '#EC4899', '#7C3AED'];
function applyTheme() {
  document.documentElement.style.setProperty('--accent', store.accent || ACCENTS[0]);
  const dark = store.mode === 'dark' || (store.mode !== 'light' && matchMedia('(prefers-color-scheme: dark)').matches);
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
}

const TABS = [['search', '🔎', 'Search'], ['insights', '💡', 'Insights'], ['tools', '🛠️', 'Tools'], ['formulas', '📚', 'Formulas'], ['settings', '⚙️', 'Settings']];
let tab = 'search';

function show(t) {
  tab = t;
  $('#tabs').innerHTML = TABS.map(([id, e, l]) => `<button data-t="${id}" class="${id === t ? 'on' : ''}"><span>${e}</span>${l}</button>`).join('');
  $('#tabs').querySelectorAll('button').forEach(b => (b.onclick = () => show(b.dataset.t)));
  VIEWS[t]($('#view'));
}

/* ---------- Search / command palette ---------- */
const groupName = id => GROUPS.find(g => g.id === id).label;
function itemHTML(t) {
  return `<button class="item" data-id="${t.id}"><div class="e">${t.e}</div><div><b>${esc(t.label)}</b><small>${esc(t.tip)}</small></div><span class="badge">${esc(t.menu || groupName(t.g))}</span></button>`;
}
function runTool(t) {
  if (t.pane) return show(t.pane === 'search' ? 'search' : t.pane);
  store.recent = [t.id, ...(store.recent || []).filter(x => x !== t.id)].slice(0, 5);
  save();
  P.exec(t);
}
const VIEWS = {
  search(v) {
    v.innerHTML = `<div class="search"><i>🔎</i><input id="q" placeholder="What do you want to do?  e.g. remove duplicates" autocomplete="off"></div><div id="list"></div>`;
    const q = $('#q'), list = $('#list');
    let act = 0, shown = [];
    const draw = () => {
      const w = q.value.toLowerCase().split(/\s+/).filter(Boolean);
      let html;
      if (!w.length) {
        const rec = (store.recent || []).map(id => TOOLS.find(t => t.id === id)).filter(Boolean);
        shown = [...rec, ...TOOLS.filter(t => !rec.includes(t))];
        html = (rec.length ? '<h3>Recent</h3>' + rec.map(itemHTML).join('') : '') +
          GROUPS.map(g => `<h3>${g.e} ${g.label}</h3>` + TOOLS.filter(t => t.g === g.id).map(itemHTML).join('')).join('');
      } else {
        shown = TOOLS.filter(t => w.every(x => `${t.label} ${t.tip} ${t.k || ''} ${t.menu || ''} ${groupName(t.g)}`.toLowerCase().includes(x)));
        html = shown.length ? shown.map(itemHTML).join('') : '<div class="empty">😕 No tool matches.<br>Try "clean", "sheet" or "chart".</div>';
      }
      list.innerHTML = html;
      act = 0;
      mark();
    };
    const mark = () => list.querySelectorAll('.item').forEach((b, i) => b.classList.toggle('act', !!q.value && i === act));
    list.onclick = e => { const b = e.target.closest('.item'); if (b) runTool(TOOLS.find(t => t.id === b.dataset.id)); };
    q.oninput = draw;
    q.onkeydown = e => {
      const n = list.querySelectorAll('.item').length;
      if (e.key === 'ArrowDown') { act = Math.min(n - 1, act + 1); mark(); e.preventDefault(); }
      else if (e.key === 'ArrowUp') { act = Math.max(0, act - 1); mark(); e.preventDefault(); }
      else if (e.key === 'Enter' && shown[act]) runTool(shown[act]);
    };
    draw();
    q.focus();
  },

  /* ---------- Insights ---------- */
  insights(v) {
    v.innerHTML = `<div id="ins" class="grid"><div class="empty" style="grid-column:span 2">Select some cells…</div></div>`;
    refreshInsights();
  },

  /* ---------- Tools (forms) ---------- */
  tools(v) {
    const delims = [['Comma', ','], ['Semicolon', ';'], ['Space', ' '], ['Tab', '\t'], ['Pipe', '|'], ['Dash', '-'], ['Slash', '/']];
    const chips = id => `<div class="chips" id="${id}">${delims.map(([n, d], i) => `<button class="chip${i === 0 ? ' on' : ''}" data-d="${esc(d)}">${n}</button>`).join('')}</div>`;
    v.innerHTML = `
      <div class="card"><h2>✂️ Split column</h2><p>Select the cells of one column. Text is split into new columns on its right (existing data shifts, nothing is overwritten).</p>
        ${chips('sd')}<input class="field" id="sc" placeholder="…or type a custom separator"><button class="btn" id="dosplit">Split selected column</button></div>
      <div class="card"><h2>🔗 Merge columns</h2><p>Select 2+ columns. They are joined into a new column on the right.</p>
        ${chips('md')}<input class="field" id="mc" placeholder="…or type a custom separator"><button class="btn" id="domerge">Merge selected columns</button></div>
      <div class="card"><h2>🔁 Find &amp; replace in selection</h2><p>Plain-text replace limited to your selection (or whole sheet if one cell).</p>
        <input class="field" id="fi" placeholder="Find"><input class="field" id="re" placeholder="Replace with (empty = delete)">
        <label><input type="checkbox" id="cs"> Match case</label><br><br><button class="btn" id="dorep">Replace all</button></div>`;
    const pick = id => { const on = $('#' + id + ' .on'); const c = $('#' + (id === 'sd' ? 'sc' : 'mc')).value; return c || (on ? on.dataset.d : ','); };
    ['sd', 'md'].forEach(id => $('#' + id).onclick = e => { if (e.target.dataset.d !== undefined) { $('#' + id).querySelectorAll('.chip').forEach(c => c.classList.toggle('on', c === e.target)); $('#' + (id === 'sd' ? 'sc' : 'mc')).value = ''; } });
    $('#dosplit').onclick = () => P.exec(TOOLS.find(t => t.id === 'split'), { delim: pick('sd') });
    $('#domerge').onclick = () => P.exec(TOOLS.find(t => t.id === 'merge'), { delim: pick('md') });
    $('#dorep').onclick = () => {
      const f = $('#fi').value, r = $('#re').value, cs = $('#cs').checked;
      if (!f) return P.toast('Enter text to find', 'err');
      const re = new RegExp(f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), cs ? 'g' : 'gi');
      P.exec({
        run: async ctx => {
          const { r: rg } = await P.target(ctx);
          const n = await P.mapCells(ctx, rg, v => (typeof v === 'string' ? v.replace(re, () => r) : v));
          return n + ' cell(s) changed';
        },
      });
    };
  },

  /* ---------- Formula library ---------- */
  formulas(v) {
    v.innerHTML = `<div class="search"><i>🔎</i><input id="fq" placeholder="Search formulas… lookup, sum, date" autocomplete="off"></div><div id="fl"></div>`;
    const draw = () => {
      const w = $('#fq').value.toLowerCase().split(/\s+/).filter(Boolean);
      const r = FORMULAS.filter(x => w.every(k => (x[0] + x[1] + x[2]).toLowerCase().includes(k)));
      $('#fl').innerHTML = r.map((x, i) => `<div class="card"><h2>${esc(x[0])}</h2><p>${esc(x[1])}</p><code>${esc(x[2])}</code><button class="btn sm" data-i="${FORMULAS.indexOf(x)}">Insert in active cell</button></div>`).join('') || '<div class="empty">No formula found</div>';
    };
    $('#fq').oninput = draw;
    $('#fl').onclick = e => {
      const b = e.target.closest('button'); if (!b) return;
      P.exec({ run: async ctx => { ctx.workbook.getSelectedRange().getCell(0, 0).formulas = [[FORMULAS[b.dataset.i][2]]]; await ctx.sync(); return 'Formula inserted - adjust the ranges'; } });
    };
    draw();
    $('#fq').focus();
  },

  /* ---------- Settings ---------- */
  settings(v) {
    v.innerHTML = `
      <div class="card"><h2>🎨 Accent colour</h2><div class="sw">${ACCENTS.map(c => `<button data-c="${c}" style="background:${c}" class="${c === (store.accent || ACCENTS[0]) ? 'on' : ''}"></button>`).join('')}</div>
        <h2>🌗 Appearance</h2><div class="chips" id="md2">${['auto', 'light', 'dark'].map(m => `<button class="chip${(store.mode || 'auto') === m ? ' on' : ''}" data-m="${m}">${m}</button>`).join('')}</div></div>
      <div class="card"><h2>💡 Tips</h2><p>• Select <b>one cell</b> and most tools work on the <b>whole sheet</b>.<br>• Select cells first for precise control.<br>• Every action can be undone with <kbd>Ctrl</kbd>+<kbd>Z</kbd>.<br>• Use <b>Search</b> to find any tool by name or purpose.</p></div>
      <div class="card"><h2>🔒 Privacy</h2><p>Everything runs locally inside Excel. No data leaves your workbook, no telemetry, no sign-in.</p></div>
      <div class="card"><h2>ℹ️ About</h2><p>Excel 100x Pro · version 1.0 · ${TOOLS.length} tools</p></div>`;
    v.querySelectorAll('[data-c]').forEach(b => (b.onclick = () => { store.accent = b.dataset.c; save(); applyTheme(); show('settings'); }));
    v.querySelectorAll('[data-m]').forEach(b => (b.onclick = () => { store.mode = b.dataset.m; save(); applyTheme(); show('settings'); }));
  },
};

async function refreshInsights() {
  const el = $('#ins');
  if (tab !== 'insights' || !el) return;
  try {
    await Excel.run(async ctx => {
      const r = ctx.workbook.getSelectedRange();
      r.load('address,rowCount,columnCount');
      await ctx.sync();
      if (r.rowCount * r.columnCount > 500000) { el.innerHTML = '<div class="empty" style="grid-column:span 2">Selection too large for live stats</div>'; return; }
      r.load('values');
      await ctx.sync();
      const all = r.values.flat(), nums = all.filter(x => typeof x === 'number').sort((a, b) => a - b);
      const sum = nums.reduce((a, b) => a + b, 0), f = x => (Number.isInteger(x) ? x : +x.toFixed(2)).toLocaleString();
      const mid = nums.length ? (nums.length % 2 ? nums[(nums.length - 1) / 2] : (nums[nums.length / 2 - 1] + nums[nums.length / 2]) / 2) : null;
      const sd = nums.length > 1 ? Math.sqrt(nums.reduce((a, b) => a + (b - sum / nums.length) ** 2, 0) / (nums.length - 1)) : null;
      const st = (l, v, big) => `<div class="stat${big ? ' big' : ''}"><label>${l}</label><b>${v}</b></div>`;
      el.innerHTML = st(esc(r.address.split('!').pop()) + ' · SUM', nums.length ? f(sum) : '—', true) +
        st('Average', nums.length ? f(sum / nums.length) : '—') + st('Median', mid === null ? '—' : f(mid)) +
        st('Min', nums.length ? f(nums[0]) : '—') + st('Max', nums.length ? f(nums[nums.length - 1]) : '—') +
        st('Numbers', nums.length) + st('Non-empty', all.filter(x => x !== '').length) +
        st('Blank', all.filter(x => x === '').length) + st('Unique', new Set(all.filter(x => x !== '')).size) +
        st('Std dev', sd === null ? '—' : f(sd)) + st('Rows × Cols', `${r.rowCount}×${r.columnCount}`);
    });
  } catch (e) { /* host busy - next selection refreshes */ }
}

const FORMULAS = [
  ['XLOOKUP', 'Modern lookup, any direction (365)', '=XLOOKUP(E2,A:A,B:B,"Not found")'],
  ['VLOOKUP', 'Classic lookup in first column', '=VLOOKUP(E2,A:C,3,FALSE)'],
  ['INDEX + MATCH', 'Flexible lookup, works everywhere', '=INDEX(C:C,MATCH(E2,A:A,0))'],
  ['SUMIFS', 'Sum with several conditions', '=SUMIFS(C:C,A:A,"North",B:B,">=100")'],
  ['COUNTIFS', 'Count with several conditions', '=COUNTIFS(A:A,"North",B:B,">100")'],
  ['AVERAGEIFS', 'Average with conditions', '=AVERAGEIFS(C:C,A:A,"North")'],
  ['IF / IFS', 'Multiple conditions without nesting', '=IFS(A2>=90,"A",A2>=75,"B",A2>=50,"C",TRUE,"Fail")'],
  ['IFERROR', 'Hide errors with a fallback', '=IFERROR(A2/B2,0)'],
  ['SWITCH', 'Map values to labels', '=SWITCH(A2,1,"Low",2,"Medium",3,"High","Other")'],
  ['FILTER', 'Rows that match a condition (365)', '=FILTER(A2:C100,B2:B100>100,"None")'],
  ['UNIQUE', 'Distinct values (365)', '=UNIQUE(A2:A100)'],
  ['SORT / SORTBY', 'Sort a list dynamically (365)', '=SORT(A2:C100,2,-1)'],
  ['TEXTJOIN', 'Join text with a separator, skip blanks', '=TEXTJOIN(", ",TRUE,A2:A10)'],
  ['TEXTSPLIT', 'Split text into columns (365)', '=TEXTSPLIT(A2,",")'],
  ['LEFT / MID / RIGHT', 'Extract part of text', '=MID(A2,FIND("-",A2)+1,5)'],
  ['TRIM + CLEAN', 'Clean imported text', '=TRIM(CLEAN(A2))'],
  ['SUBSTITUTE', 'Replace text inside a cell', '=SUBSTITUTE(A2,"old","new")'],
  ['TEXT', 'Format a number or date as text', '=TEXT(A2,"dd-mmm-yyyy")'],
  ['CONCAT with &', 'Join values', '=A2&" "&B2'],
  ['LET', 'Name intermediate results (365)', '=LET(x,A2*B2,x+x*0.18)'],
  ['SUMPRODUCT', 'Multiply then sum arrays', '=SUMPRODUCT(B2:B10,C2:C10)'],
  ['RANK', 'Rank within a list', '=RANK(B2,B$2:B$100,0)'],
  ['LARGE / SMALL', 'Nth biggest or smallest', '=LARGE(B2:B100,3)'],
  ['ROUND', 'Round to decimals', '=ROUND(A2,2)'],
  ['MROUND', 'Round to nearest multiple', '=MROUND(A2,5)'],
  ['TODAY / NOW', 'Current date / time', '=TODAY()'],
  ['EOMONTH', 'Last day of month', '=EOMONTH(A2,0)'],
  ['EDATE', 'Same day, N months later', '=EDATE(A2,3)'],
  ['DATEDIF', 'Years / months / days between dates', '=DATEDIF(A2,TODAY(),"Y")'],
  ['NETWORKDAYS', 'Working days between dates', '=NETWORKDAYS(A2,B2)'],
  ['WORKDAY', 'Date after N working days', '=WORKDAY(A2,10)'],
  ['WEEKNUM', 'Week number of the year', '=WEEKNUM(A2)'],
  ['PMT', 'Loan EMI', '=PMT(9%/12,60,-500000)'],
  ['FV', 'Future value of savings', '=FV(8%/12,120,-5000)'],
  ['NPV / IRR', 'Investment returns', '=IRR(B2:B10)'],
  ['GST add', 'Price including 18% GST', '=A2*1.18'],
  ['GST remove', 'Base price from GST-inclusive amount', '=A2/1.18'],
  ['Percentage change', 'Growth between two values', '=(B2-A2)/A2'],
  ['Running total', 'Cumulative sum', '=SUM(B$2:B2)'],
  ['Count unique', 'Number of distinct values (365)', '=COUNTA(UNIQUE(A2:A100))'],
  ['Remove duplicates in formula', 'Distinct sorted list (365)', '=SORT(UNIQUE(A2:A100))'],
  ['ISNUMBER + SEARCH', 'Does text contain a word?', '=ISNUMBER(SEARCH("urgent",A2))'],
  ['HYPERLINK', 'Clickable link', '=HYPERLINK("https://example.com","Open")'],
  ['CHOOSECOLS', 'Pick columns from a range (365)', '=CHOOSECOLS(A2:E100,1,3)'],
];

Office.onReady(() => {
  try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, refreshInsights); } catch (e) { /* not in Excel */ }
});
applyTheme();
const go = new URLSearchParams(location.search).get('go');
show(TABS.some(t => t[0] === go) ? go : 'search');
