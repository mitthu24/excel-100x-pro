// node build/build.js [--url https://host] [--icons]   -> dist/ + Excel-100x-Pro.zip
const fs = require('fs'), path = require('path'), cp = require('child_process');
const { GROUPS, MENUS, TOOLS } = require('../src/tools.js');

const root = path.join(__dirname, '..'), src = path.join(root, 'src'), dist = path.join(root, 'dist');
const arg = n => { const i = process.argv.indexOf(n); return i < 0 ? null : process.argv[i + 1] || true; };
const BASE = (arg('--url') || 'https://localhost:3100').replace(/\/$/, '');
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/* ---- icons: key -> emoji ---- */
const iconMap = {};
GROUPS.forEach(g => (iconMap['g_' + g.id] = g.e));
Object.entries(MENUS).forEach(([k, e]) => (iconMap['m_' + k] = e));
TOOLS.forEach(t => (iconMap['t_' + t.id] = t.e));
if (arg('--icons')) {
  fs.writeFileSync(path.join(__dirname, 'icons.json'), JSON.stringify(iconMap));
  cp.execFileSync('python', [path.join(__dirname, 'gen_icons.py')], { stdio: 'inherit' });
}

/* ---- manifest ---- */
const images = new Map(), urls = new Map(), shorts = new Map(), longs = new Map();
const S = (id, v) => (shorts.set(id, v), id);
const L = (id, v) => (longs.set(id, v), id);
const U = (id, v) => (urls.set(id, v), id);
const icon = key => [16, 32, 80].map(s => {
  const id = `${key}_${s}`;
  images.set(id, `${BASE}/assets/${id}.png`);
  return `<bt:Image size="${s}" resid="${id}"/>`;
}).join('');
const head = (id, label, tip, key) =>
  `<Label resid="${S(id + '.l', label)}"/><Supertip><Title resid="${id}.l"/><Description resid="${L(id + '.d', tip)}"/></Supertip><Icon>${icon(key)}</Icon>`;
const action = t => t.pane
  ? `<Action xsi:type="ShowTaskpane"><TaskpaneId>Pro100xPane</TaskpaneId><SourceLocation resid="${U('pane_' + t.pane, `${BASE}/taskpane.html?go=${t.pane}`)}"/></Action>`
  : `<Action xsi:type="ExecuteFunction"><FunctionName>${t.id}</FunctionName></Action>`;
const button = (t, tag) => `<${tag} id="b_${t.id}"${tag === 'Control' ? ' xsi:type="Button"' : ''}>${head('b_' + t.id, t.label, t.tip, 't_' + t.id)}${action(t)}</${tag}>`;

const groupsXml = GROUPS.map(g => {
  const done = new Set(), controls = [];
  TOOLS.filter(t => t.g === g.id).forEach(t => {
    if (!t.menu) return controls.push(button(t, 'Control'));
    if (done.has(t.menu)) return;
    done.add(t.menu);
    const items = TOOLS.filter(x => x.g === g.id && x.menu === t.menu).map(x => button(x, 'Item')).join('');
    const id = `m_${g.id}_${t.menu}`.slice(0, 30);
    controls.push(`<Control xsi:type="Menu" id="${id}">${head(id, t.menu, `${t.menu} tools`, 'm_' + t.menu)}<Items>${items}</Items></Control>`);
  });
  return `<Group id="g_${g.id}"><Label resid="${S('g_' + g.id + '.l', g.label)}"/><Icon>${icon('g_' + g.id)}</Icon>${controls.join('')}</Group>`;
}).join('');

// ribbon limits: ids <= 32 chars, labels <= 32, descriptions <= 250
for (const [id] of [...shorts, ...longs, ...urls, ...images]) if (id.length > 32) throw new Error('resid too long: ' + id);
for (const [id, v] of shorts) if (v.length > 32) throw new Error('label too long: ' + v);
for (const [id, v] of longs) if (v.length > 250) throw new Error('tip too long: ' + id);

const res = (tag, m, el) => `<bt:${tag}>${[...m].map(([id, v]) => `<bt:${el} id="${id}" DefaultValue="${esc(v)}"/>`).join('')}</bt:${tag}>`;
S('Tab.l', '100x Pro');
U('Commands.Url', `${BASE}/commands.html`);
images.set('logo_32', `${BASE}/assets/logo_32.png`);
images.set('logo_64', `${BASE}/assets/logo_64.png`);

const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>c6f3a1d2-7b4e-4f1a-9c3d-5e8a2b7d1f40</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Excel 100x Pro</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Excel 100x Pro"/>
  <Description DefaultValue="${TOOLS.length} productivity tools for Excel: clean data, reshape, format, audit and create business templates in one click."/>
  <IconUrl DefaultValue="${BASE}/assets/logo_32.png"/>
  <HighResolutionIconUrl DefaultValue="${BASE}/assets/logo_64.png"/>
  <SupportUrl DefaultValue="${BASE}/taskpane.html"/>
  <Hosts><Host Name="Workbook"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.9"><Set Name="ExcelApi"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="${BASE}/taskpane.html"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts><Host xsi:type="Workbook"><DesktopFormFactor>
      <FunctionFile resid="Commands.Url"/>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <CustomTab id="Pro100x.Tab">${groupsXml}<Label resid="Tab.l"/></CustomTab>
      </ExtensionPoint>
    </DesktopFormFactor></Host></Hosts>
    <Resources>${res('Images', images, 'Image')}${res('Urls', urls, 'Url')}${res('ShortStrings', shorts, 'String')}${res('LongStrings', longs, 'String')}</Resources>
  </VersionOverrides>
</OfficeApp>
`;

/* ---- dist ---- */
fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(dist, 'assets'), { recursive: true });
fs.readdirSync(src).filter(f => fs.statSync(path.join(src, f)).isFile()).forEach(f => fs.copyFileSync(path.join(src, f), path.join(dist, f)));
fs.readdirSync(path.join(src, 'assets')).forEach(f => fs.copyFileSync(path.join(src, 'assets', f), path.join(dist, 'assets', f)));
fs.writeFileSync(path.join(dist, 'manifest.xml'), manifest);

const missing = [...images.keys()].filter(k => !fs.existsSync(path.join(dist, 'assets', k + '.png')));
if (missing.length) throw new Error('missing icons (run with --icons): ' + missing.slice(0, 5));

/* ---- zip ---- */
const zip = path.join(root, 'Excel-100x-Pro.zip');
fs.rmSync(zip, { force: true });
const stage = path.join(root, '.stage', 'Excel-100x-Pro');
fs.rmSync(path.join(root, '.stage'), { recursive: true, force: true });
fs.mkdirSync(stage, { recursive: true });
['README.md', 'INSTALL-XLAM.md', 'Excel-100x-Pro.xlam', 'START-LOCAL.cmd', 'dist', 'docs', 'src', 'build', 'server', 'xlam'].forEach(n => {
  if (fs.existsSync(path.join(root, n))) fs.cpSync(path.join(root, n), path.join(stage, n), { recursive: true, filter: s => !/node_modules|icons\.json|[\\/]out[\\/]?$|[\\/]out[\\/]/.test(s) });
});
// bsdtar (ships with Windows 10+) writes real zips; GNU tar on PATH would not
const tar = process.platform === 'win32' ? path.join(process.env.SystemRoot, 'System32', 'tar.exe') : 'zip';
cp.execFileSync(tar, process.platform === 'win32' ? ['-a', '-cf', 'Excel-100x-Pro.zip', '-C', '.stage', 'Excel-100x-Pro'] : ['-qr', '../Excel-100x-Pro.zip', 'Excel-100x-Pro'], { cwd: process.platform === 'win32' ? root : path.join(root, '.stage') });
fs.rmSync(path.join(root, '.stage'), { recursive: true, force: true });
console.log(`manifest: ${TOOLS.length} tools, ${groupsXml.split('<Group ').length - 1} groups, base ${BASE}\nzip: ${zip}`);
