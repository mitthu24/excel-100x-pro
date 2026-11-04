// Assembles everything Vercel serves into .deploy/ : the web add-in, the licence API,
// the pricing page and the admin page.   node build/deploy-web.js   then deploy .deploy/
const fs = require('fs'), path = require('path');

const root = path.join(__dirname, '..'), out = path.join(root, '.deploy');
// empty it rather than remove it: Windows often holds a lock on the directory itself
fs.mkdirSync(out, { recursive: true });
for (const f of fs.readdirSync(out)) {
  if (f === '.vercel' || f === 'node_modules') continue;   // keep the project link and installs
  fs.rmSync(path.join(out, f), { recursive: true, force: true });
}

fs.cpSync(path.join(root, 'dist'), out, { recursive: true });            // web add-in + manifest
fs.cpSync(path.join(root, 'licence', 'api'), path.join(out, 'api'), { recursive: true });
['buy.html', 'admin.html', 'package.json'].forEach(f =>
  fs.copyFileSync(path.join(root, 'licence', f), path.join(out, f)));

// one-off migration endpoint, only when asked for
if (process.argv.includes('--with-setup')) {
  const sqlText = fs.readFileSync(path.join(root, 'licence', 'schema.sql'), 'utf8');
  fs.writeFileSync(path.join(out, 'api', '_schema.js'),
    `export const SCHEMA = ${JSON.stringify(sqlText)};`);
  console.log('including /api/setup (remove it with a plain deploy afterwards)');
} else {
  fs.rmSync(path.join(out, 'api', 'setup.js'), { force: true });
}

const xlam = path.join(root, 'Excel-100x-Pro.xlam');                      // direct download
if (fs.existsSync(xlam)) fs.copyFileSync(xlam, path.join(out, 'Excel-100x-Pro.xlam'));

console.log(`.deploy ready (${fs.readdirSync(out).length} entries)
next:  cd .deploy && vercel deploy --prod --yes --scope smart-bi-studio --name excel-100x-pro`);
