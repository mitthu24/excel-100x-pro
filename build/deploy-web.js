// Assembles everything Vercel serves into .deploy/ : the web add-in, the licence API,
// the pricing page and the admin page.   node build/deploy-web.js   then deploy .deploy/
const fs = require('fs'), path = require('path');

const root = path.join(__dirname, '..'), out = path.join(root, '.deploy');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });

fs.cpSync(path.join(root, 'dist'), out, { recursive: true });            // web add-in + manifest
fs.cpSync(path.join(root, 'licence', 'api'), path.join(out, 'api'), { recursive: true });
['buy.html', 'admin.html', 'package.json'].forEach(f =>
  fs.copyFileSync(path.join(root, 'licence', f), path.join(out, f)));

const xlam = path.join(root, 'Excel-100x-Pro.xlam');                      // direct download
if (fs.existsSync(xlam)) fs.copyFileSync(xlam, path.join(out, 'Excel-100x-Pro.xlam'));

console.log(`.deploy ready (${fs.readdirSync(out).length} entries)
next:  cd .deploy && vercel deploy --prod --yes --scope smart-bi-studio --name excel-100x-pro`);
