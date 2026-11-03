// Static HTTPS server for ../dist on https://localhost:3100 (trusted dev cert from office-addin-dev-certs).
const https = require('https'), fs = require('fs'), path = require('path'), os = require('os');
const dir = path.join(os.homedir(), '.office-addin-dev-certs');
const root = path.join(__dirname, '..', 'dist');
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.xml': 'text/xml' };

https.createServer({ key: fs.readFileSync(path.join(dir, 'localhost.key')), cert: fs.readFileSync(path.join(dir, 'localhost.crt')) }, (req, res) => {
  const p = path.normalize(path.join(root, decodeURIComponent(new URL(req.url, 'https://x').pathname)));
  if (!p.startsWith(root) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); return res.end('Not found'); }
  res.writeHead(200, { 'Content-Type': types[path.extname(p)] || 'application/octet-stream', 'Access-Control-Allow-Origin': '*', 'Cache-Control': 'no-cache' });
  fs.createReadStream(p).pipe(res);
}).listen(3100, () => console.log('Excel 100x Pro running at https://localhost:3100  (keep this window open)'));
