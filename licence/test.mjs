// End-to-end test of the live licence API.
//   ADMIN_TOKEN=xxxx node licence/test.mjs
// Creates a licence called "TEST - automated" and deletes it at the end.
const BASE = process.env.LICENCE_URL || 'https://excel-100x-pro.vercel.app';
const TOKEN = process.env.ADMIN_TOKEN;
if (!TOKEN) { console.error('Set ADMIN_TOKEN first.'); process.exit(1); }

let pass = 0, fail = 0;
const post = async (path, body, hdr = {}) =>
  (await fetch(BASE + path, {
    method: 'POST', headers: { 'content-type': 'application/json', ...hdr }, body: JSON.stringify(body),
  })).json();
const lic = b => post('/api/licence', b);
const adm = b => post('/api/admin', b, { 'x-admin-token': TOKEN });

const check = (label, got, want) => {
  const ok = got === want;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${ok ? '' : `   got=${JSON.stringify(got)} want=${JSON.stringify(want)}`}`);
  ok ? pass++ : fail++;
};

const r = await adm({ action: 'issue', plan: 'Y5', customer: 'TEST - automated', email: 'test@example.com' });
if (!r.ok) { console.error('could not issue test licence:', r); process.exit(1); }
const KEY = r.licence.key;
console.log(`test licence ${KEY}  seats=${r.licence.seats}  expires=${String(r.licence.expires_at).slice(0, 10)}\n`);

// 1. five seats activate
for (let i = 1; i <= 5; i++) {
  const a = await lic({ action: 'activate', key: KEY, machine: 'TESTPC' + i, machine_name: 'Test PC ' + i, user: 'tester' });
  check(`activate PC${i}`, a.ok && a.seats_used, i);
}
// 2. sixth is refused
check('6th PC refused', (await lic({ action: 'activate', key: KEY, machine: 'TESTPC6' })).error, 'seat_limit');
// 3. re-activating an existing PC does not consume a seat
check('re-activate PC3 reuses seat', (await lic({ action: 'activate', key: KEY, machine: 'TESTPC3' })).seats_used, 5);
// 4. daily check works
check('check PC2 valid', (await lic({ action: 'check', key: KEY, machine: 'TESTPC2' })).ok, true);
// 5. machine list
const m = await lic({ action: 'machines', key: KEY });
check('machine list has 5', m.ok && m.machines.length, 5);
// 6. transfer: free PC1, activate PC6
check('release PC1', (await lic({ action: 'release', key: KEY, machine: 'TESTPC1' })).ok, true);
check('PC6 activates after release', (await lic({ action: 'activate', key: KEY, machine: 'TESTPC6' })).ok, true);
// 7. the released machine is locked out
check('released PC1 blocked', (await lic({ action: 'check', key: KEY, machine: 'TESTPC1' })).error, 'not_activated');
// 8. transfer rate limit (3 per 30 days; 1 used)
await lic({ action: 'release', key: KEY, machine: 'TESTPC2' });
await lic({ action: 'release', key: KEY, machine: 'TESTPC3' });
check('4th transfer refused', (await lic({ action: 'release', key: KEY, machine: 'TESTPC4' })).error, 'transfer_limit');
// 9. suspend / reactivate
await adm({ action: 'status', key: KEY, status: 'suspended' });
check('suspended key blocked', (await lic({ action: 'check', key: KEY, machine: 'TESTPC5' })).error, 'suspended');
await adm({ action: 'status', key: KEY, status: 'active' });
check('reactivated key works', (await lic({ action: 'check', key: KEY, machine: 'TESTPC5' })).ok, true);
// 10. unknown key and bad input
check('unknown key', (await lic({ action: 'activate', key: 'EXL-ZZZZ-ZZZZ-ZZZZ', machine: 'X' })).error, 'not_found');
check('empty key', (await lic({ action: 'activate', key: '', machine: 'X' })).error, 'bad_request');
check('lowercase key accepted', (await lic({ action: 'check', key: KEY.toLowerCase(), machine: 'TESTPC5' })).ok, true);
// 11. admin is protected
check('admin needs token', (await post('/api/admin', { action: 'list' })).message, 'Wrong admin token.');
// 12. renew
const rn = await adm({ action: 'renew', key: KEY, months: 12 });
check('renew adds a year', rn.ok && new Date(rn.licence.expires_at) > new Date(Date.now() + 700 * 864e5), true);
// 13. overview shows seats
const list = await adm({ action: 'list' });
const row = list.rows.find(x => x.key === KEY);
check('overview seat count', Number(row.seats_used), 3);

console.log('\ncleaning up…', (await adm({ action: 'delete', key: KEY })).message);
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
