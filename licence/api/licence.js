// POST /api/licence  { action, key, machine, machine_name, user }
// Public endpoint used by the add-in. All logic and locking live in the SQL functions.
import { neon } from '@neondatabase/serverless';

let _sql;
const db = () => {
  if (!process.env.DATABASE_URL) throw new Error('not_configured');
  return (_sql ??= neon(process.env.DATABASE_URL));
};
const s = (v, max) => (typeof v === 'string' ? v.trim().slice(0, max) : null);

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'method', message: 'POST only.' });

  try {
    const b = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : req.body || {};
    const key = s(b.key, 32);
    const machine = s(b.machine, 64);
    let out;

    switch (b.action) {
      case 'activate':
        out = await db()`select lic_activate(${key}, ${machine}, ${s(b.machine_name, 64)}, ${s(b.user, 64)}) as r`;
        break;
      case 'check':
        out = await db()`select lic_check(${key}, ${machine}) as r`;
        break;
      case 'machines':
        out = await db()`select lic_machines(${key}) as r`;
        break;
      case 'release':
        out = await db()`select lic_release(${key}, ${machine}) as r`;
        break;
      default:
        return res.status(400).json({ ok: false, error: 'action', message: 'Unknown action.' });
    }
    return res.status(200).json(out[0].r);
  } catch (e) {
    console.error(e);
    if (e.message === 'not_configured')
      return res.status(503).json({ ok: false, error: 'setup',
        message: 'The licence server is not set up yet. Please contact support.' });
    return res.status(500).json({ ok: false, error: 'server', message: 'Licence server error. Please try again.' });
  }
}
