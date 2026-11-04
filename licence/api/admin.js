// POST /api/admin  { action, ... }   header: x-admin-token
// Private: issue, renew, list, set status, release a seat. Guarded by ADMIN_TOKEN.
import { neon } from '@neondatabase/serverless';
import { timingSafeEqual } from 'node:crypto';

let _sql;
const db = () => {
  if (!process.env.DATABASE_URL) throw new Error('not_configured');
  return (_sql ??= neon(process.env.DATABASE_URL));
};
const s = (v, max) => (typeof v === 'string' ? v.trim().slice(0, max) : null);

const tokenOk = given => {
  const want = process.env.ADMIN_TOKEN || '';
  if (!want || typeof given !== 'string') return false;
  const a = Buffer.from(given), b = Buffer.from(want);
  return a.length === b.length && timingSafeEqual(a, b);
};

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, message: 'POST only.' });
  if (!tokenOk(req.headers['x-admin-token'])) return res.status(401).json({ ok: false, message: 'Wrong admin token.' });

  try {
    const b = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : req.body || {};
    const key = s(b.key, 32);

    switch (b.action) {
      case 'list':
        return res.json({ ok: true, rows: await db()`select * from overview limit 500` });

      case 'plans':
        return res.json({ ok: true, rows: await db()`select * from plans where active order by seats` });

      case 'issue': {
        if (!s(b.customer, 120)) return res.json({ ok: false, message: 'Customer name is required.' });
        const r = await db()`select to_jsonb(issue_licence(${s(b.plan, 16)}, ${s(b.customer, 120)},
                                                           ${s(b.email, 160)}, ${s(b.notes, 400)})) as l`;
        return res.json({ ok: true, licence: r[0].l });
      }

      case 'renew': {
        const r = await db()`select to_jsonb(renew_licence(${key}, ${Number(b.months) || 12})) as l`;
        return res.json({ ok: true, licence: r[0].l });
      }

      case 'status': {
        const st = s(b.status, 16);
        if (!['active', 'suspended', 'cancelled'].includes(st)) return res.json({ ok: false, message: 'Bad status.' });
        await db()`update licences set status = ${st} where key = ${key}`;
        return res.json({ ok: true, message: `Licence set to ${st}.` });
      }

      case 'seats':
        return res.json({
          ok: true,
          rows: await db()`select machine_id, machine_name, user_name, activated_at, last_seen_at
                            from activations where licence_key = ${key} and status = 'active'
                           order by last_seen_at desc`,
        });

      case 'delete':   // for a key issued by mistake; activations go with it
        await db()`delete from licences where key = ${key}`;
        return res.json({ ok: true, message: 'Licence deleted.' });

      case 'kick':
        await db()`update activations set status = 'released', released_at = now()
                   where licence_key = ${key} and machine_id = ${s(b.machine, 64)} and status = 'active'`;
        return res.json({ ok: true, message: 'Seat released.' });

      default:
        return res.status(400).json({ ok: false, message: 'Unknown action.' });
    }
  } catch (e) {
    console.error(e);
    if (e.message === 'not_configured')
      return res.status(503).json({ ok: false, message: 'DATABASE_URL is not set on this Vercel project yet.' });
    return res.status(500).json({ ok: false, message: String(e.message || e).slice(0, 300) });
  }
}
