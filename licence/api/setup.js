// POST /api/setup  header: x-admin-token      ONE-OFF: applies schema.sql to the database.
// Deployed only when build/deploy-web.js is run with --with-setup, so it is not live normally.
// The SQL is idempotent (create if not exists / create or replace), so re-running is safe.
import { Pool, neonConfig } from '@neondatabase/serverless';
import ws from 'ws';
import { SCHEMA } from './_schema.js';
import { timingSafeEqual } from 'node:crypto';

neonConfig.webSocketConstructor = ws;

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
  if (!process.env.DATABASE_URL) return res.status(503).json({ ok: false, message: 'DATABASE_URL is not set.' });

  const pool = new Pool({ connectionString: process.env.DATABASE_URL });
  try {
    await pool.query(SCHEMA);
    const plans = await pool.query('select code, name, seats, price_inr from plans order by seats');
    const funcs = await pool.query(
      `select routine_name from information_schema.routines
        where routine_schema = 'public' and routine_name in
              ('lic_activate','lic_check','lic_machines','lic_release','issue_licence','renew_licence')
        order by routine_name`);
    return res.json({ ok: true, plans: plans.rows, functions: funcs.rows.map(r => r.routine_name) });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: String(e.message || e).slice(0, 400) });
  } finally {
    await pool.end();
  }
}
