# Licence system

**Status: live and working.** The database, API and admin page are set up and tested.
Nothing below needs doing again — it is here so you know how it fits together and how to
rebuild it if you ever move hosts.

| What | Where |
|---|---|
| Admin page | <https://excel-100x-pro.vercel.app/admin.html> |
| Pricing page | <https://excel-100x-pro.vercel.app/buy.html> |
| Database | Neon `excel-100xpro`, linked to the Vercel project |
| Plans | ₹999 / year for 5 computers · free 14-day trial (1 computer) |

## How it works

```
Excel add-in  ──POST──>  excel-100x-pro.vercel.app/api/licence  ──>  Neon Postgres
  (ProLicence.bas)            (licence/api/licence.js)              (licence/schema.sql)

You            ──────>  excel-100x-pro.vercel.app/admin.html   ──>  /api/admin  ──> Neon
```

- Each PC sends a **machine id**: a hash of the hardware UUID plus the computer name. The raw
  hardware id never leaves the PC.
- The server counts seats. A 10-computer key allows 10 live activations.
- **Dead PC:** on the new computer, enter the key, pick the old PC from the list, and its seat is
  freed. The old PC stops at its next daily check. Limited to 3 moves per 30 days.
- The add-in checks in **once a day** and keeps working **offline for up to 14 days**.

## How it was set up (for reference, or to rebuild elsewhere)

1. **Database** — Vercel project → Storage → Create Database → Neon. Vercel then sets
   `DATABASE_URL` and the other Postgres variables on the project by itself.
2. **Tables** — [`schema.sql`](schema.sql) applied to that database. Vercel masks the Neon
   password, so it was applied by deploying a one-off endpoint that runs server-side:
   ```bash
   node build/deploy-web.js --with-setup
   cd .deploy && vercel deploy --prod --yes --scope smart-bi-studio --name excel-100x-pro
   curl -s -X POST https://excel-100x-pro.vercel.app/api/setup -H "x-admin-token: YOUR-TOKEN" -d '{}'
   node build/deploy-web.js        # plain build drops /api/setup again
   ```
   The endpoint is **not** deployed normally. The SQL is idempotent, so re-running is safe.
   You can also just paste `schema.sql` into the Neon SQL editor.
3. **Admin password** — `ADMIN_TOKEN` set on the Vercel project for all three environments.

### Health check

```bash
curl -s -X POST https://excel-100x-pro.vercel.app/api/licence \
  -H "content-type: application/json" \
  -d '{"action":"check","key":"EXL-TEST-TEST-TEST","machine":"abc"}'
```

Expected: `{"ok":false,"error":"not_found","message":"Licence key not recognised."}`
If it says *not set up yet*, `DATABASE_URL` is missing or a redeploy has not finished.

## Day-to-day: issuing keys

Go to **https://excel-100x-pro.vercel.app/admin.html**, enter your `ADMIN_TOKEN`, and you can:

- **Issue** a key — pick the plan, type the customer name, get a key like `EXL-K7PM-3QRT-9WDN`
- **See** every licence, how many seats are used, and days left
- **Renew** a licence by any number of months
- **Suspend** a licence (non-payment) and reactivate it later
- **Computers** — see each activated PC and release a seat yourself

Email the key to the customer. They open Excel → **100x Pro Tools → Licence** → paste it.

## Changing prices or plans

You sell one paid plan: **₹999 a year for 5 computers**, plus a free 14-day single-PC trial.

Prices live in the `plans` table and on the pricing page `licence/buy.html`. To change the price:

```sql
update plans set price_inr = 1299 where code = 'Y5';
```

To add a bigger tier (it appears in the admin drop-down straight away):

```sql
insert into plans (code, name, seats, price_inr, days)
values ('Y10', 'Yearly - 10 computers', 10, 1799, 365);
```

Changing a plan never affects licences already issued; each licence keeps the seat count and
expiry date it was given. Remember to add the tier to `buy.html` as well.

## What this protects against, and what it does not

**It stops** a key being shared beyond its seat count, use after expiry, and use by a PC whose
seat you revoked.

**It does not stop** a determined person who edits the add-in. A `.xlam` is a VBA file; VBA project
passwords can be removed with freely available tools, and the licence check can then be deleted.
This is true of every VBA add-in sold anywhere. Treat the system as a way to run a clean
subscription business with honest customers, not as copy protection.

Stronger options, if that ever matters: compile the tools into a COM add-in (.NET/VSTO, much harder
to patch), or move the valuable work server-side so the add-in is only a front end.

## Files

| File | What it is |
|---|---|
| `schema.sql` | Database tables, seat logic, admin functions |
| `api/licence.js` | Public endpoint the add-in calls |
| `api/admin.js` | Private endpoint, guarded by `ADMIN_TOKEN` |
| `api/setup.js` | One-off migration runner, deployed only with `--with-setup` |
| `admin.html` | Your licence manager |
| `buy.html` | Public pricing page |
| `../xlam/ProLicence.bas` | Add-in side: machine id, activation, daily check, transfer |

## Tests

```bash
ADMIN_TOKEN=your-token node licence/test.mjs
```

[`test.mjs`](test.mjs) was run against the live API. It covers: five seats
activating, the sixth being refused, re-activation not consuming a seat, the daily check,
the machine list, releasing a seat and reusing it, a released PC being locked out, the
three-transfers-per-month cap, suspend and reactivate, unknown and empty keys, lowercase keys,
admin needing a token, renewal, and the seat count in the overview. All passed.

If you change `schema.sql` or the API, re-run those checks before deploying. They create a
licence named **TEST - automated** and delete it at the end.
