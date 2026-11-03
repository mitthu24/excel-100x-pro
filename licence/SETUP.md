# Licence system — setup (about 10 minutes, one time)

Everything is built and deployed. It needs a database and two settings before it will run.

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

## Step 1 — create the database

1. In Vercel, open the **excel-100x-pro** project → **Storage** → **Create Database** → **Neon**.
   (Your Neon org is already linked to Vercel, so this just adds a database.)
2. Vercel sets `DATABASE_URL` on the project automatically. Nothing to copy.

If you prefer the Neon console, create a database there and add its pooled connection string to
Vercel as `DATABASE_URL` (Settings → Environment Variables → Production).

## Step 2 — create the tables

Open the Neon console → your database → **SQL Editor** → paste the whole of
[`schema.sql`](schema.sql) → **Run**. It creates the tables, the plans and the functions.

## Step 3 — set your admin password

Vercel → project → **Settings → Environment Variables** → add:

| Name | Value |
|---|---|
| `ADMIN_TOKEN` | a long random password you invent |

Then **Deployments → ⋯ → Redeploy** so both variables take effect.

## Step 4 — check it

```bash
curl -s -X POST https://excel-100x-pro.vercel.app/api/licence \
  -H "content-type: application/json" \
  -d '{"action":"check","key":"EXL-TEST-TEST-TEST","machine":"abc"}'
```

Expected: `{"ok":false,"error":"not_found","message":"Licence key not recognised."}`
If it says *not set up yet*, `DATABASE_URL` is missing or the redeploy has not finished.

## Day-to-day: issuing keys

Go to **https://excel-100x-pro.vercel.app/admin.html**, enter your `ADMIN_TOKEN`, and you can:

- **Issue** a key — pick the plan, type the customer name, get a key like `EXL-K7PM-3QRT-9WDN`
- **See** every licence, how many seats are used, and days left
- **Renew** a licence by any number of months
- **Suspend** a licence (non-payment) and reactivate it later
- **Computers** — see each activated PC and release a seat yourself

Email the key to the customer. They open Excel → **100x Pro Tools → Licence** → paste it.

## Changing prices or plans

Prices live in the `plans` table and in `licence/buy.html`. To change one:

```sql
update plans set price_inr = 1999 where code = 'Y10';
```

To add a plan:

```sql
insert into plans (code, name, seats, price_inr, days)
values ('Y25', 'Yearly - 25 computers', 25, 4999, 365);
```

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
| `admin.html` | Your licence manager |
| `buy.html` | Public pricing page |
| `../xlam/ProLicence.bas` | Add-in side: machine id, activation, daily check, transfer |
