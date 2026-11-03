-- Excel 100x Pro - licensing database (Neon Postgres)
-- Run once: Neon Console -> your project -> SQL Editor -> paste -> Run.
-- Nothing connects to this database except the Vercel API, so there are no public grants.

create table if not exists plans (
  code      text primary key,
  name      text not null,
  seats     int  not null check (seats > 0),
  price_inr numeric(10,2) not null,
  days      int  not null default 365,
  active    boolean not null default true
);

create table if not exists licences (
  key            text primary key,
  plan_code      text not null references plans(code),
  seats          int  not null check (seats > 0),
  customer_name  text,
  customer_email text,
  issued_at      timestamptz not null default now(),
  expires_at     timestamptz not null,
  status         text not null default 'active'
                 check (status in ('active','suspended','cancelled')),
  notes          text
);

create table if not exists activations (
  id           bigserial primary key,
  licence_key  text not null references licences(key) on delete cascade,
  machine_id   text not null,
  machine_name text,
  user_name    text,
  activated_at timestamptz not null default now(),
  last_seen_at timestamptz not null default now(),
  released_at  timestamptz,
  status       text not null default 'active' check (status in ('active','released'))
);

-- one live activation per machine per licence; released rows stay as history
create unique index if not exists activations_one_active_per_machine
  on activations (licence_key, machine_id) where status = 'active';
create index if not exists activations_by_key on activations (licence_key);

insert into plans (code, name, seats, price_inr, days) values
  ('Y5',   'Yearly - 5 computers',   5, 999, 365),
  ('Y10',  'Yearly - 10 computers', 10, 999, 365),
  ('TRIAL','14-day trial',           1,   0,  14)
on conflict (code) do nothing;

-- how many seats one licence may free in a rolling 30 days (dead-PC transfers)
create or replace function transfer_limit() returns int language sql immutable as $$ select 3 $$;

-- ---------------------------------------------------------------- used by the API

-- Claims a seat for this machine. Locks the licence row so two PCs cannot take the
-- last seat at the same time.
create or replace function lic_activate(
  p_key text, p_machine text, p_machine_name text default null, p_user text default null)
returns json language plpgsql as $$
declare l licences%rowtype; used int; pname text;
begin
  if coalesce(p_key,'') = '' or coalesce(p_machine,'') = '' then
    return json_build_object('ok', false, 'error', 'bad_request',
      'message', 'Licence key and machine id are required.');
  end if;

  select * into l from licences where key = upper(trim(p_key)) for update;
  if not found then
    return json_build_object('ok', false, 'error', 'not_found',
      'message', 'This licence key was not recognised. Please check it and try again.');
  end if;
  if l.status <> 'active' then
    return json_build_object('ok', false, 'error', l.status,
      'message', 'This licence is ' || l.status || '. Please contact support.');
  end if;
  if l.expires_at < now() then
    return json_build_object('ok', false, 'error', 'expired',
      'expires', to_char(l.expires_at,'YYYY-MM-DD'),
      'message', 'This licence expired on ' || to_char(l.expires_at,'DD Mon YYYY') || '. Please renew to continue.');
  end if;

  select name into pname from plans where code = l.plan_code;

  -- already on this machine: just refresh it, no seat used
  update activations
     set last_seen_at = now(),
         machine_name = coalesce(p_machine_name, machine_name),
         user_name    = coalesce(p_user, user_name)
   where licence_key = l.key and machine_id = p_machine and status = 'active';
  if found then
    select count(*) into used from activations where licence_key = l.key and status = 'active';
    return json_build_object('ok', true, 'status', 'active',
      'expires', to_char(l.expires_at,'YYYY-MM-DD'), 'seats_used', used,
      'seats_total', l.seats, 'plan', pname, 'message', 'Already activated on this computer.');
  end if;

  select count(*) into used from activations where licence_key = l.key and status = 'active';
  if used >= l.seats then
    return json_build_object('ok', false, 'error', 'seat_limit',
      'seats_used', used, 'seats_total', l.seats,
      'message', 'All ' || l.seats || ' computers on this licence are in use. Release one to activate this computer.');
  end if;

  insert into activations (licence_key, machine_id, machine_name, user_name)
  values (l.key, p_machine, p_machine_name, p_user);

  return json_build_object('ok', true, 'status', 'active',
    'expires', to_char(l.expires_at,'YYYY-MM-DD'), 'seats_used', used + 1,
    'seats_total', l.seats, 'plan', pname, 'message', 'Activated.');
end $$;

-- Daily re-check from the add-in. Fails if the seat was released or the licence lapsed.
create or replace function lic_check(p_key text, p_machine text)
returns json language plpgsql as $$
declare l licences%rowtype; hit boolean;
begin
  select * into l from licences where key = upper(trim(p_key));
  if not found then
    return json_build_object('ok', false, 'error', 'not_found', 'message', 'Licence key not recognised.');
  end if;
  if l.status <> 'active' then
    return json_build_object('ok', false, 'error', l.status,
      'message', 'This licence is ' || l.status || '.');
  end if;
  if l.expires_at < now() then
    return json_build_object('ok', false, 'error', 'expired',
      'expires', to_char(l.expires_at,'YYYY-MM-DD'),
      'message', 'Licence expired on ' || to_char(l.expires_at,'DD Mon YYYY') || '.');
  end if;

  update activations set last_seen_at = now()
   where licence_key = l.key and machine_id = p_machine and status = 'active'
  returning true into hit;
  if hit is null then
    return json_build_object('ok', false, 'error', 'not_activated',
      'message', 'This computer is not activated on the licence. It may have been released.');
  end if;

  return json_build_object('ok', true, 'status', 'active',
    'expires', to_char(l.expires_at,'YYYY-MM-DD'),
    'seats_total', l.seats, 'message', 'Licence valid.');
end $$;

-- Machines currently holding a seat, so the add-in can offer one to replace.
create or replace function lic_machines(p_key text)
returns json language plpgsql as $$
declare rows json;
begin
  if not exists (select 1 from licences where key = upper(trim(p_key))) then
    return json_build_object('ok', false, 'error', 'not_found', 'message', 'Licence key not recognised.');
  end if;
  select coalesce(json_agg(json_build_object(
           'machine_id', machine_id,
           'name',       coalesce(machine_name, '(unknown)'),
           'user',       coalesce(user_name, ''),
           'last_seen',  to_char(last_seen_at, 'YYYY-MM-DD'))
         order by last_seen_at desc), '[]'::json)
    into rows from activations where licence_key = upper(trim(p_key)) and status = 'active';
  return json_build_object('ok', true, 'machines', rows);
end $$;

-- Frees a seat so a replacement computer can take it. Works when the old PC is dead,
-- because it is called from the new one. Rate limited to stop a key floating freely.
create or replace function lic_release(p_key text, p_machine text)
returns json language plpgsql as $$
declare l licences%rowtype; recent int; lim int := transfer_limit();
begin
  select * into l from licences where key = upper(trim(p_key)) for update;
  if not found then
    return json_build_object('ok', false, 'error', 'not_found', 'message', 'Licence key not recognised.');
  end if;

  select count(*) into recent from activations
   where licence_key = l.key and status = 'released' and released_at > now() - interval '30 days';
  if recent >= lim then
    return json_build_object('ok', false, 'error', 'transfer_limit',
      'message', 'This licence has already moved ' || lim || ' computers in the last 30 days. Please contact support.');
  end if;

  update activations set status = 'released', released_at = now()
   where licence_key = l.key and machine_id = p_machine and status = 'active';
  if not found then
    return json_build_object('ok', false, 'error', 'not_activated',
      'message', 'That computer is not using a seat.');
  end if;
  return json_build_object('ok', true, 'message', 'Seat released. You can now activate the new computer.');
end $$;

-- ---------------------------------------------------------------- admin

-- select * from issue_licence('Y10', 'Acme Traders', 'acme@example.com');
create or replace function issue_licence(
  p_plan text, p_customer text, p_email text default null, p_notes text default null)
returns licences language plpgsql as $$
declare p plans%rowtype; k text; row licences%rowtype;
        alphabet text := 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';  -- no I/O/0/1
begin
  select * into p from plans where code = upper(trim(p_plan));
  if not found then
    raise exception 'Unknown plan %. Available: %', p_plan, (select string_agg(code, ', ') from plans);
  end if;
  loop
    k := 'EXL';
    for i in 1..12 loop
      if i % 4 = 1 then k := k || '-'; end if;
      k := k || substr(alphabet, 1 + floor(random() * length(alphabet))::int, 1);
    end loop;
    exit when not exists (select 1 from licences where key = k);
  end loop;
  insert into licences (key, plan_code, seats, customer_name, customer_email, expires_at, notes)
  values (k, p.code, p.seats, p_customer, p_email, now() + (p.days || ' days')::interval, p_notes)
  returning * into row;
  return row;
end $$;

-- select * from renew_licence('EXL-XXXX-XXXX-XXXX', 12);
create or replace function renew_licence(p_key text, p_months int default 12)
returns licences language plpgsql as $$
declare row licences%rowtype;
begin
  update licences
     set expires_at = greatest(expires_at, now()) + (p_months || ' months')::interval,
         status = 'active'
   where key = upper(trim(p_key)) returning * into row;
  if not found then raise exception 'Licence % not found', p_key; end if;
  return row;
end $$;

create or replace view overview as
select l.key, l.customer_name, l.customer_email, p.name as plan, l.seats,
       count(a.id) filter (where a.status = 'active') as seats_used,
       l.status, l.expires_at::date as expires,
       (l.expires_at::date - current_date) as days_left
  from licences l
  join plans p on p.code = l.plan_code
  left join activations a on a.licence_key = l.key
 group by l.key, l.customer_name, l.customer_email, p.name, l.seats, l.status, l.expires_at
 order by l.issued_at desc;
