-- ============================================================================
-- Capital Edge Stellar - Supabase schema (multi-user pass)
--
-- Run this once in your Supabase project's SQL Editor (Dashboard -> SQL
-- Editor -> New query -> paste this whole file -> Run).
--
-- What changed from the single-user version: companies used to belong to
-- exactly one owner. Now a company has a company_members table - multiple
-- real logins, each with a role - and company_invitations lets an Admin
-- invite someone by email without needing a backend email service (they
-- share a link; the invited person's role is applied the moment they sign
-- up through it).
--
-- Model: one row per company, holding its entire app state as JSONB (the
-- exact same shape the app already reads/writes as one JSON object today).
-- This keeps almost all of the app's business logic untouched - only how
-- that JSON object is loaded, saved, and who's allowed to touch it changes.
-- A fully normalized schema (separate tables per record type, so a
-- Bookkeeper's browser never even receives Payroll data) is the real
-- long-term architecture and is worth doing later - this pass gets you
-- genuine multi-user logins and database-enforced membership/read-write
-- control now, without waiting for that larger rewrite.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. companies: one row per company, holding all of its data. owner_id is
--    kept only as a record of who originally created it - it no longer
--    controls access on its own; company_members does that now.
-- ----------------------------------------------------------------------------
create table if not exists companies (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  name text not null default 'My Company',
  data jsonb not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists companies_owner_id_idx on companies(owner_id);

-- ----------------------------------------------------------------------------
-- 2. company_members: who can actually access a company, and what role they
--    hold. This is the real, database-enforced source of truth for access -
--    the Users & Roles screen in the app is what an Admin uses to manage it.
-- ----------------------------------------------------------------------------
-- email and name are duplicated here on purpose - auth.users isn't
-- directly queryable from the client for privacy reasons, so this is what
-- lets the Users & Roles screen show who's actually on the team without
-- needing a separate backend endpoint just to look up a name.
create table if not exists company_members (
  id uuid primary key default gen_random_uuid(),
  company_id uuid not null references companies(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  email text not null,
  name text not null default '',
  role text not null default 'admin' check (role in ('admin', 'accountant', 'bookkeeper', 'viewer')),
  created_at timestamptz not null default now(),
  unique (company_id, user_id)
);

create index if not exists company_members_company_id_idx on company_members(company_id);
create index if not exists company_members_user_id_idx on company_members(user_id);

-- profiles: kept from the original schema as a quick "which company is my
-- home company" lookup for a user with exactly one company. If someone is
-- ever a member of more than one (accepted an invite to a second company),
-- this still points at whichever one they were LAST added to - a company
-- switcher for genuinely multi-company users is a real, separate feature
-- this schema doesn't attempt to solve yet.
create table if not exists profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  company_id uuid not null references companies(id) on delete cascade,
  created_at timestamptz not null default now()
);

-- ----------------------------------------------------------------------------
-- 3. company_invitations: how an Admin invites someone without a backend
--    email service. The Admin creates a row here (via the app), shares the
--    resulting link (yoursite.com/?invite=TOKEN) with the invited person
--    directly - email, WhatsApp, whatever - and when that person signs up
--    through the link, the trigger below reads the token, links them to
--    this company with the stated role, and marks the invite used, instead
--    of creating them a brand new empty company.
-- ----------------------------------------------------------------------------
create table if not exists company_invitations (
  id uuid primary key default gen_random_uuid(),
  company_id uuid not null references companies(id) on delete cascade,
  email text not null,
  role text not null check (role in ('admin', 'accountant', 'bookkeeper', 'viewer')),
  token uuid not null default gen_random_uuid() unique,
  created_by uuid references auth.users(id),
  used boolean not null default false,
  created_at timestamptz not null default now()
);

create index if not exists company_invitations_token_idx on company_invitations(token);

-- ----------------------------------------------------------------------------
-- 4. Row Level Security - this is what actually prevents one company's
--    members from ever seeing another company's data, enforced by the
--    database itself, not just the app's own code. It also now enforces
--    something the UI alone never could: a Viewer's browser genuinely
--    cannot write a change to the database, even if someone bypassed the
--    app entirely and called the API directly.
-- ----------------------------------------------------------------------------
alter table companies enable row level security;
alter table company_members enable row level security;
alter table company_invitations enable row level security;
alter table profiles enable row level security;

-- A user can see a company if they're a member of it, in any role.
create policy "Members can view their company"
  on companies for select
  using (exists (select 1 from company_members where company_id = companies.id and user_id = auth.uid()));

-- A user can save changes to a company only if they're a member AND their
-- role isn't Viewer. This is real, database-level enforcement of "Viewers
-- can't write" - the one role restriction that's fully secure today,
-- because it's a simple yes/no on the whole row. A Bookkeeper being
-- write-blocked from Payroll specifically still relies on the app's own UI
-- (see the note at the top of this file about why, and what closes that gap).
create policy "Non-viewer members can update their company"
  on companies for update
  using (exists (
    select 1 from company_members
    where company_id = companies.id and user_id = auth.uid() and role <> 'viewer'
  ));

-- Company rows are only ever created via handle_new_user() below, never
-- inserted directly by the client - this stops a malicious client from
-- creating extra companies or attaching to someone else's.

-- A member can see who else is on their own company.
create policy "Members can view their company's member list"
  on company_members for select
  using (exists (select 1 from company_members m2 where m2.company_id = company_members.company_id and m2.user_id = auth.uid()));

-- Only an Admin of the company can add, change, or remove members.
create policy "Admins can add members"
  on company_members for insert
  with check (exists (
    select 1 from company_members
    where company_id = company_members.company_id and user_id = auth.uid() and role = 'admin'
  ));

create policy "Admins can change member roles"
  on company_members for update
  using (exists (
    select 1 from company_members m2
    where m2.company_id = company_members.company_id and m2.user_id = auth.uid() and m2.role = 'admin'
  ));

create policy "Admins can remove members"
  on company_members for delete
  using (exists (
    select 1 from company_members m2
    where m2.company_id = company_members.company_id and m2.user_id = auth.uid() and m2.role = 'admin'
  ));

-- Only an Admin of the company can create an invitation for it.
create policy "Admins can create invitations"
  on company_invitations for insert
  with check (exists (
    select 1 from company_members
    where company_id = company_invitations.company_id and user_id = auth.uid() and role = 'admin'
  ));

-- Admins can see (and revoke, by deleting) their own company's invitations.
create policy "Admins can view their company's invitations"
  on company_invitations for select
  using (exists (
    select 1 from company_members
    where company_id = company_invitations.company_id and user_id = auth.uid() and role = 'admin'
  ));

create policy "Admins can revoke invitations"
  on company_invitations for delete
  using (exists (
    select 1 from company_members
    where company_id = company_invitations.company_id and user_id = auth.uid() and role = 'admin'
  ));

-- A user can only see their own profile row.
create policy "Users can view their own profile"
  on profiles for select
  using (id = auth.uid());

-- ----------------------------------------------------------------------------
-- 5. lookup_invitation: lets the signup page show "You've been invited to
--    join [Company] as [Role]" BEFORE the person has an account or is
--    logged in - which normal RLS wouldn't allow, since they're anonymous
--    at that point. This function runs with elevated privilege but only
--    ever returns the company name and role for a valid, unused token -
--    never anything else about the company.
-- ----------------------------------------------------------------------------
create or replace function lookup_invitation(invite_token uuid)
returns table (company_name text, role text)
language plpgsql
security definer set search_path = public
as $$
begin
  return query
    select c.name, i.role
    from company_invitations i
    join companies c on c.id = i.company_id
    where i.token = invite_token and i.used = false;
end;
$$;

-- ----------------------------------------------------------------------------
-- 6. handle_new_user: runs automatically every time someone signs up via
--    Supabase Auth. Branches two ways:
--
--    a) Signed up through an invite link (raw_user_meta_data has an
--       invite_token that matches a real, unused invitation) - link them to
--       that EXISTING company with the invited role. No new company, no
--       empty data - they join what's already there.
--
--    b) A normal signup - create their own brand new company, starting
--       completely EMPTY (no demo data), and make them its Admin.
--
--    IMPORTANT: the placeholder JSONB below must be replaced with the real
--    output of emptyCompanyData() from src/emptyCompanyData.js before this
--    is useful - see DEPLOYMENT-GUIDE.md for the exact command. Keeping it
--    as a SQL default means new signups don't depend on any client-side
--    code to get set up correctly.
-- ----------------------------------------------------------------------------
create or replace function handle_new_user()
returns trigger
language plpgsql
security definer set search_path = public
as $$
declare
  new_company_id uuid;
  invite_token_text text;
  invite_row company_invitations;
begin
  invite_token_text := new.raw_user_meta_data->>'invite_token';

  -- Validate the shape before casting to uuid - a malformed or tampered
  -- token should just be ignored (falling through to a normal signup),
  -- not throw a cast error that fails the whole signup.
  if invite_token_text is not null and invite_token_text ~* '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$' then
    select * into invite_row from company_invitations
      where token = invite_token_text::uuid and used = false
      limit 1;
  end if;

  if invite_row.id is not null then
    -- Path (a): join the existing company that invited them.
    insert into company_members (company_id, user_id, email, name, role)
      values (invite_row.company_id, new.id, new.email, coalesce(new.raw_user_meta_data->>'name', ''), invite_row.role);
    update company_invitations set used = true where id = invite_row.id;
    insert into profiles (id, company_id) values (new.id, invite_row.company_id)
      on conflict (id) do update set company_id = excluded.company_id;
  else
    -- Path (b): brand new company, empty, this person is its Admin.
    insert into companies (owner_id, name, data)
    values (
      new.id,
      coalesce(new.raw_user_meta_data->>'company_name', 'My Company'),
      '{"settings":{"companyName":"My Company","currencyCode":"NGN","mode":"light","accent":"#2563EB","reportOptions":{"showCodes":false,"hideZeroLines":true},"userName":"","accountingBasis":"accrual","payrollCountry":"NG","industry":"general","country":"NG","corporateTaxRate":30,"eclRates":{"Current":0.5,"1-30 days":1,"31-60 days":5,"61-90 days":15,"90+ days":40},"lockDate":null},"accounts":[{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1000","code":"1000","name":"Main Operating Account","type":"asset","subtype":"bank","category":"Bank","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1010","code":"1010","name":"Reserve Account","type":"asset","subtype":"bank","category":"Bank","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1050","code":"1050","name":"Petty Cash","type":"asset","subtype":"current","category":"Cash","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1100","code":"1100","name":"Accounts Receivable","type":"asset","subtype":"current","category":"Accounts Receivable","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1200","code":"1200","name":"Finished Goods Inventory","type":"asset","subtype":"fixed","category":"Inventory","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1201","code":"1201","name":"Raw Materials Inventory","type":"asset","subtype":"fixed","category":"Inventory","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1202","code":"1202","name":"Work In Progress Inventory","type":"asset","subtype":"fixed","category":"Inventory","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1203","code":"1203","name":"Consumables & Supplies Inventory","type":"asset","subtype":"fixed","category":"Inventory","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1300","code":"1300","name":"Fixed Assets","type":"asset","subtype":"fixed","category":"Fixed Assets","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1310","code":"1310","name":"Accumulated Depreciation","type":"asset","subtype":"fixed","normal":"credit","contra":true,"category":"Fixed Assets","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2000","code":"2000","name":"Accounts Payable","type":"liability","subtype":"current","category":"Accounts Payable","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2100","code":"2100","name":"VAT Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2200","code":"2200","name":"WHT Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2210","code":"2210","name":"PAYE Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2220","code":"2220","name":"Pension Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2230","code":"2230","name":"NHF Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2240","code":"2240","name":"Other Payroll Deductions Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2250","code":"2250","name":"Current Tax Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2260","code":"2260","name":"Deferred Tax Liability","type":"liability","subtype":"noncurrent","category":"Deferred Tax Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1450","code":"1450","name":"Deferred Tax Asset","type":"asset","subtype":"fixed","category":"Other Assets","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1320","code":"1320","name":"Right-of-Use Assets","type":"asset","category":"Fixed Assets","subtype":"fixed","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1330","code":"1330","name":"Accumulated Depreciation - ROU Assets","type":"asset","contra":true,"normal":"credit","category":"Fixed Assets","subtype":"fixed","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2270","code":"2270","name":"Lease Liability","type":"liability","category":"Non-Current Liabilities","subtype":"noncurrent","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2300","code":"2300","name":"Stamp Duty Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"3000","code":"3000","name":"Owner''s Equity","type":"equity","category":"Owner''s Equity","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"4000","code":"4000","name":"Sales Revenue","type":"revenue","category":"Operating Income","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"4100","code":"4100","name":"Other Income","type":"revenue","category":"Operating Income","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"4200","code":"4200","name":"Sales Returns & Allowances","type":"revenue","contra":true,"normal":"debit","category":"Operating Income","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5000","code":"5000","name":"Cost of Goods Sold","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5010","code":"5010","name":"Direct Materials","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5020","code":"5020","name":"Direct Labour","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5030","code":"5030","name":"Carriage / Freight Inwards","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5040","code":"5040","name":"Production Overheads","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5100","code":"5100","name":"Rent Expense","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5200","code":"5200","name":"Utilities Expense","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5300","code":"5300","name":"Payroll Expense","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5400","code":"5400","name":"Office Supplies","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5500","code":"5500","name":"Marketing Expense","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5600","code":"5600","name":"Bank Fees","type":"expense","category":"Finance Costs","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5700","code":"5700","name":"Depreciation Expense","type":"expense","category":"Depreciation & Amortization","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5900","code":"5900","name":"Purchase Returns & Allowances","type":"expense","subtype":"cogs","contra":true,"normal":"credit","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5850","code":"5850","name":"Income Tax Expense","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5650","code":"5650","name":"Lease Interest Expense","type":"expense","category":"Finance Costs","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5750","code":"5750","name":"Inventory Write-down","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5950","code":"5950","name":"Bad Debt Expense (ECL)","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1150","code":"1150","name":"Allowance for Doubtful Accounts","type":"asset","contra":true,"normal":"credit","category":"Accounts Receivable","subtype":"current","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2290","code":"2290","name":"Deferred Revenue","type":"liability","category":"Current Liabilities","subtype":"current","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2296","code":"2296","name":"Goods Received Not Invoiced","type":"liability","category":"Current Liabilities","subtype":"current","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"3900","code":"3900","name":"Opening Balance Equity","type":"equity","category":"Opening Balance Equity","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"6010","code":"6010","name":"Foreign Exchange Gain / (Loss)","type":"revenue","subtype":"other","category":"Other Income","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5760","code":"5760","name":"Impairment Loss","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5770","code":"5770","name":"Provision Expense","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5780","code":"5780","name":"Revaluation Loss","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2295","code":"2295","name":"Provisions","type":"liability","category":"Non-Current Liabilities","subtype":"noncurrent","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"3200","code":"3200","name":"Revaluation Surplus","type":"equity","category":"Reserves","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"6000","code":"6000","name":"Realized Gains & Losses","type":"revenue","subtype":"other","category":"Other Income","status":"active"}],"banks":[],"transactions":[],"invoices":[],"bills":[],"expenses":[],"inventory":[],"inventoryLots":[],"fixedAssets":[],"payments":[],"taxGroups":[],"projects":[],"locations":[],"departments":[],"users":[],"messages":[],"budgets":{},"budgetAccounts":[],"favoriteReports":["pl","ar-aging","inventory-summary"],"bankFeed":[],"categoryRules":[],"bin":[],"employees":[],"payrollRuns":[],"recurringJournals":[],"reconciliations":[],"auditLog":[],"taxProvisions":[],"leases":[],"deferredRevenueSchedules":[],"provisions":[],"salesOrders":[],"salesReceipts":[],"creditNotes":[],"salesReturns":[],"purchaseOrders":[],"purchaseReceipts":[],"vendorCredits":[],"timesheets":[],"productionRecords":[],"nextProductionNum":1,"openingBalances":{"asOfDate":"2026-07-30","accountAmounts":{},"customerBalances":[],"vendorBalances":[],"posted":false,"postedDate":null},"nextInvoiceNum":1001,"nextBillNum":2001,"nextSalesOrderNum":1,"nextSalesReceiptNum":1,"nextCreditNoteNum":1,"nextSalesReturnNum":1,"nextPurchaseOrderNum":1,"nextPurchaseReceiptNum":1,"nextVendorCreditNum":1}
'::jsonb
    )
    returning id into new_company_id;

    insert into company_members (company_id, user_id, email, name, role)
      values (new_company_id, new.id, new.email, coalesce(new.raw_user_meta_data->>'name', ''), 'admin');
    insert into profiles (id, company_id) values (new.id, new_company_id);
  end if;

  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure handle_new_user();

-- ----------------------------------------------------------------------------
-- 7. Keep updated_at current on every save, useful later for conflict
--    detection or a "last saved at" indicator in the UI.
-- ----------------------------------------------------------------------------
create or replace function set_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists companies_set_updated_at on companies;
create trigger companies_set_updated_at
  before update on companies
  for each row execute procedure set_updated_at();

-- ----------------------------------------------------------------------------
-- 8. Migration backfill - REQUIRED if you already ran an earlier version of
--    this schema and have a real company in production.
--
--    The company_members table is brand new. Without this, an existing
--    company's original owner would have zero rows in it, and the new RLS
--    policies above would lock them out of their own data entirely - they
--    were only ever recognized by owner_id before, which no longer grants
--    access on its own.
--
--    This finds any company whose owner isn't yet in company_members and
--    adds them as Admin. If this is a brand new project with no existing
--    companies, this simply does nothing - safe to run either way, and
--    safe to run more than once.
-- ----------------------------------------------------------------------------
insert into company_members (company_id, user_id, email, name, role)
select c.id, c.owner_id, u.email, '', 'admin'
from companies c
join auth.users u on u.id = c.owner_id
where not exists (
  select 1 from company_members m where m.company_id = c.id and m.user_id = c.owner_id
);
