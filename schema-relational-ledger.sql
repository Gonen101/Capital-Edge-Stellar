-- ============================================================================
-- Capital Edge Stellar - relational ledger (accounts + transactions)
--
-- Run this AFTER schema.sql (and schema-bank-feeds.sql if you're using
-- that). This is the first, highest-value piece of moving off the
-- single-JSON-blob storage model: the chart of accounts and every journal
-- entry get real tables, with real foreign keys and real indexes - the
-- data everything else in the app (every report, every balance) is
-- actually derived from.
--
-- WHY THIS SCOPE, NOT EVERYTHING AT ONCE: normalizing all ~30 collections
-- (invoices, bills, inventory, payroll, and the rest) in a single pass
-- would be a multi-week project even done carefully, and rushing it risks
-- corrupting real financial data. Accounts and transactions are
-- deliberately the highest-value target - the most-queried data, and the
-- foundation everything else is computed from. The remaining collections
-- stay in the JSON blob for now, following the exact same pattern this one
-- establishes whenever they're tackled next.
--
-- WHAT DOESN'T CHANGE: nothing about the app's own business logic. The
-- data-access layer in AuthGate.jsx reconstructs the exact same
-- data.accounts / data.transactions shape the app has always worked with,
-- and translates changes back into targeted table writes - App.jsx itself
-- is untouched. All the risk from this migration is contained to that one
-- new, independently testable layer.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. gl_accounts - the chart of accounts, one row per account. id is the
--    app's own account id (e.g. "1000", "4000") - only unique within a
--    company, not globally, so the primary key is the pair.
-- ----------------------------------------------------------------------------
create table if not exists gl_accounts (
  id text not null,
  company_id uuid not null references companies(id) on delete cascade,
  code text not null,
  name text not null,
  type text not null check (type in ('asset', 'liability', 'equity', 'revenue', 'expense')),
  subtype text,
  category text,
  status text not null default 'active',
  parent_id text,
  description text not null default '',
  tax_account_id text not null default '',
  currency text not null default '',
  normal text check (normal in ('debit', 'credit') or normal is null),
  contra boolean not null default false,
  primary key (company_id, id)
);

create index if not exists gl_accounts_company_id_idx on gl_accounts(company_id);

-- ----------------------------------------------------------------------------
-- 2. gl_transactions - one row per journal entry (the header). id is the
--    app's own uid("txn") value.
-- ----------------------------------------------------------------------------
create table if not exists gl_transactions (
  id text not null,
  company_id uuid not null references companies(id) on delete cascade,
  date date not null,
  memo text not null default '',
  source text not null default 'manual',
  doc_id text,
  created_at timestamptz not null default now(),
  primary key (company_id, id)
);

create index if not exists gl_transactions_company_date_idx on gl_transactions(company_id, date);
create index if not exists gl_transactions_doc_id_idx on gl_transactions(company_id, doc_id) where doc_id is not null;

-- ----------------------------------------------------------------------------
-- 3. gl_transaction_lines - the actual debit/credit lines. line_order
--    preserves the exact sequence the app displays lines in, since a
--    relational table's row order isn't otherwise guaranteed to match the
--    array order the app originally saved.
-- ----------------------------------------------------------------------------
create table if not exists gl_transaction_lines (
  id uuid primary key default gen_random_uuid(),
  company_id uuid not null references companies(id) on delete cascade,
  transaction_id text not null,
  account_id text not null,
  debit numeric not null default 0,
  credit numeric not null default 0,
  line_order int not null default 0,
  foreign key (company_id, transaction_id) references gl_transactions(company_id, id) on delete cascade
);

create index if not exists gl_transaction_lines_txn_idx on gl_transaction_lines(company_id, transaction_id);
create index if not exists gl_transaction_lines_account_idx on gl_transaction_lines(company_id, account_id);

-- ----------------------------------------------------------------------------
-- 4. Row Level Security - the same boundary as the rest of the app: any
--    member can read, only non-viewer members can write. This is now
--    enforced per-table instead of on one shared JSON blob.
-- ----------------------------------------------------------------------------
alter table gl_accounts enable row level security;
alter table gl_transactions enable row level security;
alter table gl_transaction_lines enable row level security;

create policy "Members can view their company's accounts"
  on gl_accounts for select
  using (exists (select 1 from company_members where company_id = gl_accounts.company_id and user_id = auth.uid()));
create policy "Non-viewer members can manage their company's accounts"
  on gl_accounts for all
  using (exists (select 1 from company_members where company_id = gl_accounts.company_id and user_id = auth.uid() and role <> 'viewer'))
  with check (exists (select 1 from company_members where company_id = gl_accounts.company_id and user_id = auth.uid() and role <> 'viewer'));

create policy "Members can view their company's transactions"
  on gl_transactions for select
  using (exists (select 1 from company_members where company_id = gl_transactions.company_id and user_id = auth.uid()));
create policy "Non-viewer members can manage their company's transactions"
  on gl_transactions for all
  using (exists (select 1 from company_members where company_id = gl_transactions.company_id and user_id = auth.uid() and role <> 'viewer'))
  with check (exists (select 1 from company_members where company_id = gl_transactions.company_id and user_id = auth.uid() and role <> 'viewer'));

create policy "Members can view their company's transaction lines"
  on gl_transaction_lines for select
  using (exists (select 1 from company_members where company_id = gl_transaction_lines.company_id and user_id = auth.uid()));
create policy "Non-viewer members can manage their company's transaction lines"
  on gl_transaction_lines for all
  using (exists (select 1 from company_members where company_id = gl_transaction_lines.company_id and user_id = auth.uid() and role <> 'viewer'))
  with check (exists (select 1 from company_members where company_id = gl_transaction_lines.company_id and user_id = auth.uid() and role <> 'viewer'));

-- ----------------------------------------------------------------------------
-- 5. Backfill - copies each existing company's accounts and transactions
--    (currently living inside companies.data as JSON) into the new tables.
--
--    Deliberately non-destructive: this only READS from companies.data, it
--    never modifies or clears it. If anything about this migration looks
--    wrong afterward, your original data is still sitting there untouched -
--    nothing has been deleted. Safe to re-run; already-migrated rows are
--    just re-upserted to the same values.
-- ----------------------------------------------------------------------------
insert into gl_accounts (id, company_id, code, name, type, subtype, category, status, parent_id, description, tax_account_id, currency, normal, contra)
select
  a->>'id', c.id, a->>'code', a->>'name', a->>'type', a->>'subtype', a->>'category',
  coalesce(a->>'status', 'active'), a->>'parentId', coalesce(a->>'description', ''),
  coalesce(a->>'taxAccountId', ''), coalesce(a->>'currency', ''), a->>'normal',
  coalesce((a->>'contra')::boolean, false)
from companies c, jsonb_array_elements(c.data->'accounts') a
on conflict (company_id, id) do update set
  code = excluded.code, name = excluded.name, type = excluded.type, subtype = excluded.subtype,
  category = excluded.category, status = excluded.status, parent_id = excluded.parent_id,
  description = excluded.description, tax_account_id = excluded.tax_account_id,
  currency = excluded.currency, normal = excluded.normal, contra = excluded.contra;

insert into gl_transactions (id, company_id, date, memo, source, doc_id)
select t->>'id', c.id, (t->>'date')::date, coalesce(t->>'memo', ''), coalesce(t->>'source', 'manual'), t->>'docId'
from companies c, jsonb_array_elements(c.data->'transactions') t
on conflict (company_id, id) do update set
  date = excluded.date, memo = excluded.memo, source = excluded.source, doc_id = excluded.doc_id;

insert into gl_transaction_lines (company_id, transaction_id, account_id, debit, credit, line_order)
select c.id, t->>'id', l->>'accountId', coalesce((l->>'debit')::numeric, 0), coalesce((l->>'credit')::numeric, 0), ord - 1
from companies c,
     jsonb_array_elements(c.data->'transactions') t,
     jsonb_array_elements(t->'lines') with ordinality as line_set(l, ord)
where not exists (
  -- Skip transactions already backfilled with their lines, so re-running
  -- this doesn't duplicate lines for transactions that haven't changed.
  select 1 from gl_transaction_lines existing
  where existing.company_id = c.id and existing.transaction_id = t->>'id'
);
