-- ============================================================================
-- Capital Edge Stellar - live bank feeds (Mono)
--
-- Run this AFTER schema.sql, once you're ready to set up live bank feeds.
-- This is a genuinely optional layer - the app works completely fine
-- without it, using CSV/Excel import instead. Run this only when you've
-- signed up for a Mono account and are ready to configure it (see
-- DEPLOYMENT-GUIDE.md for the full walkthrough, including the Edge
-- Functions this depends on).
-- ============================================================================

create table if not exists bank_connections (
  id uuid primary key default gen_random_uuid(),
  company_id uuid not null references companies(id) on delete cascade,
  provider text not null default 'mono',
  external_account_id text not null,
  institution_name text not null default 'Connected account',
  account_number_masked text not null default '',
  currency text not null default 'NGN',
  status text not null default 'active' check (status in ('active', 'disconnected')),
  needs_sync boolean not null default false,
  last_synced_at timestamptz,
  last_webhook_at timestamptz,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  unique (company_id, external_account_id)
);

create index if not exists bank_connections_company_id_idx on bank_connections(company_id);

alter table bank_connections enable row level security;

-- Any non-viewer member of the company can see and manage its bank
-- connections - the same boundary as everywhere else in the app: viewers
-- can look, everyone else with real access can act.
create policy "Non-viewer members can view their company's bank connections"
  on bank_connections for select
  using (exists (
    select 1 from company_members
    where company_id = bank_connections.company_id and user_id = auth.uid()
  ));

create policy "Non-viewer members can manage their company's bank connections"
  on bank_connections for all
  using (exists (
    select 1 from company_members
    where company_id = bank_connections.company_id and user_id = auth.uid() and role <> 'viewer'
  ))
  with check (exists (
    select 1 from company_members
    where company_id = bank_connections.company_id and user_id = auth.uid() and role <> 'viewer'
  ));
