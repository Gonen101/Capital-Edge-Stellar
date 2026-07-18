-- ============================================================================
-- Capital Edge Stellar - Supabase schema
--
-- Run this once in your Supabase project's SQL Editor (Dashboard -> SQL
-- Editor -> New query -> paste this whole file -> Run).
--
-- Model: one row per company, holding its entire app state as JSONB (the
-- exact same shape the app already reads/writes as one JSON object today).
-- This is the fastest, lowest-risk path to a real backend, because almost
-- none of the app's existing business logic has to change - only how that
-- one JSON object is loaded and saved. A fully normalized schema (separate
-- tables per record type) is a better long-term architecture for complex
-- reporting and is worth doing later, but is a much bigger rewrite and
-- isn't needed to get a real, secure, multi-tenant app live.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. companies: one row per signed-up company, holding all of its data.
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
-- 2. profiles: links each auth user to exactly one company. Kept as its own
--    table (rather than just a column on auth.users, which you can't alter)
--    so a future "invite a teammate to my company" feature has somewhere to
--    extend into - for now, one signup = one owner = one company.
-- ----------------------------------------------------------------------------
create table if not exists profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  company_id uuid not null references companies(id) on delete cascade,
  created_at timestamptz not null default now()
);

-- ----------------------------------------------------------------------------
-- 3. Row Level Security - this is what actually prevents user A from ever
--    seeing user B's data, enforced by the database itself, not just by the
--    app's own code. Even a direct API call with a stolen session token
--    can't read another company's row.
-- ----------------------------------------------------------------------------
alter table companies enable row level security;
alter table profiles enable row level security;

-- A user can only see/edit the company they own.
create policy "Users can view their own company"
  on companies for select
  using (owner_id = auth.uid());

create policy "Users can update their own company"
  on companies for update
  using (owner_id = auth.uid());

-- Company rows are only ever created via the signup function below (see
-- section 4), never inserted directly by the client - this stops a
-- malicious client from creating extra companies or attaching to someone
-- else's. No insert policy is defined for regular users on purpose.

-- A user can only see their own profile row.
create policy "Users can view their own profile"
  on profiles for select
  using (id = auth.uid());

-- ----------------------------------------------------------------------------
-- 4. handle_new_user: runs automatically every time someone signs up via
--    Supabase Auth. Creates their company (starting EMPTY - no demo data)
--    and links their profile to it, atomically, so a user is never left in
--    a half-set-up state.
--
--    IMPORTANT: replace the placeholder JSONB below with the real output of
--    emptyCompanyData() from src/emptyCompanyData.js (run it once in Node
--    and paste the JSON - see DEPLOYMENT-GUIDE.md step 4 for the exact
--    command). Keeping it as a SQL default means new signups don't depend
--    on any client-side code to get set up correctly.
-- ----------------------------------------------------------------------------
create or replace function handle_new_user()
returns trigger
language plpgsql
security definer set search_path = public
as $$
declare
  new_company_id uuid;
begin
  insert into companies (owner_id, name, data)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'company_name', 'My Company'),
    '{"settings":{"companyName":"My Company","currencyCode":"NGN","mode":"light","accent":"#2563EB","reportOptions":{"showCodes":false,"hideZeroLines":true},"userName":"","accountingBasis":"accrual","payrollCountry":"NG","industry":"general","country":"NG","corporateTaxRate":30,"eclRates":{"Current":0.5,"1-30 days":1,"31-60 days":5,"61-90 days":15,"90+ days":40},"lockDate":null},"accounts":[{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1000","code":"1000","name":"Main Operating Account","type":"asset","subtype":"bank","category":"Bank","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1010","code":"1010","name":"Reserve Account","type":"asset","subtype":"bank","category":"Bank","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1050","code":"1050","name":"Petty Cash","type":"asset","subtype":"current","category":"Cash","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1100","code":"1100","name":"Accounts Receivable","type":"asset","subtype":"current","category":"Accounts Receivable","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1200","code":"1200","name":"Finished Goods Inventory","type":"asset","subtype":"fixed","category":"Inventory","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1201","code":"1201","name":"Raw Materials Inventory","type":"asset","subtype":"fixed","category":"Inventory","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1202","code":"1202","name":"Work In Progress Inventory","type":"asset","subtype":"fixed","category":"Inventory","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1203","code":"1203","name":"Consumables & Supplies Inventory","type":"asset","subtype":"fixed","category":"Inventory","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1300","code":"1300","name":"Fixed Assets","type":"asset","subtype":"fixed","category":"Fixed Assets","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1310","code":"1310","name":"Accumulated Depreciation","type":"asset","subtype":"fixed","normal":"credit","contra":true,"category":"Fixed Assets","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2000","code":"2000","name":"Accounts Payable","type":"liability","subtype":"current","category":"Accounts Payable","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2100","code":"2100","name":"VAT Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2200","code":"2200","name":"WHT Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2210","code":"2210","name":"PAYE Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2220","code":"2220","name":"Pension Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2230","code":"2230","name":"NHF Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2240","code":"2240","name":"Other Payroll Deductions Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2250","code":"2250","name":"Current Tax Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2260","code":"2260","name":"Deferred Tax Liability","type":"liability","subtype":"noncurrent","category":"Deferred Tax Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1450","code":"1450","name":"Deferred Tax Asset","type":"asset","subtype":"fixed","category":"Other Assets","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1320","code":"1320","name":"Right-of-Use Assets","type":"asset","category":"Fixed Assets","subtype":"fixed","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1330","code":"1330","name":"Accumulated Depreciation - ROU Assets","type":"asset","contra":true,"normal":"credit","category":"Fixed Assets","subtype":"fixed","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2270","code":"2270","name":"Lease Liability","type":"liability","category":"Non-Current Liabilities","subtype":"noncurrent","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2300","code":"2300","name":"Stamp Duty Payable","type":"liability","subtype":"current","category":"Current Liabilities","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"3000","code":"3000","name":"Owner''s Equity","type":"equity","category":"Owner''s Equity","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"4000","code":"4000","name":"Sales Revenue","type":"revenue","category":"Operating Income","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"4100","code":"4100","name":"Other Income","type":"revenue","category":"Operating Income","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"4200","code":"4200","name":"Sales Returns & Allowances","type":"revenue","contra":true,"normal":"debit","category":"Operating Income","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5000","code":"5000","name":"Cost of Goods Sold","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5010","code":"5010","name":"Direct Materials","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5020","code":"5020","name":"Direct Labour","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5030","code":"5030","name":"Carriage / Freight Inwards","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5040","code":"5040","name":"Production Overheads","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5100","code":"5100","name":"Rent Expense","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5200","code":"5200","name":"Utilities Expense","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5300","code":"5300","name":"Payroll Expense","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5400","code":"5400","name":"Office Supplies","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5500","code":"5500","name":"Marketing Expense","type":"expense","category":"Operating Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5600","code":"5600","name":"Bank Fees","type":"expense","category":"Finance Costs","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5700","code":"5700","name":"Depreciation Expense","type":"expense","category":"Depreciation & Amortization","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5900","code":"5900","name":"Purchase Returns & Allowances","type":"expense","subtype":"cogs","contra":true,"normal":"credit","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5850","code":"5850","name":"Income Tax Expense","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5650","code":"5650","name":"Lease Interest Expense","type":"expense","category":"Finance Costs","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5750","code":"5750","name":"Inventory Write-down","type":"expense","subtype":"cogs","category":"Cost of Goods Sold (COGS)","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5950","code":"5950","name":"Bad Debt Expense (ECL)","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"1150","code":"1150","name":"Allowance for Doubtful Accounts","type":"asset","contra":true,"normal":"credit","category":"Accounts Receivable","subtype":"current","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2290","code":"2290","name":"Deferred Revenue","type":"liability","category":"Current Liabilities","subtype":"current","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2296","code":"2296","name":"Goods Received Not Invoiced","type":"liability","category":"Current Liabilities","subtype":"current","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"3900","code":"3900","name":"Opening Balance Equity","type":"equity","category":"Opening Balance Equity","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"6010","code":"6010","name":"Foreign Exchange Gain / (Loss)","type":"revenue","subtype":"other","category":"Other Income","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5760","code":"5760","name":"Impairment Loss","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5770","code":"5770","name":"Provision Expense","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"5780","code":"5780","name":"Revaluation Loss","type":"expense","category":"Other Expenses","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"2295","code":"2295","name":"Provisions","type":"liability","category":"Non-Current Liabilities","subtype":"noncurrent","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"3200","code":"3200","name":"Revaluation Surplus","type":"equity","category":"Reserves","status":"active"},{"parentId":null,"description":"","taxAccountId":"","currency":"","id":"6000","code":"6000","name":"Realized Gains & Losses","type":"revenue","subtype":"other","category":"Other Income","status":"active"}],"banks":[],"transactions":[],"invoices":[],"bills":[],"expenses":[],"inventory":[],"inventoryLots":[],"fixedAssets":[],"payments":[],"taxGroups":[],"projects":[],"locations":[],"departments":[],"budgets":{},"budgetAccounts":[],"favoriteReports":["pl","ar-aging","inventory-summary"],"bankFeed":[],"categoryRules":[],"bin":[],"employees":[],"payrollRuns":[],"recurringJournals":[],"reconciliations":[],"auditLog":[],"taxProvisions":[],"leases":[],"deferredRevenueSchedules":[],"provisions":[],"salesOrders":[],"salesReceipts":[],"creditNotes":[],"salesReturns":[],"purchaseOrders":[],"purchaseReceipts":[],"vendorCredits":[],"timesheets":[],"productionRecords":[],"nextProductionNum":1,"openingBalances":{"asOfDate":"2026-07-18","accountAmounts":{},"customerBalances":[],"vendorBalances":[],"posted":false,"postedDate":null},"nextInvoiceNum":1001,"nextBillNum":2001,"nextSalesOrderNum":1,"nextSalesReceiptNum":1,"nextCreditNoteNum":1,"nextSalesReturnNum":1,"nextPurchaseOrderNum":1,"nextPurchaseReceiptNum":1,"nextVendorCreditNum":1}
'::jsonb
  )
  returning id into new_company_id;

  insert into profiles (id, company_id) values (new.id, new_company_id);

  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure handle_new_user();

-- ----------------------------------------------------------------------------
-- 5. Keep updated_at current on every save, useful later for conflict
--    detection or an "last saved at" indicator in the UI.
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
