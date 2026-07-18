# Deploying CapitalEdge Stellar as a real, multi-tenant website

This turns the single-file artifact into a real hosted app where:
- Anyone can sign up and gets their own brand-new, **completely empty** company
- No one can ever see another user's data (enforced by the database itself)
- The first thing anyone sees is the real landing page - animated hero, brand mark, login/signup - not a bare form

## This is a refreshed build, not the original

The app has grown substantially since this kit was first built - Inventory types with their own ledger accounts, Production/usage tracking, Budgets with a lock-until-edit flow, a Subscription panel, Time Sheets, Transaction Locking, Bulk Update, Locations and Departments, Opening Balances, and dozens of new reports. Everything in this folder was **regenerated from the current app**, not patched forward from the old version:

- `src/App.jsx` - the current app, patched the same three small ways as before (load/save through Supabase instead of browser storage, plus a real sign-out button and a Factory Reset that clears to an empty company instead of ever restoring demo data). Along the way, a real gap was caught and fixed in the source app itself: "Start Fresh" hadn't been updated to clear Sales Orders, Purchase Orders, Time Sheets, Production Records, and several other collections added since it was first written - it does now, in both this hosted copy and the original artifact.
- `src/emptyCompanyData.js` - completely rebuilt against the current chart of accounts (56 accounts now, including the inventory-type accounts and Opening Balance Equity). **Verified two ways**: every account checked field-by-field against the app's own output, and a full comparison confirming every top-level key in a real company's data has a matching key here, with nothing missing and nothing extra.
- `src/AuthGate.jsx` - the landing page and login/signup, described below.
- `src/supabaseClient.js` - connects to your Supabase project (unchanged).
- `supabase/schema.sql` - the database schema, Row Level Security, and the signup trigger, **re-embedded with the newly-verified empty-company data** - it works as soon as you run it, no editing required.

You do not need to touch any of the business logic. Everything below is account setup, configuration, and deployment - no accounting code to write.

## The landing page

`AuthGate.jsx` now renders the full animated landing page you approved as a standalone preview - the same hero on the left (scattered figures rising into a growth chart, a small figure climbing it book in hand, arriving at a trophy), the CapitalEdge Stellar brand mark, and the same scene running blurred behind the sign-in form on the right. It's the same code, ported directly into a React component rather than rebuilt - what you saw is what ships. The only change: the decorative sign-in box from the preview is replaced with the real, working Supabase auth logic (email/password, signup creates a company via the trigger below, error and loading states).

The wordmark uses Playfair Display, loaded in `index.html`.

---

## 1. Create your Supabase project

1. Go to [supabase.com](https://supabase.com) and create a free account.
2. Click **New Project**. Pick a name, a database password (save it somewhere), and a region close to your users.
3. Wait ~2 minutes for it to provision.

## 2. Run the database schema

1. In your Supabase project, open **SQL Editor** (left sidebar) → **New query**.
2. Open `supabase/schema.sql` from this folder, copy the whole file, paste it in, and click **Run**.
3. You should see "Success. No rows returned." This created two tables (`companies`, `profiles`), turned on Row Level Security, and set up the trigger that gives every new signup their own empty company automatically.

## 3. Get your API keys

1. In Supabase: **Project Settings** (gear icon) → **API**.
2. Copy the **Project URL** and the **anon / public** key. (Not the `service_role` key - that one is a secret and should never go in frontend code.)
3. Just keep these two values somewhere handy for a few minutes - a sticky note, a scratch text file, anything. They go into Vercel's dashboard in step 6, not into any file in this project.

## 4. Configure email confirmation (recommended for a real launch)

By default, Supabase requires users to click a confirmation link before they can log in. For testing, you can turn this off to move faster:

**Authentication** → **Providers** → **Email** → turn off "Confirm email".

Turn it back on before you have real users, or configure a custom SMTP provider (Supabase's docs cover this) so confirmation emails don't land in spam.

## 5. Push to GitHub

You don't need to run anything on your own computer for this - just get the folder onto GitHub.

```bash
git init
git add .
git commit -m "Initial commit"
```

Create a new repository on GitHub, then follow its instructions to push this folder there. (GitHub Desktop, if you prefer a visual tool over the terminal, works just as well for this step.)

## 6. Deploy to Vercel - this is where your two API keys actually go

This is the easiest way to get the two Supabase values into the app - no hidden files, no terminal, just a normal web form.

1. Go to [vercel.com](https://vercel.com), sign in with GitHub, click **Add New → Project**, and pick your repository.
2. Vercel auto-detects Vite - leave the build settings as default.
3. Before clicking Deploy, look for **Environment Variables** on that same page. Add two:
   - Name: `VITE_SUPABASE_URL` — Value: your Project URL from Supabase
   - Name: `VITE_SUPABASE_ANON_KEY` — Value: your anon/public key from Supabase
4. Click **Deploy**.

A few minutes later you have a real URL (`your-project.vercel.app`) with the animated landing page live on it. Every push to your GitHub repo redeploys automatically.

## 7. Add a custom domain (optional)

**Vercel Project → Settings → Domains** → add your domain and follow the DNS instructions it gives you. Takes effect within a few minutes to a few hours depending on your DNS provider.

## Optional: running it on your own computer first

You don't need this to go live - it's only useful if you want to preview changes before pushing them. If you want to try it:

```bash
npm install
npm run dev
```

This does need a `.env.local` file in the project root with the same two values as above. If creating a file that starts with a dot is giving you trouble (many file managers hide dotfiles by default), skip this step entirely and just use Vercel - it doesn't need this file at all.

---

## What "no user can see another user's data" actually rests on

Not just app code - the database itself. In `schema.sql`, every table has Row Level Security turned on with policies like:

```sql
create policy "Users can view their own company"
  on companies for select
  using (owner_id = auth.uid());
```

This means even a direct API call using a stolen session token, or a bug in the frontend code, cannot return another company's row - Postgres itself refuses the query.

## What happens on signup, exactly

1. User fills in the signup form on the landing page → `supabase.auth.signUp()` is called.
2. Supabase creates the auth user.
3. This fires the `handle_new_user()` trigger (defined in `schema.sql`), which:
   - Creates one row in `companies` with the embedded empty-company JSON as its starting `data`
   - Creates one row in `profiles` linking that user to that company
4. The user logs in, `AuthGate.jsx` fetches their (and only their) company row, and renders the full app with it.

Nothing in this path can produce demo data - verified directly: the empty-company JSON was loaded through the app's real `migrateData()` function and confirmed to produce a Balance Sheet of exactly 0 = 0 and a trivially reconciling Cash Flow Statement, with zero invoices, bills, or banks.

## If you make more changes to the app before going further

The same rule as last time: this is a snapshot. If you keep building on the artifact after this, `App.jsx` here won't have those changes until it's re-patched. When you're ready for another release, say so and the same three small edits (plus a fresh field-by-field verification of `emptyCompanyData.js` against whatever's changed in the chart of accounts) get reapplied to the latest version - not a rebuild from scratch.

## A note on the storage model, honestly

This setup stores each company's entire dataset as one JSON object in a single database column - the same shape the app already used with browser storage, just moved to Postgres. This was the right call to get you live **quickly and with near-zero risk to the tested business logic**. It has a real limitation worth knowing about before you scale: you can't run direct SQL reports across companies, and two people editing the *same* company's books at the exact same moment could overwrite each other's changes (last write wins).

If usage grows past solo bookkeeping - a small team working the same books simultaneously, or cross-company analytics - the next real step is splitting this into proper relational tables (one row per transaction, per invoice, etc.), with Row Level Security scoped by `company_id` on each. That's a genuine, separate project. Nothing about today's setup blocks that migration later.

## Known limitations carried over from testing this session

- If you view the Cash Flow Statement for a period that includes the exact date you post an Opening Balance entry, it may show a small reconciliation difference - this is disclosed in the app's own Opening Balances screen. For any period after that date (normal day-to-day use), it reconciles exactly.
- Transaction Locking is wired into the highest-traffic edit/delete points (Journal, Invoices, Bills, Expenses, Inventory adjustments, Production records) but not yet into every single document type in the app (fixed asset disposals and a few older sales/purchase documents aren't gated). Extending it fully is a real, doable follow-up rather than something silently assumed to be complete.
