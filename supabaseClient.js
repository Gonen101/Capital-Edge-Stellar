// ============================================================================
// supabaseClient.js
//
// One shared Supabase client for the whole app. Reads its connection details
// from environment variables so the same code works locally and on Vercel
// without editing anything - see DEPLOYMENT-GUIDE.md for where these values
// come from and how to set them.
//
// The "anon" key below is meant to be public (it ships in your frontend
// bundle - anyone can see it in the browser). It is NOT a secret. What
// actually protects your data is the Row Level Security policies in
// supabase/schema.sql, which run inside the database itself regardless of
// what key or code is used to call it.
// ============================================================================
import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  throw new Error(
    'Missing Supabase environment variables. Set VITE_SUPABASE_URL and ' +
    'VITE_SUPABASE_ANON_KEY in a .env.local file (for local dev) or in your ' +
    'Vercel project settings (for the deployed site). See DEPLOYMENT-GUIDE.md.'
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
