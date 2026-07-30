// mono-exchange
//
// Called right after the Mono Connect widget succeeds on the client, with
// the one-time `code` it returned. This is the one step that genuinely
// cannot happen in the browser: exchanging that code for a permanent
// Account ID requires your Mono SECRET key, which must never reach client
// code. This function holds that secret (as a Supabase Edge Function
// secret, not a database value, not a .env file that ships to the
// browser) and is the only place it's ever used.
//
// Deploy with: supabase functions deploy mono-exchange
// Set the secret with: supabase secrets set MONO_SECRET_KEY=your_real_secret_key
import { createClient } from "npm:@supabase/supabase-js@2";
import { corsHeaders } from "../_shared/cors.ts";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const { code } = await req.json();
    if (!code) return json({ error: "Missing code" }, 400);

    const monoSecretKey = Deno.env.get("MONO_SECRET_KEY");
    if (!monoSecretKey) return json({ error: "MONO_SECRET_KEY is not configured on this project" }, 500);

    // Identify who's calling and which company they belong to, using their
    // own auth token - this reuses the exact same RLS the rest of the app
    // relies on, so this function can never be tricked into attaching a
    // bank account to a company the caller isn't actually a member of.
    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_ANON_KEY")!,
      { global: { headers: { Authorization: req.headers.get("Authorization")! } } }
    );
    const { data: { user }, error: userErr } = await supabase.auth.getUser();
    if (userErr || !user) return json({ error: "Not authenticated" }, 401);

    const { data: membership, error: memErr } = await supabase
      .from("company_members")
      .select("company_id, role")
      .eq("user_id", user.id)
      .limit(1)
      .maybeSingle();
    if (memErr || !membership) return json({ error: "No company found for this account" }, 404);
    if (membership.role === "viewer") return json({ error: "Your role doesn't allow connecting a bank account" }, 403);

    // Exchange the widget's one-time code for a permanent Account ID.
    // Documented at https://docs.mono.co/docs/financial-data/integration-guide
    const exchangeRes = await fetch("https://api.withmono.com/v2/accounts/auth", {
      method: "POST",
      headers: { "Content-Type": "application/json", "accept": "application/json", "mono-sec-key": monoSecretKey },
      body: JSON.stringify({ code }),
    });
    const exchangeBody = await exchangeRes.json();
    if (!exchangeRes.ok) return json({ error: exchangeBody?.message || "Mono rejected the exchange" }, 502);
    const accountId = exchangeBody?.data?.id || exchangeBody?.id;
    if (!accountId) return json({ error: "Mono didn't return an account id" }, 502);

    // Fetch account details for a friendlier display (bank name, masked
    // account number) - not required for the connection itself to work,
    // so a failure here doesn't block the connection from being saved.
    let institutionName = "Connected account", accountNumberMasked = "", currency = "NGN";
    try {
      const detailsRes = await fetch(`https://api.withmono.com/v2/accounts/${accountId}`, {
        headers: { accept: "application/json", "mono-sec-key": monoSecretKey },
      });
      const details = await detailsRes.json();
      const acct = details?.data?.account || details?.account;
      if (acct) {
        institutionName = acct.institution?.name || institutionName;
        currency = acct.currency || currency;
        if (acct.accountNumber) accountNumberMasked = "•••• " + String(acct.accountNumber).slice(-4);
      }
    } catch (_e) { /* non-fatal - connection is still saved below */ }

    const { data: connection, error: insertErr } = await supabase
      .from("bank_connections")
      .upsert(
        { company_id: membership.company_id, provider: "mono", external_account_id: accountId, institution_name: institutionName, account_number_masked: accountNumberMasked, currency, created_by: user.id, status: "active" },
        { onConflict: "company_id,external_account_id" }
      )
      .select()
      .single();
    if (insertErr) return json({ error: insertErr.message }, 500);

    return json({ connection });
  } catch (e) {
    return json({ error: String(e?.message || e) }, 500);
  }
});

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { ...corsHeaders, "Content-Type": "application/json" } });
}
