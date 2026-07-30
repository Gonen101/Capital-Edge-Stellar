// Shared CORS headers - Edge Functions are called directly from the
// browser (via supabase.functions.invoke), so they need to answer
// preflight OPTIONS requests and allow cross-origin calls from your app.
export const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};
