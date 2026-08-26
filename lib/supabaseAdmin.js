import { createClient } from "@supabase/supabase-js";

// SERVER-SIDE ONLY. This uses the secret service role key, which bypasses
// Row Level Security — never import this file from browser/client code.
// It is only used inside pages/api/* routes, and every route that uses it
// manually checks the caller's identity first (see verifyUser below).
const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY,
  { auth: { autoRefreshToken: false, persistSession: false } }
);

export default supabaseAdmin;

// Verifies the Supabase access token sent by the browser (Authorization:
// Bearer <token>) and returns the authenticated user, or null if invalid.
export async function verifyUser(req) {
  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!token) return null;
  const { data, error } = await supabaseAdmin.auth.getUser(token);
  if (error || !data?.user) return null;
  return data.user;
}
