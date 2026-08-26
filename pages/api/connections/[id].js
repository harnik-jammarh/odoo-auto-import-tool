import supabaseAdmin, { verifyUser } from "../../../lib/supabaseAdmin";
import { encryptSecret } from "../../../lib/crypto";

export default async function handler(req, res) {
  const user = await verifyUser(req);
  if (!user) { res.status(401).json({ error: "Not signed in" }); return; }
  const { id } = req.query;

  // Ownership check — a user can only touch their own rows, belt-and-braces
  // alongside the database's own Row Level Security policy.
  const { data: existing, error: findErr } = await supabaseAdmin
    .from("saved_connections")
    .select("id, user_id")
    .eq("id", id)
    .single();
  if (findErr || !existing || existing.user_id !== user.id) {
    res.status(404).json({ error: "Connection not found" });
    return;
  }

  if (req.method === "PUT") {
    const { label, url, db, username, apiKey, autoCreateSafe } = req.body || {};
    const updates = { label, url, db, username, auto_create_safe: autoCreateSafe !== false };
    if (apiKey) updates.api_key_encrypted = encryptSecret(apiKey);
    const { error } = await supabaseAdmin.from("saved_connections").update(updates).eq("id", id);
    if (error) { res.status(500).json({ error: error.message }); return; }
    res.status(200).json({ ok: true });
    return;
  }

  if (req.method === "DELETE") {
    const { error } = await supabaseAdmin.from("saved_connections").delete().eq("id", id);
    if (error) { res.status(500).json({ error: error.message }); return; }
    res.status(200).json({ ok: true });
    return;
  }

  res.status(405).json({ error: "Method not allowed" });
}
