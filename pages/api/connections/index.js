import supabaseAdmin, { verifyUser } from "../../../lib/supabaseAdmin";
import { encryptSecret, decryptSecret } from "../../../lib/crypto";

// GET  -> list the logged-in user's saved connections (API key decrypted for use)
// POST -> save a new connection (API key encrypted before it touches the database)
export default async function handler(req, res) {
  const user = await verifyUser(req);
  if (!user) {
    res.status(401).json({ error: "Not signed in" });
    return;
  }

  if (req.method === "GET") {
    const { data, error } = await supabaseAdmin
      .from("saved_connections")
      .select("*")
      .eq("user_id", user.id)
      .order("created_at", { ascending: true });
    if (error) { res.status(500).json({ error: error.message }); return; }

    const connections = data.map((row) => ({
      id: row.id,
      label: row.label,
      url: row.url,
      db: row.db,
      username: row.username,
      autoCreateSafe: row.auto_create_safe,
      apiKey: safeDecrypt(row.api_key_encrypted),
    }));
    res.status(200).json({ connections });
    return;
  }

  if (req.method === "POST") {
    const { label, url, db, username, apiKey, autoCreateSafe } = req.body || {};
    if (!label || !url || !db || !username || !apiKey) {
      res.status(400).json({ error: "label, url, db, username and apiKey are all required" });
      return;
    }
    const { data, error } = await supabaseAdmin
      .from("saved_connections")
      .insert({
        user_id: user.id,
        label,
        url,
        db,
        username,
        api_key_encrypted: encryptSecret(apiKey),
        auto_create_safe: autoCreateSafe !== false,
      })
      .select()
      .single();
    if (error) { res.status(500).json({ error: error.message }); return; }
    res.status(200).json({ id: data.id });
    return;
  }

  res.status(405).json({ error: "Method not allowed" });
}

function safeDecrypt(value) {
  try { return decryptSecret(value); } catch (e) { return ""; }
}
