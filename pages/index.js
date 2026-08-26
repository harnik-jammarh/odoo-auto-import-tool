import supabaseAdmin from "../../../lib/supabaseAdmin";

// POST -> anyone can submit an email to request access. This does NOT touch
// Supabase Auth at all — it only inserts a "pending" row here. No Auth user
// (and no login ability) exists until the admin approves the request from
// /admin/waitlist, which is the only thing that actually calls
// supabaseAdmin.auth.admin.inviteUserByEmail().
export default async function handler(req, res) {
  if (req.method !== "POST") { res.status(405).json({ error: "Method not allowed" }); return; }

  const { email, note } = req.body || {};
  if (!email || typeof email !== "string" || !/^\S+@\S+\.\S+$/.test(email)) {
    res.status(400).json({ error: "A valid email is required" });
    return;
  }
  const cleanEmail = email.trim().toLowerCase();

  const { data: existing } = await supabaseAdmin
    .from("waitlist_requests")
    .select("id, status")
    .eq("email", cleanEmail)
    .maybeSingle();

  if (existing) {
    if (existing.status === "pending") { res.status(200).json({ ok: true, message: "You're already on the waitlist — hang tight." }); return; }
    if (existing.status === "approved") { res.status(200).json({ ok: true, message: "You're already approved — check your email for the invite, or just sign in." }); return; }
    // previously rejected: allow a fresh request by resetting it to pending
    const { error } = await supabaseAdmin
      .from("waitlist_requests")
      .update({ status: "pending", note: note || null, decided_at: null })
      .eq("id", existing.id);
    if (error) { res.status(500).json({ error: error.message }); return; }
    res.status(200).json({ ok: true, message: "Request received — you'll get an email once you're approved." });
    return;
  }

  const { error } = await supabaseAdmin
    .from("waitlist_requests")
    .insert({ email: cleanEmail, note: note || null, status: "pending" });
  if (error) { res.status(500).json({ error: error.message }); return; }

  res.status(200).json({ ok: true, message: "Request received — you'll get an email once you're approved." });
}
