import supabaseAdmin, { verifyUser } from "../../../lib/supabaseAdmin";

// Only the account whose email matches ADMIN_EMAIL (set in your environment
// variables, never in client code) can list or decide on waitlist requests.
// Everyone else — including other signed-in users — gets a 403 here.
async function requireAdmin(req) {
  const user = await verifyUser(req);
  if (!user) return { error: 401, message: "Not signed in" };
  const adminEmail = (process.env.ADMIN_EMAIL || "").trim().toLowerCase();
  if (!adminEmail || (user.email || "").toLowerCase() !== adminEmail) {
    return { error: 403, message: "Not authorized" };
  }
  return { user };
}

export default async function handler(req, res) {
  const check = await requireAdmin(req);
  if (check.error) { res.status(check.error).json({ error: check.message }); return; }

  if (req.method === "GET") {
    const { data, error } = await supabaseAdmin
      .from("waitlist_requests")
      .select("*")
      .order("created_at", { ascending: false });
    if (error) { res.status(500).json({ error: error.message }); return; }
    res.status(200).json({ requests: data });
    return;
  }

  if (req.method === "POST") {
    const { id, action } = req.body || {}; // action: "approve" | "reject"
    if (!id || !["approve", "reject"].includes(action)) {
      res.status(400).json({ error: "id and a valid action are required" });
      return;
    }
    const { data: reqRow, error: fetchErr } = await supabaseAdmin
      .from("waitlist_requests")
      .select("*")
      .eq("id", id)
      .single();
    if (fetchErr || !reqRow) { res.status(404).json({ error: "Request not found" }); return; }

    if (action === "reject") {
      const { error } = await supabaseAdmin
        .from("waitlist_requests")
        .update({ status: "rejected", decided_at: new Date().toISOString() })
        .eq("id", id);
      if (error) { res.status(500).json({ error: error.message }); return; }
      res.status(200).json({ ok: true });
      return;
    }

    // Approve: this is the ONLY place a Supabase Auth user gets created for
    // a new signup. inviteUserByEmail creates the account and emails the
    // person a link to set their own password — we never see or store it.
    const { error: inviteErr } = await supabaseAdmin.auth.admin.inviteUserByEmail(reqRow.email);
    if (inviteErr && !/already registered|already exists/i.test(inviteErr.message || "")) {
      res.status(500).json({ error: inviteErr.message });
      return;
    }
    const { error } = await supabaseAdmin
      .from("waitlist_requests")
      .update({ status: "approved", decided_at: new Date().toISOString() })
      .eq("id", id);
    if (error) { res.status(500).json({ error: error.message }); return; }
    res.status(200).json({ ok: true });
    return;
  }

  res.status(405).json({ error: "Method not allowed" });
}
