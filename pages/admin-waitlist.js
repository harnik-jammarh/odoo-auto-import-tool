import { useState, useEffect, useCallback } from "react";
import { useRouter } from "next/router";
import { supabase } from "../lib/supabaseClient";

export default function AdminWaitlist() {
  const router = useRouter();
  const [session, setSession] = useState(undefined);
  const [requests, setRequests] = useState([]);
  const [error, setError] = useState("");
  const [busyId, setBusyId] = useState(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (session === null) router.replace("/login");
  }, [session, router]);

  const authedFetch = useCallback(async (url, options = {}) => {
    const token = session?.access_token;
    const res = await fetch(url, {
      ...options,
      headers: { "Content-Type": "application/json", ...(options.headers || {}), Authorization: `Bearer ${token}` },
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Request failed");
    return data;
  }, [session]);

  const load = useCallback(async () => {
    setError("");
    try {
      const data = await authedFetch("/api/admin/waitlist");
      setRequests(data.requests);
    } catch (e) {
      setError(e.message); // e.g. "Not authorized" if this isn't the admin account
    }
  }, [authedFetch]);

  useEffect(() => {
    if (session) load();
  }, [session, load]);

  async function decide(id, action) {
    setBusyId(id);
    try {
      await authedFetch("/api/admin/waitlist", { method: "POST", body: JSON.stringify({ id, action }) });
      await load();
    } catch (e) {
      setError(e.message);
    } finally {
      setBusyId(null);
    }
  }

  if (session === undefined) return <div className="page">Loading...</div>;
  if (!session) return null;

  const pending = requests.filter((r) => r.status === "pending");
  const decided = requests.filter((r) => r.status !== "pending");

  return (
    <div className="page">
      <div className="header-row">
        <div className="logo">O</div>
        <h1>Odoo Auto-Import</h1>
      </div>
      <p className="subtitle">Waitlist Approvals</p>

      <div className="top-row">
        <div />
        <button className="link-btn" onClick={() => router.push("/dashboard")}>← Back to Dashboard</button>
      </div>

      {error && <div className="error-box">{error}</div>}

      <h3>Pending ({pending.length})</h3>
      {!pending.length && <p className="subtitle">No pending requests.</p>}
      <div className="db-list">
        {pending.map((r) => (
          <div key={r.id} className="db-item">
            <div className="db-item-main">
              <div className="db-item-label">{r.email}</div>
              <div className="db-item-sub">Requested {new Date(r.created_at).toLocaleString()}{r.note ? ` — "${r.note}"` : ""}</div>
            </div>
            <div className="db-item-actions">
              <button className="btn btn-primary" disabled={busyId === r.id} onClick={() => decide(r.id, "approve")}>
                {busyId === r.id ? "..." : "Approve"}
              </button>
              <button className="btn btn-danger" disabled={busyId === r.id} onClick={() => decide(r.id, "reject")}>
                Reject
              </button>
            </div>
          </div>
        ))}
      </div>

      {decided.length > 0 && (
        <>
          <h3 style={{ marginTop: 24 }}>Decided</h3>
          <div className="db-list">
            {decided.map((r) => (
              <div key={r.id} className="db-item">
                <div className="db-item-main">
                  <div className="db-item-label">{r.email}</div>
                  <div className="db-item-sub">
                    {r.status === "approved" ? "✅ Approved" : "❌ Rejected"} {r.decided_at ? `on ${new Date(r.decided_at).toLocaleString()}` : ""}
                  </div>
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}
