import { useState, useEffect } from "react";
import { useRouter } from "next/router";
import { supabase } from "../lib/supabaseClient";

export default function Login() {
  const router = useRouter();
  const [mode, setMode] = useState("signin"); // "signin" | "signup"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [note, setNote] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      if (data.session) router.replace("/dashboard");
    });
  }, [router]);

  async function handleSubmit(e) {
    e.preventDefault();
    setError(""); setInfo(""); setBusy(true);
    try {
      if (mode === "signup") {
        // Signup never talks to Supabase Auth directly — it only files a
        // waitlist request. The account itself is only created once the
        // admin approves it (see /api/admin/waitlist), at which point an
        // invite email goes out to set a password.
        const res = await fetch("/api/waitlist", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ email, note }),
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(data.error || "Request failed");
        setInfo(data.message || "Request received — you'll get an email once you're approved.");
        setMode("signin");
      } else {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        router.replace("/dashboard");
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="page">
      <div className="auth-wrap">
        <div className="header-row">
          <div className="logo">O</div>
          <h1>Odoo Auto-Import</h1>
        </div>
        <p className="subtitle">Sign in to access your saved databases.</p>

        <div className="auth-tabs">
          <div className={`auth-tab ${mode === "signin" ? "active" : ""}`} onClick={() => setMode("signin")}>Sign In</div>
          <div className={`auth-tab ${mode === "signup" ? "active" : ""}`} onClick={() => setMode("signup")}>Request Access</div>
        </div>

        <form className="card" onSubmit={handleSubmit}>
          <div className="field">
            <label>Email</label>
            <input type="email" required value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@company.com" />
          </div>
          {mode === "signin" ? (
            <div className="field">
              <label>Password</label>
              <input type="password" required minLength={6} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="At least 6 characters" />
            </div>
          ) : (
            <div className="field">
              <label>What do you want to use this for? (optional)</label>
              <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="e.g. importing sale orders from our ERP" />
            </div>
          )}
          <button className="btn btn-primary" type="submit" disabled={busy} style={{ width: "100%" }}>
            {busy ? "Please wait..." : mode === "signup" ? "Request Access" : "Sign In"}
          </button>
          {error && <div className="error-text">{error}</div>}
          {info && <div className="note-box" style={{ marginTop: 10 }}>{info}</div>}
          {mode === "signup" && (
            <p className="subtitle" style={{ marginTop: 10 }}>
              New signups are approved manually. You'll get an email invite to set your password once you're approved.
            </p>
          )}
        </form>
      </div>
    </div>
  );
}

