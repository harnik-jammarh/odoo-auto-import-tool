import { useState, useEffect } from "react";
import { useRouter } from "next/router";
import { supabase } from "../lib/supabaseClient";
import Header from "../components/Header";

export default function Profile() {
  const router = useRouter();
  const [session, setSession] = useState(undefined); // undefined = loading, null = signed out
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (session === null) router.replace("/login");
  }, [session, router]);

  async function handleChangePassword(e) {
    e.preventDefault();
    setError(""); setInfo("");
    if (newPassword.length < 6) { setError("New password must be at least 6 characters."); return; }
    if (newPassword !== confirmPassword) { setError("Passwords don't match."); return; }
    setBusy(true);
    try {
      const { error } = await supabase.auth.updateUser({ password: newPassword });
      if (error) throw error;
      setInfo("Password updated.");
      setNewPassword(""); setConfirmPassword("");
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  }

  if (session === undefined) return <div className="page">Loading...</div>;
  if (!session) return null;

  return (
    <div className="page">
      <Header session={session} active="profile" />
      <p className="subtitle">My Profile</p>

      <div className="card">
        <div className="field">
          <label>Email</label>
          <input type="email" value={session.user.email} disabled />
        </div>
      </div>

      <div className="card" style={{ marginTop: 16 }}>
        <h3 style={{ marginTop: 0 }}>Change Password</h3>
        <form onSubmit={handleChangePassword}>
          <div className="field">
            <label>New Password</label>
            <input
              type="password"
              required
              minLength={6}
              value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)}
              placeholder="At least 6 characters"
            />
          </div>
          <div className="field">
            <label>Confirm New Password</label>
            <input
              type="password"
              required
              minLength={6}
              value={confirmPassword}
              onChange={(e) => setConfirmPassword(e.target.value)}
              placeholder="Retype new password"
            />
          </div>
          <button className="btn btn-primary" type="submit" disabled={busy}>
            {busy ? "Updating..." : "Update Password"}
          </button>
          {error && <div className="error-text">{error}</div>}
          {info && <div className="note-box" style={{ marginTop: 10 }}>{info}</div>}
        </form>
      </div>
    </div>
  );
}
