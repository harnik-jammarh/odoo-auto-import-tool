import { useRouter } from "next/router";
import { supabase } from "../lib/supabaseClient";

// Shared top nav. Logo/title = Home button (always). When signed in, shows
// My Databases / My Profile / Sign Out tabs. When signed out, shows Sign In.
export default function Header({ session, active }) {
  const router = useRouter();

  async function signOut() {
    await supabase.auth.signOut();
    router.push("/login");
  }

  return (
    <div className="top-row" style={{ marginBottom: 20 }}>
      <div className="header-row" style={{ margin: 0, cursor: "pointer" }} onClick={() => router.push("/")}>
        <div className="logo">O</div>
        <h1>Odoo Auto-Import</h1>
      </div>
      <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
        {session === undefined ? null : session ? (
          <>
            <button className={`link-btn ${active === "databases" ? "nav-active" : ""}`} onClick={() => router.push("/dashboard")}>My Databases</button>
            <button className={`link-btn ${active === "profile" ? "nav-active" : ""}`} onClick={() => router.push("/profile")}>My Profile</button>
            <button className="link-btn" onClick={signOut}>Sign Out</button>
          </>
        ) : (
          <button className="btn btn-primary" onClick={() => router.push("/login")}>Sign In</button>
        )}
      </div>
    </div>
  );
}
