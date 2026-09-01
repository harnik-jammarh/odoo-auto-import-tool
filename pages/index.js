import { useState, useEffect } from "react";
import { useRouter } from "next/router";
import { supabase } from "../lib/supabaseClient";
import Header from "../components/Header";

const STEPS = [
  {
    title: "1. Request access",
    body: "Click \"Request Access\" and enter your email. Signups are approved manually — you'll get an email invite to set your password once approved.",
  },
  {
    title: "2. Sign in and add your Odoo database",
    body: "Once you're in, go to Databases -> + New Database. You'll need your Odoo URL, database name, your Odoo username, and an API key (Odoo -> your user icon -> My Profile -> Account Security -> New API Key).",
  },
  {
    title: "3. Upload a file",
    body: "Pick a saved database, then drag in an Excel or CSV file. Each sheet is analyzed automatically and matched to an Odoo module (Contacts, Inventory, Sales, Purchases, Accounting, Leads, Physical Inventory).",
  },
  {
    title: "4. Check the column mapping",
    body: "Review \"Maps to\" for each column — fix any wrong guesses using the dropdown, or tick \"Skip\" for columns you don't want imported.",
  },
  {
    title: "5. Choose what happens on duplicates",
    body: "Decide whether existing records should be updated, skipped, or handled \"smart\" (skip only if the record already has real transactions against it).",
  },
  {
    title: "6. Upload to Odoo",
    body: "Click Upload. You'll get a summary of created, updated, skipped and failed rows, plus notes on anything that needed a judgment call (like an auto-created location or category).",
  },
];

export default function Home() {
  const router = useRouter();
  const [session, setSession] = useState(undefined);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
  }, []);

  return (
    <div className="page">
      <Header session={session} active="home" />

      <h3>Getting started</h3>
      <div className="db-list">
        {STEPS.map((s) => (
          <div key={s.title} className="db-item" style={{ cursor: "default" }}>
            <div className="db-item-main" style={{ cursor: "default" }}>
              <div className="db-item-label">{s.title}</div>
              <div className="db-item-sub">{s.body}</div>
            </div>
          </div>
        ))}
      </div>

      <div className="note-box" style={{ marginTop: 20 }}>
        Modules supported today: Contacts, Inventory (including attributes/variants and vendor pricing), Accounting, Sales Orders, Purchase Orders, Leads, and Physical Inventory counts.
      </div>
    </div>
  );
}
