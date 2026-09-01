import { useState } from "react";
import { EXPIRY_OPTIONS, computeExpiryFromOption } from "../lib/apiExpiry";

function toDatetimeLocal(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

export default function ConnectionForm({ initial, onSave, onSaved, onCancel }) {
  const [form, setForm] = useState(initial || { url: "", db: "", username: "", apiKey: "", driveApiKey: "", autoCreateSafe: true, apiKeyExpiresAt: null });
  const [expiryOption, setExpiryOption] = useState(initial?.apiKeyExpiresAt ? "Custom Date" : initial ? "Persistent Key" : "1 Month");
  const [customDate, setCustomDate] = useState(toDatetimeLocal(initial?.apiKeyExpiresAt));
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [saved, setSaved] = useState(false);
  const isEdit = !!initial?.id;

  function set(field, value) { setForm((f) => ({ ...f, [field]: value })); }

  async function submit(e) {
    e.preventDefault();
    setError("");
    setSaved(false);
    setBusy(true);
    try {
      const apiKeyExpiresAt = computeExpiryFromOption(expiryOption, customDate);
      const savedConn = await onSave({ ...form, label: form.db, apiKeyExpiresAt });
      setSaved(true);
      setBusy(false);
      // Let the success message show briefly before handing off to the parent
      // to navigate to the database connection screen.
      setTimeout(() => onSaved && onSaved(savedConn), 900);
    } catch (err) {
      setError(err?.message || "Something went wrong while saving. Please try again.");
      setBusy(false);
    }
  }

  return (
    <form className="card" onSubmit={submit}>
      <h3 style={{ marginTop: 0 }}>{isEdit ? "Edit Database" : "New Database"}</h3>
      <div className="field">
        <label>Odoo URL (e.g. https://acme.odoo.com)</label>
        <input required value={form.url} onChange={(e) => set("url", e.target.value)} placeholder="https://yourcompany.odoo.com" />
      </div>
      <div className="field">
        <label>Database name</label>
        <input required value={form.db} onChange={(e) => set("db", e.target.value)} />
      </div>
      <div className="field">
        <label>Username / Email</label>
        <input required value={form.username} onChange={(e) => set("username", e.target.value)} />
      </div>
      <div className="field">
        <label>API Key {isEdit && "(leave blank to keep the current one)"}</label>
        <input type="password" required={!isEdit} value={form.apiKey} onChange={(e) => set("apiKey", e.target.value)} />
      </div>
      <div className="field">
        <label>API Key Expiry (match whatever duration you picked in Odoo when generating this key)</label>
        <select value={expiryOption} onChange={(e) => setExpiryOption(e.target.value)}>
          {EXPIRY_OPTIONS.map((opt) => <option key={opt} value={opt}>{opt}</option>)}
        </select>
      </div>
      {expiryOption === "Custom Date" && (
        <div className="field">
          <label>Custom expiry date &amp; time</label>
          <input type="datetime-local" required value={customDate} onChange={(e) => setCustomDate(e.target.value)} />
        </div>
      )}
      <div className="field">
        <label>Google Drive API Key (optional — only needed to expand Drive FOLDER links into every image inside, for Inventory's Extra Images column)</label>
        <input type="password" value={form.driveApiKey || ""} onChange={(e) => set("driveApiKey", e.target.value)} placeholder={isEdit ? "leave blank to keep the current one" : "leave blank if you won't be importing Drive folders"} />
      </div>
      <div className="checkbox-row">
        <input type="checkbox" checked={form.autoCreateSafe !== false} onChange={(e) => set("autoCreateSafe", e.target.checked)} />
        <label>Allow auto-creating missing categories/tags/vendors while importing</label>
      </div>
      {error && <div className="error-box">{error}</div>}
      {saved && <div className="success-box">✓ Saved changes successfully — connecting...</div>}

      <div className="actions-row">
        <button className="btn btn-primary" type="submit" disabled={busy || saved}>
          {busy ? "Saving..." : saved ? "Saved ✓" : "Save & Connect"}
        </button>
        <button className="btn btn-secondary" type="button" onClick={onCancel} disabled={busy || saved}>Cancel</button>
      </div>
    </form>
  );
}
