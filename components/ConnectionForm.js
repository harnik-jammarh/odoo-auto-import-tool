import { useState } from "react";

export default function ConnectionForm({ initial, onSave, onCancel }) {
  const [form, setForm] = useState(initial || { label: "", url: "", db: "", username: "", apiKey: "", driveApiKey: "", autoCreateSafe: true });
  const [busy, setBusy] = useState(false);
  const isEdit = !!initial?.id;

  function set(field, value) { setForm((f) => ({ ...f, [field]: value })); }

  async function submit(e) {
    e.preventDefault();
    setBusy(true);
    try { await onSave(form); } finally { setBusy(false); }
  }

  return (
    <form className="card" onSubmit={submit}>
      <h3 style={{ marginTop: 0 }}>{isEdit ? "Edit Database" : "New Database"}</h3>
      <div className="field">
        <label>Label (just for you, e.g. "Acme Corp — Production")</label>
        <input required value={form.label} onChange={(e) => set("label", e.target.value)} />
      </div>
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
        <label>Google Drive API Key (optional — only needed to expand Drive FOLDER links into every image inside, for Inventory's Extra Images column)</label>
        <input type="password" value={form.driveApiKey || ""} onChange={(e) => set("driveApiKey", e.target.value)} placeholder={isEdit ? "leave blank to keep the current one" : "leave blank if you won't be importing Drive folders"} />
      </div>
      <div className="checkbox-row">
        <input type="checkbox" checked={form.autoCreateSafe !== false} onChange={(e) => set("autoCreateSafe", e.target.checked)} />
        <label>Allow auto-creating missing categories/tags/vendors while importing</label>
      </div>
      <div className="actions-row">
        <button className="btn btn-primary" type="submit" disabled={busy}>{busy ? "Saving..." : "Save & Connect"}</button>
        <button className="btn btn-secondary" type="button" onClick={onCancel}>Cancel</button>
      </div>
    </form>
  );
}
