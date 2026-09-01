import { useState, useEffect, useCallback, useRef } from "react";
import { useRouter } from "next/router";
import { supabase } from "../lib/supabaseClient";
import { createEngine, parseWorkbookFile } from "../lib/odooEngine";
import ConnectionForm from "../components/ConnectionForm";
import SheetCard from "../components/SheetCard";

export default function Dashboard() {
  const router = useRouter();
  const [session, setSession] = useState(undefined); // undefined = loading, null = signed out
  const [view, setView] = useState("databases"); // databases | form | main
  const [connections, setConnections] = useState([]);
  const [editingConn, setEditingConn] = useState(null); // null = add new
  const [activeConn, setActiveConn] = useState(null);
  const [engine, setEngine] = useState(null);
  const [connectionStatus, setConnectionStatus] = useState("unknown");
  const [stockLocations, setStockLocations] = useState([]);
  const [stockLocationsLoading, setStockLocationsLoading] = useState(false);
  const [journals, setJournals] = useState([]);
  const [journalsLoading, setJournalsLoading] = useState(false);
  const [customFieldsStatus, setCustomFieldsStatus] = useState("idle"); // idle | loading | done
  const [sheets, setSheets] = useState([]);
  const [fileName, setFileName] = useState("");
  const [processing, setProcessing] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [loadError, setLoadError] = useState("");
  const fileInputRef = useRef(null);

  // ---- auth guard -----------------------------------------------------
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

  const loadConnections = useCallback(async () => {
    try {
      const data = await authedFetch("/api/connections");
      setConnections(data.connections);
    } catch (e) {
      setLoadError(e.message);
    }
  }, [authedFetch]);

  useEffect(() => {
    if (session) loadConnections();
  }, [session, loadConnections]);

  async function signOut() {
    await supabase.auth.signOut();
    router.replace("/login");
  }

  // ---- connection management -------------------------------------------
  async function saveConnection(form) {
    if (editingConn?.id) {
      await authedFetch(`/api/connections/${editingConn.id}`, { method: "PUT", body: JSON.stringify(form) });
      await loadConnections();
      activateConnection({ ...editingConn, ...form });
    } else {
      const { id } = await authedFetch("/api/connections", { method: "POST", body: JSON.stringify(form) });
      await loadConnections();
      activateConnection({ id, ...form });
    }
  }

  async function deleteConnection(id) {
    if (!confirm("Delete this saved database? This can't be undone.")) return;
    await authedFetch(`/api/connections/${id}`, { method: "DELETE" });
    if (activeConn?.id === id) { setActiveConn(null); setEngine(null); setConnectionStatus("unknown"); }
    await loadConnections();
  }

  function activateConnection(conn) {
    setActiveConn(conn);
    const eng = createEngine(conn);
    setEngine(eng);
    setView("main");
    setSheets([]);
    setFileName("");
    setConnectionStatus("unknown");
    checkConn(eng, conn);
  }

  async function checkConn(eng, conn) {
    setConnectionStatus("unknown");
    if (!conn.url || !conn.db || !conn.username || !conn.apiKey) { setConnectionStatus("offline"); return; }
    const ok = await eng.checkConnection();
    setConnectionStatus(ok ? "online" : "offline");
    if (ok) {
      setStockLocationsLoading(true);
      const locs = await eng.loadStockLocations();
      setStockLocations(locs);
      setStockLocationsLoading(false);

      setJournalsLoading(true);
      eng.loadJournals().then(setJournals).catch(() => setJournals([])).finally(() => setJournalsLoading(false));

      setCustomFieldsStatus("loading");
      eng.refreshCustomFields()
        .then((foundAny) => setCustomFieldsStatus(foundAny ? "done" : "idle"))
        .catch(() => setCustomFieldsStatus("error"));
    }
  }

  // ---- file handling -----------------------------------------------------
  async function handleFile(file) {
    if (!file) return;
    setProcessing(true);
    setFileName(file.name);
    try {
      const parsed = await parseWorkbookFile(file);
      setSheets(parsed);
    } catch (e) {
      alert("Could not read that file: " + e.message);
    } finally {
      setProcessing(false);
    }
  }

  function resetFile() {
    setSheets([]);
    setFileName("");
  }

  function updateSheet(idx, updated) {
    setSheets((prev) => prev.map((s, i) => (i === idx ? updated : s)));
  }

  if (session === undefined) return <div className="page">Loading...</div>;
  if (!session) return null;

  return (
    <div className="page">
      <div className="header-row">
        <div className="logo">O</div>
        <h1>Odoo Auto-Import</h1>
      </div>
      <p className="subtitle">Excel/CSV uploads straight into Odoo, from anywhere — no extension to install.</p>

      <div className="top-row">
        <div className="status-row">
          {view === "main" && (
            <>
              <span className={`dot ${connectionStatus}`} />
              <span>
                {connectionStatus === "online" ? `Connected — ${activeConn?.label}` : connectionStatus === "offline" ? "Not connected" : "Checking..."}
              </span>
              {connectionStatus === "online" && customFieldsStatus === "loading" && (
                <span className="sheet-meta" style={{ marginLeft: 10 }}>· checking for custom fields...</span>
              )}
              {connectionStatus === "online" && customFieldsStatus === "done" && (
                <span className="sheet-meta" style={{ marginLeft: 10 }}>· custom fields loaded</span>
              )}
              {connectionStatus === "online" && customFieldsStatus === "error" && (
                <span className="sheet-meta" style={{ marginLeft: 10, color: "#b91c1c" }}>· custom field check failed</span>
              )}
            </>
          )}
        </div>
        <div style={{ display: "flex", gap: 14 }}>
          <button className="link-btn" onClick={() => setView(view === "main" ? "databases" : "main")}>
            {view === "main" ? "Databases" : "Close"}
          </button>
          <button className="link-btn" onClick={() => router.push("/profile")}>My Profile</button>
          <button className="link-btn" onClick={signOut}>Sign Out</button>
        </div>
      </div>

      {loadError && <div className="error-box">{loadError}</div>}

      {view === "databases" && (
        <DatabasesList
          connections={connections}
          activeId={activeConn?.id}
          onAdd={() => { setEditingConn(null); setView("form"); }}
          onEdit={(c) => { setEditingConn(c); setView("form"); }}
          onActivate={activateConnection}
          onDelete={deleteConnection}
        />
      )}

      {view === "form" && (
        <ConnectionForm
          initial={editingConn}
          onSave={saveConnection}
          onCancel={() => setView(editingConn ? "main" : "databases")}
        />
      )}

      {view === "main" && (
        <>
          {connectionStatus !== "online" && <div className="note-box">Select or connect a database above to get started.</div>}

          {connectionStatus === "online" && !sheets.length && (
            <div
              className={`dropzone ${dragOver ? "drag" : ""}`}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
              onClick={() => fileInputRef.current?.click()}
            >
              <p>📄 Drag an Excel/CSV file here, or click to choose one</p>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                style={{ display: "none" }}
                onChange={(e) => handleFile(e.target.files[0])}
              />
            </div>
          )}

          {processing && <p className="subtitle">Processing {fileName}...</p>}

          {sheets.length > 0 && !processing && (
            <>
              <div className="file-bar">
                <span>📄 {fileName} · {sheets.length} sheet(s) detected</span>
                <button className="link-btn" onClick={resetFile}>Upload a different file</button>
              </div>
              {sheets.map((sheet, idx) => (
                <SheetCard
                  key={idx}
                  sheet={sheet}
                  idx={idx}
                  engine={engine}
                  stockLocations={stockLocations}
                  stockLocationsLoading={stockLocationsLoading}
                  journals={journals}
                  journalsLoading={journalsLoading}
                  onChange={updateSheet}
                />
              ))}
            </>
          )}
        </>
      )}
    </div>
  );
}

function DatabasesList({ connections, activeId, onAdd, onEdit, onActivate, onDelete }) {
  return (
    <div>
      <button className="btn btn-primary" onClick={onAdd}>+ New Database</button>
      <div className="db-list">
        {connections.map((c) => (
          <div key={c.id} className={`db-item ${c.id === activeId ? "active" : ""}`}>
            <div className="db-item-main" onClick={() => onActivate(c)}>
              <div className="db-item-label">{c.label}</div>
              <div className="db-item-sub">{c.db} · {c.username}</div>
            </div>
            <div className="db-item-actions">
              <button className="btn btn-secondary" onClick={() => onEdit(c)}>Edit</button>
              <button className="btn btn-danger" onClick={() => onDelete(c.id)}>Delete</button>
            </div>
          </div>
        ))}
      </div>
      {!connections.length && <p className="subtitle">No databases saved yet — add one to get started.</p>}
    </div>
  );
}
