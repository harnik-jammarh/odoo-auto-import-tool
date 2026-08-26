import { useState } from "react";
import { ODOO_SCHEMAS, mapForModule, confidenceColor } from "../lib/odooEngine";

export default function SheetCard({ sheet, idx, engine, stockLocations, stockLocationsLoading, onChange }) {
  const [expanded, setExpanded] = useState(true);
  const [uploadState, setUploadState] = useState(null);

  const schema = ODOO_SCHEMAS[sheet.analysis.moduleKey];
  const mapping = sheet.analysis.mapping;

  function patch(fields) {
    onChange(idx, { ...sheet, ...fields });
  }

  function changeModule(moduleKey) {
    patch({ analysis: { moduleKey, mapping: mapForModule(sheet.headers, moduleKey) } });
    setUploadState(null);
  }

  const needsLocation = mapping.some((m) => m.field?.isQuantity) && schema.model === "product.template";
  const requiredFields = schema.fields.filter((f) => f.required);
  const missingRequired = requiredFields.filter((f) => !mapping.some((m) => m.field?.name === f.name));
  const blockers = [];
  if (missingRequired.length) blockers.push(`Required field(s) not mapped: ${missingRequired.map((f) => f.label).join(", ")}`);
  if (needsLocation && !sheet.stockLocationId) blockers.push("Choose a location for On Hand Quantity below");

  async function doUpload() {
    setUploadState({ status: "checking" });
    const result = await engine.uploadSheetToOdoo(sheet, (partial) => setUploadState(partial));
    setUploadState(result);
  }

  function confirmCreateAndContinue() {
    patch({ autoCreateProducts: true });
    setTimeout(doUpload, 0);
  }

  return (
    <div className="sheet-card">
      <div className="sheet-card-header" onClick={() => setExpanded((e) => !e)}>
        <div>
          <strong>{schema.icon} {sheet.name}</strong>{" "}
          <span className="sheet-meta">→ {schema.label} · {sheet.rows.length} row(s)</span>
        </div>
        <span className="link-btn">{expanded ? "Collapse" : "Expand"}</span>
      </div>

      {expanded && (
        <div className="sheet-card-body">
          <div className="module-row">
            <label className="module-label">Import as:</label>
            <select className="module-select" value={sheet.analysis.moduleKey} onChange={(e) => changeModule(e.target.value)}>
              {Object.entries(ODOO_SCHEMAS).map(([key, s]) => (
                <option key={key} value={key}>{s.icon} {s.label}</option>
              ))}
            </select>
          </div>
          <div className="note-box">{schema.note}</div>

          <table className="mapping-table">
            <thead>
              <tr><th>Your column</th><th>Maps to</th><th>Confidence</th></tr>
            </thead>
            <tbody>
              {mapping.map((m, i) => (
                <tr key={i}>
                  <td>{m.header}</td>
                  <td>{m.field ? m.field.label : <em style={{ color: "#9CA3AF" }}>not mapped</em>}</td>
                  <td>
                    {m.field && (
                      <span className="confidence-pill" style={{ background: confidenceColor(m.confidence) }}>
                        {Math.round(m.confidence * 100)}%
                      </span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          {blockers.map((b, i) => <div key={i} className="warn-box">⚠️ {b}</div>)}

          {needsLocation && (
            <>
              <div className="module-row">
                <label className="module-label">Location:</label>
                {stockLocationsLoading ? (
                  <span className="sheet-meta">Loading locations...</span>
                ) : (
                  <select
                    className="module-select"
                    value={sheet.stockLocationId || ""}
                    onChange={(e) => patch({ stockLocationId: e.target.value ? Number(e.target.value) : null })}
                  >
                    <option value="">{stockLocations.length ? "— choose a location —" : "No internal locations found"}</option>
                    {stockLocations.map((loc) => (
                      <option key={loc.id} value={loc.id}>{loc.display_name}</option>
                    ))}
                  </select>
                )}
              </div>
              <div className="note-box">All On Hand Quantity values in this sheet will be set at the location chosen above — this replaces that product's counted quantity there, it does not add to existing stock elsewhere.</div>
            </>
          )}

          <div className="dedupe-row">
            <label>
              <input
                type="checkbox"
                checked={sheet.skipDuplicates !== false}
                onChange={(e) => patch({ skipDuplicates: e.target.checked })}
              />{" "}
              Check for records that already exist in Odoo
            </label>
          </div>

          {sheet.skipDuplicates !== false && (
            <>
              <div className="module-row">
                <label className="module-label">If it already exists:</label>
                <select
                  className="module-select"
                  value={sheet.duplicateAction || "smart"}
                  onChange={(e) => patch({ duplicateAction: e.target.value })}
                >
                  <option value="smart">Update it — unless it has transactions, then skip it (recommended)</option>
                  <option value="update">Always update it, even if it has transactions</option>
                  <option value="skip">Always skip it (leave the existing record untouched)</option>
                </select>
              </div>
              {(sheet.duplicateAction || "smart") === "smart" && schema.model === "product.template" && (
                <div className="note-box">For each existing product this checks real warehouse activity — receipts, deliveries, transfers, sales/purchase orders, invoices/bills. No activity found → the row's mapped fields are written onto it. Activity found → it's left untouched and the reason is listed in the results below.</div>
              )}
              {(sheet.duplicateAction || "smart") === "update" && (
                <div className="warn-box">⚠️ Existing records will be overwritten even if they already have transactions against them — this can make those transactions inconsistent.</div>
              )}
            </>
          )}

          <div className="actions-row">
            <button
              className="btn btn-primary"
              disabled={blockers.length > 0 || uploadState?.status === "uploading" || uploadState?.status === "checking" || uploadState?.status === "confirm-create"}
              onClick={doUpload}
            >
              {uploadState?.status === "uploading" ? `Uploading ${uploadState.progress}/${uploadState.total}...`
                : uploadState?.status === "checking" ? "Checking products..."
                : "Upload to Odoo"}
            </button>
          </div>

          {uploadState?.status === "confirm-create" && (
            <>
              <div className="warn-box">
                This sheet references {uploadState.missingProducts.length} product(s) that don't exist in Odoo yet. Creating them from transaction data alone means guessing sales tax, unit of measure, and inventory tracking — worth a quick look first:
                <ul className="confirm-create-list">
                  {uploadState.missingProducts.map((n) => <li key={n}>{n}</li>)}
                </ul>
              </div>
              <div className="actions-row">
                <button className="btn btn-primary" onClick={confirmCreateAndContinue}>Create them as basic products & continue</button>
                <button className="btn btn-secondary" onClick={() => setUploadState(null)}>Cancel — I'll create/fix them in Odoo first</button>
              </div>
              <div className="note-box">Products created this way are Goods, no sales/purchase tax, tracking off, price = 0 — plain placeholders. Review each one afterward (a later import can update them again once you've set them up).</div>
            </>
          )}

          {uploadState?.status === "done" && (
            <>
              <div className="success-box">
                ✅ {uploadState.result.created_count} of {uploadState.result.total} record(s) created.
                {uploadState.result.duplicate_count > 0 && (
                  uploadState.result.duplicate_action === "skip"
                    ? ` ${uploadState.result.duplicate_count} skipped as duplicates.`
                    : ` ${uploadState.result.duplicate_count} existing record(s) matched — see notes below.`
                )}
                {uploadState.result.failed_count > 0 && ` ${uploadState.result.failed_count} failed — see below.`}
              </div>
              {uploadState.result.notes.map((n, i) => <div key={i} className="note-box">ℹ️ {n}</div>)}
              {uploadState.result.errors.map((e, i) => <div key={i} className="error-box">Row {e.row + 1}: {e.error}</div>)}
            </>
          )}
        </div>
      )}
    </div>
  );
}
