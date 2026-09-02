// Image import's server-side proxy — the same fix as /api/odoo-proxy, for a
// different problem. A browser page on this app's own domain cannot fetch()
// Google Drive, googleusercontent.com, docs.google.com, or an arbitrary
// product-photo URL directly: none of those origins send back an
// Access-Control-Allow-Origin header for this app's domain, so the browser
// blocks it. (The Chrome extension version of this tool never had this
// problem — it runs in a privileged extension context with
// host_permissions covering "https://*/*", which isn't subject to CORS at
// all.) A serverless function has no such restriction, so the browser
// calls THIS route for three things, and this route makes the real call on
// the server and hands the result back:
//
//   action: "fetch"       -> download one image URL, return its bytes as base64
//   action: "driveList"   -> list the image files inside a public Drive folder
//                             (Drive API v3 files.list, one page at a time)
//   action: "googleSheet" -> download a Google Sheet's CSV export (see
//                             fetchGoogleSheet in lib/odooEngine.js) — same
//                             CORS problem, docs.google.com doesn't send an
//                             Access-Control-Allow-Origin header either.

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const { action } = req.body || {};

  if (action === "fetch") {
    const { url } = req.body || {};
    let parsed;
    try {
      parsed = new URL(url);
      if (parsed.protocol !== "https:" && parsed.protocol !== "http:") throw new Error("bad protocol");
    } catch (e) {
      res.status(400).json({ error: "Invalid image URL" });
      return;
    }
    try {
      const imgRes = await fetch(parsed.toString());
      if (!imgRes.ok) {
        res.status(200).json({ error: `HTTP ${imgRes.status}` });
        return;
      }
      const contentType = imgRes.headers.get("content-type") || "";
      // A large Drive file (roughly >25-100MB, sometimes smaller) returns an
      // HTML "can't scan this file for viruses" confirmation page instead of
      // the image — catch that here rather than silently saving the HTML
      // page as the product photo.
      if (contentType.includes("text/html")) {
        res.status(200).json({ error: "got an HTML page instead of an image" });
        return;
      }
      const buf = Buffer.from(await imgRes.arrayBuffer());
      res.status(200).json({ base64: buf.toString("base64"), contentType });
    } catch (e) {
      res.status(200).json({ error: e.message || "Could not fetch that image" });
    }
    return;
  }

  if (action === "driveList") {
    const { folderId, apiKey, pageToken } = req.body || {};
    if (!folderId || !apiKey) {
      res.status(400).json({ error: "Missing folderId or apiKey" });
      return;
    }
    try {
      const q = encodeURIComponent(`'${folderId}' in parents and mimeType contains 'image/' and trashed = false`);
      const url = `https://www.googleapis.com/drive/v3/files?q=${q}&key=${encodeURIComponent(apiKey)}&fields=nextPageToken,files(id,name)&pageSize=100${pageToken ? `&pageToken=${encodeURIComponent(pageToken)}` : ""}`;
      const driveRes = await fetch(url);
      const data = await driveRes.json();
      if (!driveRes.ok) {
        res.status(200).json({ error: data.error?.message || `HTTP ${driveRes.status}` });
        return;
      }
      res.status(200).json({ files: data.files || [], nextPageToken: data.nextPageToken || null });
    } catch (e) {
      res.status(200).json({ error: e.message || "Drive API folder lookup failed" });
    }
    return;
  }

  if (action === "googleSheet") {
    const { url } = req.body || {};
    let parsed;
    try {
      parsed = new URL(url);
      if (parsed.hostname !== "docs.google.com") throw new Error("not a docs.google.com URL");
    } catch (e) {
      res.status(400).json({ error: "Invalid Google Sheets export URL" });
      return;
    }
    try {
      const sheetRes = await fetch(parsed.toString());
      // A private (or link-restricted) sheet's export endpoint 302s to a
      // Google accounts sign-in page rather than erroring outright — fetch()
      // follows that redirect automatically, so both the final response
      // status/URL and the body itself (checked back in fetchGoogleSheet)
      // need checking, not just sheetRes.ok.
      const finalUrlLooksLikeLogin = /accounts\.google\.com/i.test(sheetRes.url || "");
      if (!sheetRes.ok || finalUrlLooksLikeLogin) {
        res.status(200).json({ error: `HTTP ${sheetRes.status}` });
        return;
      }
      const csvText = await sheetRes.text();
      res.status(200).json({ csvText });
    } catch (e) {
      res.status(200).json({ error: e.message || "Could not fetch that sheet" });
    }
    return;
  }

  res.status(400).json({ error: "Unknown or missing action" });
}
