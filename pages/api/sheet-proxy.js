// Google Sheets import's server-side proxy — the same fix as
// /api/image-proxy, for the same class of problem. A browser page on this
// app's own domain cannot fetch() docs.google.com's CSV export endpoint
// directly: docs.google.com doesn't send back an Access-Control-Allow-Origin
// header for arbitrary third-party origins, so the browser blocks it. (The
// Chrome extension version of this tool never had this problem — it runs in
// a privileged extension context with host_permissions covering
// "https://*/*", which isn't subject to CORS at all.) A serverless function
// has no such restriction, so the browser calls THIS route with a sheet ID
// (and optional gid), and this route makes the real call on the server and
// hands the CSV text back.
//
// Two failure modes are checked here, not just the HTTP status: (a) a
// private/restricted sheet's export endpoint 302-redirects to a Google
// sign-in page rather than erroring outright, so the final redirected URL
// is checked too; (b) even past that check, the body itself can be an HTML
// sign-in/permission page rather than CSV — checked again client-side in
// importGoogleSheet as a backstop.

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const { sheetId, gid } = req.body || {};
  if (!sheetId || typeof sheetId !== "string") {
    res.status(400).json({ error: "Missing sheetId" });
    return;
  }

  const exportUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/export?format=csv${gid ? `&gid=${encodeURIComponent(gid)}` : ""}`;

  try {
    const sheetRes = await fetch(exportUrl);
    // A private (or link-restricted) sheet's export endpoint 302s to a
    // Google accounts sign-in page rather than erroring outright — the
    // fetch itself "succeeds" with HTML content, so both the HTTP status
    // and the final URL are checked, not just res.ok.
    const finalUrlLooksLikeLogin = /accounts\.google\.com/i.test(sheetRes.url);
    if (!sheetRes.ok || finalUrlLooksLikeLogin) {
      res.status(200).json({
        error: `couldn't fetch this sheet (HTTP ${sheetRes.status}) — make sure it's shared as "Anyone with the link" (Viewer access is enough), not just to specific people.`,
      });
      return;
    }
    const csvText = await sheetRes.text();
    if (/^\s*<!doctype html/i.test(csvText) || /^\s*<html/i.test(csvText)) {
      res.status(200).json({
        error: 'got a sign-in/permission page back instead of your sheet\'s data — it needs to be shared as "Anyone with the link" first.',
      });
      return;
    }
    res.status(200).json({ csv: csvText });
  } catch (e) {
    res.status(200).json({ error: e.message || "Could not fetch that Google Sheet" });
  }
}
