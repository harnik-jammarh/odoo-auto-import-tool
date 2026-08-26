// This is the one piece that makes the whole web-app approach possible.
// A browser page on your Vercel domain cannot fetch() an Odoo server
// directly — Odoo doesn't send back the CORS headers a browser requires
// for a cross-origin request, so the browser blocks it. A serverless
// function has no such restriction (server-to-server calls aren't subject
// to CORS at all), so the browser calls THIS route, and this route makes
// the real call to Odoo and hands the result back.
//
// This route does not require Supabase login by itself — the Odoo
// credentials it forwards (db/username/apiKey) are the real gatekeeper,
// exactly as they are for the extension. Anyone calling it without valid
// Odoo credentials just gets an Odoo authentication error back.

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: { message: "Method not allowed" } });
    return;
  }

  const { url, service, method, args } = req.body || {};

  if (!url || !service || !method || !Array.isArray(args)) {
    res.status(400).json({ error: { message: "Missing url, service, method, or args" } });
    return;
  }

  let odooUrl;
  try {
    odooUrl = new URL(url);
    if (odooUrl.protocol !== "https:" && odooUrl.protocol !== "http:") throw new Error("bad protocol");
  } catch (e) {
    res.status(400).json({ error: { message: "Invalid Odoo URL" } });
    return;
  }

  try {
    const odooRes = await fetch(`${url.replace(/\/$/, "")}/jsonrpc`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        jsonrpc: "2.0",
        method: "call",
        params: { service, method, args },
        id: Math.floor(Math.random() * 1e9),
      }),
    });
    const data = await odooRes.json();
    res.status(200).json(data);
  } catch (e) {
    res.status(502).json({ error: { message: `Could not reach Odoo at ${url}: ${e.message}` } });
  }
}
