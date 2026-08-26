# Odoo Auto-Import — Web App

A hosted, team-shared version of the Odoo Auto-Import Chrome extension.
Everything about the actual import logic — module schemas, fuzzy column
matching, GST tax resolution, attribute/variant parsing, account
auto-creation, on-hand quantity, relocation, and the "smart" duplicate check
— is unchanged from the extension. See `lib/odooEngine.js`.

**For non-technical setup instructions, see [SETUP-GUIDE.md](./SETUP-GUIDE.md).**

## Architecture

| Layer | Tool | Role |
|---|---|---|
| Code + versioning | GitHub | Source of truth |
| Hosting + serverless functions | Vercel | Serves the frontend, runs the Odoo proxy |
| Auth | Supabase Auth | Per-teammate email/password login |
| Data | Supabase Postgres + Row Level Security | Each user's saved Odoo connections, visible only to them |
| Secrets | AES-256-GCM (`lib/crypto.js`) | Odoo API keys encrypted before they're stored, decrypted only server-side |
| File parsing | SheetJS (`xlsx`) | 100% client-side, exactly like the extension — files never touch the server |

### Why a server-side proxy is required

A browser page can't call an Odoo server's `/jsonrpc` endpoint directly —
Odoo doesn't return the CORS headers a browser requires for cross-origin
requests (the extension only worked because its manifest's
`host_permissions` bypasses that check). `pages/api/odoo-proxy.js` is a
small serverless function that makes the actual call to Odoo server-to-server
(no CORS restriction applies there) and returns the result. `lib/odooEngine.js`
calls this proxy instead of calling Odoo directly; nothing else about the
import logic changed.

## Local development

```
npm install
cp .env.example .env.local   # then fill in the values (see SETUP-GUIDE.md)
npm run dev
```

## Folder guide

- `lib/odooEngine.js` — the ported import engine (schemas, matching, cleaning, upload logic)
- `lib/crypto.js` — encrypt/decrypt for stored Odoo API keys
- `lib/supabaseClient.js` / `lib/supabaseAdmin.js` — browser vs. server Supabase clients
- `pages/api/odoo-proxy.js` — the CORS-bypassing proxy to Odoo
- `pages/api/connections/*` — CRUD for saved databases, encrypted at rest, ownership-checked
- `pages/dashboard.js`, `pages/login.js` — the app UI
- `components/SheetCard.js`, `components/ConnectionForm.js` — UI pieces
- `supabase-schema.sql` — run once in Supabase's SQL editor to create the table + RLS policies

## Security notes

- Odoo API keys are AES-256-GCM encrypted server-side before being written to Supabase, using a secret (`CONNECTION_ENCRYPTION_KEY`) that only lives in Vercel's environment variables.
- Row Level Security in Postgres ensures a user's saved connections are only ever returned to that same user, enforced by the database itself, not just app code.
- The `/api/connections/*` routes additionally verify the caller's Supabase session token and check row ownership before any read/write — belt and braces alongside RLS.
