# Setup Guide — everything done in your browser, no terminal needed

This turns the files in this project into a live website your whole team can
log into. You'll use three free/cheap websites: **GitHub** (stores the code),
**Vercel** (hosts the website), **Supabase** (logins + saved databases).
Total time: about 20–30 minutes, done once. After that, updating the app is
just re-uploading files to GitHub.

---

## Part 1 — Put the code on GitHub

1. Go to [github.com](https://github.com) and sign in (or create a free account).
2. Click the **+** icon top-right → **New repository**.
3. Name it something like `odoo-auto-import-web`. Keep it **Private** (recommended, since it's internal team tooling). Click **Create repository**.
4. On the new (empty) repo page, click **uploading an existing file** (a blue link in the middle of the page).
5. Open the folder of files you downloaded from this conversation on your computer. Select **all files and folders inside it** (not the outer folder itself) and drag them into the browser window.
   - GitHub will show a list of files being uploaded. Wait for it to finish — this can take a minute for larger folders.
6. Scroll down, type a commit message like "Initial upload", and click **Commit changes**.

Your code is now on GitHub. Any time you want to update the app later, come back to this repo, click into a file (or use the same "Add file → Upload files" button), replace it, and commit — Vercel will automatically redeploy.

---

## Part 2 — Create your Supabase project (logins + saved databases)

1. Go to [supabase.com](https://supabase.com) → **Start your project** → sign in with GitHub (easiest) or email.
2. Click **New project**. Pick any name (e.g. "odoo-import"), set a database password (save it somewhere, though you won't need it day-to-day), pick a region close to you, and click **Create new project**. Wait ~2 minutes while it sets up.
3. Once it's ready, click the **SQL Editor** icon in the left sidebar → **New query**.
4. Open the file `supabase-schema.sql` from this project (in GitHub, click on it, then click the "Raw" button to see the plain text, and copy it all).
5. Paste it into the Supabase SQL editor and click **Run** (bottom right). You should see "Success. No rows returned."
6. In the left sidebar, click the gear icon **Project Settings** → **API**.
   - You'll see **Project URL** — copy it somewhere (you'll need it soon).
   - Under **Project API keys**, copy the **anon / public** key.
   - Also copy the **service_role** key (click "Reveal" first). This one is secret — never share it or put it in code that's visible to visitors.

**Optional but recommended — restrict who can sign up:**
Since anyone with the link can currently create an account, you may want to limit signups to your company's email domain. In Supabase: **Authentication** (left sidebar) → **Providers** → **Email**, and/or **Authentication** → **Policies**, you can add checks later. For now, simplest option: just don't share the URL publicly, and share it directly with your teammates.

---

## Part 3 — Generate your encryption key

Saved Odoo API keys are encrypted before they're stored, using a secret key only you control. You need to generate a random one:

1. Go to [https://www.uuidgenerator.net/](https://www.uuidgenerator.net/) or any "random hex string generator" site, or simply open this project's `CONNECTION_ENCRYPTION_KEY` helper: go to [https://generate-random.org/api-key-generator?count=1&length=64&type=hex](https://generate-random.org/api-key-generator?count=1&length=64&type=hex) and copy the 64-character result.
2. Save this string somewhere safe (e.g. a password manager). You'll paste it into Vercel in the next part. If you ever lose it, previously saved Odoo connections can't be decrypted and everyone will need to re-enter their database details (nothing else breaks).

---

## Part 4 — Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) → sign up/sign in with your **GitHub** account (this makes the next step one click).
2. Click **Add New...** → **Project**.
3. Find your `odoo-auto-import-web` repo in the list and click **Import**.
4. Vercel will detect it's a Next.js project automatically — you don't need to change any build settings.
5. Before clicking Deploy, open **Environment Variables** and add these four, one at a time (Name on the left, Value on the right):

   | Name | Value |
   |---|---|
   | `NEXT_PUBLIC_SUPABASE_URL` | the Project URL from Supabase (Part 2, step 6) |
   | `NEXT_PUBLIC_SUPABASE_ANON_KEY` | the anon/public key from Supabase |
   | `SUPABASE_SERVICE_ROLE_KEY` | the service_role key from Supabase |
   | `CONNECTION_ENCRYPTION_KEY` | the 64-character random string from Part 3 |

6. Click **Deploy**. Wait 1–2 minutes.
7. When it finishes, click **Visit** — you'll see your live site! The URL will look like `https://odoo-auto-import-web.vercel.app`. Bookmark it and share it with your teammates.

---

## Part 5 — First-time use

1. Open your new site's URL. Click **Create Account**, enter your work email and a password, and confirm your email (check your inbox for a confirmation link — Supabase sends this automatically).
2. Sign in, click **+ New Database**, and fill in your Odoo URL, database name, username, and API key — same details you used in the old extension.
3. Click **Save & Connect**. You should see a green dot and "Connected."
4. Drag in your Excel/CSV file exactly like before — same modules, same column detection, same everything.

Each teammate repeats step 1–3 with their own login and their own database(s) — nobody sees anyone else's saved connections.

---

## Updating the app later (no terminal, ever)

1. Go to your GitHub repo → find the file you want to change → click the pencil (✏️) icon to edit it directly in the browser, or use **Add file → Upload files** to replace several files at once.
2. Commit the change.
3. That's it — Vercel automatically rebuilds and redeploys within about a minute. Everyone's browser gets the new version the next time they load the page. No uninstalling, no reinstalling, no re-entering credentials.

---

## Troubleshooting

- **"Not connected" / red dot after saving a database** — double check the Odoo URL (must start with `https://`), database name, username, and API key. Try authenticating with the same details directly in Odoo to confirm the API key is still valid.
- **Sign-in confirmation email never arrives** — check spam, or in Supabase go to **Authentication → Users**, find the account, and manually confirm it.
- **A client's Odoo has IP restrictions** — if a specific client's Odoo instance only allows certain IP addresses to connect, Vercel's servers use dynamic IPs, so you may need to ask that client's IT team to allowlist Vercel's IP ranges, or skip IP restriction for the API user.
- **Something looks broken after an update** — in Vercel, go to your project → **Deployments**, find the last working deployment, click the **⋯** menu → **Promote to Production** to instantly roll back.
