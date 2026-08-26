-- Run this once in Supabase: Dashboard -> SQL Editor -> New query -> paste
-- this whole file -> Run. See SETUP-GUIDE.md Part 3 for step-by-step
-- screenshots-style instructions if you're not comfortable with SQL.

create table if not exists saved_connections (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  label text not null,
  url text not null,
  db text not null,
  username text not null,
  api_key_encrypted text not null,   -- AES-256-GCM encrypted, never plaintext
  auto_create_safe boolean not null default true,
  created_at timestamptz not null default now()
);

-- Row Level Security: turn it on, then only allow a user to see/change
-- rows where user_id matches their own logged-in id. Even if two people
-- share this same database, teammate B can never read teammate A's rows.
alter table saved_connections enable row level security;

create policy "Users can view their own connections"
  on saved_connections for select
  using (auth.uid() = user_id);

create policy "Users can insert their own connections"
  on saved_connections for insert
  with check (auth.uid() = user_id);

create policy "Users can update their own connections"
  on saved_connections for update
  using (auth.uid() = user_id);

create policy "Users can delete their own connections"
  on saved_connections for delete
  using (auth.uid() = user_id);

-- Note: the app's /api/connections routes actually use Supabase's secret
-- "service role" key on the server, which bypasses RLS by design (that's
-- what lets the server encrypt/decrypt on the user's behalf) — those routes
-- do their own manual "does this row belong to this logged-in user" check
-- in code. The RLS policies above are a second, independent safety net:
-- even if a future feature ever queries this table directly from the
-- browser with the public anon key, the database itself still refuses to
-- return another user's rows.

-- ---------------------------------------------------------------------------
-- Waitlist: new signups don't hit Supabase Auth directly. They land here as
-- a pending request; only the admin (ADMIN_EMAIL env var) can approve one,
-- which is what actually creates the Supabase Auth user (via
-- /api/admin/waitlist, using the service role key — see that route).
-- ---------------------------------------------------------------------------
create table if not exists waitlist_requests (
  id uuid primary key default gen_random_uuid(),
  email text not null unique,
  note text,
  status text not null default 'pending', -- pending | approved | rejected
  created_at timestamptz not null default now(),
  decided_at timestamptz
);

-- RLS on: nobody can read/write this table directly from the browser with
-- the anon key. Both the public "request access" route and the admin
-- approve/reject route run server-side with the service role key, which
-- bypasses RLS by design (same pattern as saved_connections above).
alter table waitlist_requests enable row level security;

