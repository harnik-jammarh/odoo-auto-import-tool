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
