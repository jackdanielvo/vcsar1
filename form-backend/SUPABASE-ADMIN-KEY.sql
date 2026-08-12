-- VCSAR Team 1 — admin key for the "Email Candidates" panel
--
-- WHAT THIS DOES
-- Stores one private key that proves a request came from a signed-in admin.
-- Row Level Security means ONLY logged-in users can read it — anonymous
-- visitors to the website cannot, so they can never trigger candidate emails.
--
-- HOW TO RUN IT
-- 1. Open your Supabase project
-- 2. Left sidebar: "SQL Editor" -> "New query"
-- 3. Paste this whole file and click "Run"
-- You should see "Success. No rows returned."

create table if not exists public.app_secrets (
  key   text primary key,
  value text not null
);

alter table public.app_secrets enable row level security;

-- Only signed-in admins can read the key. (No public/anon access at all.)
drop policy if exists "authenticated read" on public.app_secrets;
create policy "authenticated read" on public.app_secrets
  for select to authenticated using (true);

-- The key itself — must match ANNOUNCE_SECRET in Code.gs exactly.
insert into public.app_secrets (key, value)
values ('announce', 'hbyXNMF17eL9bvZRWyy8NifgXAtGxdg8IYBvLdwtY00')
on conflict (key) do update set value = excluded.value;
