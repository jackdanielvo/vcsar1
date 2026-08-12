-- VCSAR Team 1 — admin allowlist
--
-- WHY: previously ANY signed-in Supabase user had full access. This makes
-- access depend on being on an approved list, so you can add/remove admins
-- from the website's admin page instead of the Supabase dashboard.
--
-- SAFE TO RUN: it first copies every EXISTING login into the allowlist, so
-- you cannot lock yourself out.
--
-- HOW TO RUN
-- Supabase -> SQL Editor -> New query -> paste this whole file -> Run.

-- 1. The allowlist ------------------------------------------------------
create table if not exists public.admins (
  email    text primary key,
  added_by text,
  added_at timestamptz default now()
);

alter table public.admins enable row level security;

-- Everyone who can already log in stays an admin (no lock-out).
insert into public.admins (email, added_by)
select email, 'initial setup' from auth.users
on conflict (email) do nothing;

-- 2. Helper: is the current user on the allowlist? ----------------------
-- SECURITY DEFINER lets this read the table without triggering its own
-- row-level rules (which would otherwise recurse forever).
create or replace function public.is_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.admins a
    where lower(a.email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  );
$$;

-- 3. Only admins can see or change the allowlist ------------------------
drop policy if exists "admins read"   on public.admins;
drop policy if exists "admins insert" on public.admins;
drop policy if exists "admins delete" on public.admins;

create policy "admins read"   on public.admins for select to authenticated using (public.is_admin());
create policy "admins insert" on public.admins for insert to authenticated with check (public.is_admin());
create policy "admins delete" on public.admins for delete to authenticated using (public.is_admin());

-- 4. Lock the roster down to admins -------------------------------------
drop policy if exists "auth insert" on public.members;
drop policy if exists "auth update" on public.members;
drop policy if exists "auth delete" on public.members;

create policy "admin insert" on public.members for insert to authenticated with check (public.is_admin());
create policy "admin update" on public.members for update to authenticated using (public.is_admin());
create policy "admin delete" on public.members for delete to authenticated using (public.is_admin());
-- (public read stays as-is so the Team page still works for visitors)

-- 5. Lock the candidate-email key down to admins ------------------------
drop policy if exists "authenticated read" on public.app_secrets;
drop policy if exists "admins read secret" on public.app_secrets;
create policy "admins read secret" on public.app_secrets
  for select to authenticated using (public.is_admin());

-- Check who is currently an admin:
-- select * from public.admins;
