-- Repair profile rows and friend IDs for existing and future users.
-- Run this in Supabase SQL Editor or through the Supabase migration flow.

alter table public.profiles
  add column if not exists friend_id text;

create or replace function public.compute_friend_id(p_user_id uuid)
returns text
language plpgsql
immutable
set search_path = public
as $$
declare
  clean text := replace(p_user_id::text, '-', '');
  result integer := 0;
  digit integer;
  ch text;
  alphabet constant text := '0123456789abcdef';
begin
  for i in 1..length(clean) loop
    ch := lower(substr(clean, i, 1));
    digit := strpos(alphabet, ch) - 1;
    if digit < 0 then
      raise exception 'invalid uuid character in %', p_user_id;
    end if;
    result := mod((result * 16) + digit, 1000000);
  end loop;

  return lpad(result::text, 6, '0');
end;
$$;

insert into public.profiles (id, email, full_name, friend_id)
select
  u.id,
  u.email,
  nullif(u.raw_user_meta_data ->> 'full_name', ''),
  public.compute_friend_id(u.id)
from auth.users u
where not exists (
  select 1 from public.profiles p where p.id = u.id
)
on conflict (id) do update set
  email = coalesce(public.profiles.email, excluded.email),
  full_name = coalesce(public.profiles.full_name, excluded.full_name),
  friend_id = coalesce(nullif(public.profiles.friend_id, ''), excluded.friend_id),
  updated_at = now();

update public.profiles
set friend_id = public.compute_friend_id(id),
    updated_at = now()
where friend_id is null or btrim(friend_id) = '';

create unique index if not exists profiles_friend_id_idx
  on public.profiles (friend_id)
  where friend_id is not null;

drop policy if exists "profiles_select_own" on public.profiles;
drop policy if exists "profiles_select_authenticated" on public.profiles;
create policy "profiles_select_authenticated" on public.profiles
  for select to authenticated using (true);

drop policy if exists "profiles_insert_own" on public.profiles;
create policy "profiles_insert_own" on public.profiles
  for insert to authenticated
  with check (auth.uid() = id);

drop policy if exists "profiles_update_own" on public.profiles;
create policy "profiles_update_own" on public.profiles
  for update to authenticated
  using (auth.uid() = id)
  with check (auth.uid() = id);

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email, full_name, friend_id)
  values (
    new.id,
    new.email,
    nullif(new.raw_user_meta_data ->> 'full_name', ''),
    public.compute_friend_id(new.id)
  )
  on conflict (id) do update set
    email = coalesce(public.profiles.email, excluded.email),
    full_name = coalesce(public.profiles.full_name, excluded.full_name),
    friend_id = coalesce(nullif(public.profiles.friend_id, ''), excluded.friend_id),
    updated_at = now();
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

drop function if exists public.find_profile_by_friend_id(text);
create function public.find_profile_by_friend_id(p_friend_id text)
returns table (
  id uuid,
  full_name text,
  email text,
  avatar_url text,
  friend_id text
)
language sql
security definer
stable
set search_path = public
as $$
  select p.id, p.full_name, p.email, p.avatar_url, p.friend_id
  from public.profiles p
  where p.friend_id = p_friend_id
  limit 1;
$$;

grant execute on function public.find_profile_by_friend_id(text) to authenticated;