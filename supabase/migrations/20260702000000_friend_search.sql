-- =====================================================================
-- Friend search: adiciona coluna friend_id e libera SELECT de perfis
-- =====================================================================

-- 1. Adicionar coluna friend_id na tabela profiles (se não existir)
alter table public.profiles
  add column if not exists friend_id text;

-- 2. Índice único para busca rápida por friend_id
create unique index if not exists profiles_friend_id_idx
  on public.profiles (friend_id)
  where friend_id is not null;

-- 3. Substituir política restrita (só o próprio) por política que permite
--    qualquer usuário logado ler perfis — necessário para busca de amigos
drop policy if exists "profiles_select_own" on public.profiles;
drop policy if exists "profiles_select_authenticated" on public.profiles;

create policy "profiles_select_authenticated" on public.profiles
  for select to authenticated using (true);

-- 4. Manter update restrito ao próprio perfil
drop policy if exists "profiles_update_own" on public.profiles;
create policy "profiles_update_own" on public.profiles
  for update to authenticated
  using (auth.uid() = id)
  with check (auth.uid() = id);

-- 5. Função auxiliar para busca por friend_id (usa SECURITY DEFINER para
--    garantir acesso mesmo que a política mude no futuro)
create or replace function public.find_profile_by_friend_id(p_friend_id text)
returns table (
  id         uuid,
  full_name  text,
  avatar_url text,
  friend_id  text
)
language sql
security definer
stable
set search_path = public
as $$
  select p.id, p.full_name, p.avatar_url, p.friend_id
  from public.profiles p
  where p.friend_id = p_friend_id
  limit 1;
$$;

-- Permite que usuários logados chamem a função
grant execute on function public.find_profile_by_friend_id(text) to authenticated;
