-- Permite que usuários logados leiam perfis públicos (necessário para buscar amigos por ID e exibir nomes em listas/transações compartilhadas)
drop policy if exists "profiles_select_authenticated" on public.profiles;

create policy "profiles_select_authenticated" on public.profiles
  for select to authenticated using (true);
