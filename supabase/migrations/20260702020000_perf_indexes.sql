-- =====================================================================
-- Otimização de performance — índices complementares
-- =====================================================================
-- As tabelas de domínio (accounts, categories, transactions, goals,
-- goal_movements, friendships, profiles.friend_id, shopping_lists.user_id)
-- JÁ possuem índices nas migrations anteriores. Este arquivo só fecha as
-- lacunas restantes. Todos os comandos são idempotentes (IF NOT EXISTS).
--
-- Rodar no SQL Editor do Supabase (o deploy da Vercel não aplica migrations).
-- =====================================================================

-- shopping_lists: a policy de SELECT usa `auth.uid() = ANY(participant_ids)`.
-- Um btree não cobre esse operador de array; um GIN sim.
create index if not exists shopping_lists_participants_gin
  on public.shopping_lists using gin (participant_ids);

-- goal_movements: consultas por usuário ordenadas/filtradas por data.
create index if not exists idx_goal_movements_user_date
  on public.goal_movements (user_id, date);

-- transactions: reforço para filtros por conta (account_id) e categoria.
create index if not exists idx_transactions_user_account
  on public.transactions (user_id, account_id);
create index if not exists idx_transactions_user_category
  on public.transactions (user_id, category_id);

-- Atualiza estatísticas do planejador para as tabelas indexadas.
analyze public.shopping_lists;
analyze public.goal_movements;
analyze public.transactions;
