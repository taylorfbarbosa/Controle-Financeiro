-- =====================================================================
-- Caesar Finance — Estrutura completa do backend (Supabase / PostgreSQL)
-- =====================================================================
-- Cobre todas as abas do app: Contas, Categorias, Transações, Metas e Perfil.
-- Princípios:
--   1. Cada usuário só enxerga os próprios dados (RLS por auth.uid()).
--   2. Ao criar conta NENHUM dado é semeado — sem carteira, categoria ou
--      transação padrão. O usuário começa com tudo zerado.
--   3. Script idempotente: pode rodar novamente sem quebrar.
--
-- ATENÇÃO: este script recria as tabelas do domínio. As tabelas atuais
-- estão vazias (0 registros), então o DROP é seguro. Se um dia tiver dados,
-- faça backup antes.
-- =====================================================================

-- ---------------------------------------------------------------------
-- 0. Extensões
-- ---------------------------------------------------------------------
create extension if not exists "pgcrypto"; -- gen_random_uuid()

-- ---------------------------------------------------------------------
-- 1. Tipos (enums) do domínio
-- ---------------------------------------------------------------------
do $$ begin
  create type public.account_type   as enum ('wallet', 'checking', 'savings', 'investment');
exception when duplicate_object then null; end $$;

do $$ begin
  create type public.category_kind  as enum ('income', 'expense', 'transfer');
exception when duplicate_object then null; end $$;

do $$ begin
  create type public.transaction_type as enum ('income', 'expense');
exception when duplicate_object then null; end $$;

do $$ begin
  create type public.recurrence_type as enum ('single', 'fixed', 'installment');
exception when duplicate_object then null; end $$;

do $$ begin
  create type public.transaction_status as enum ('open', 'settled');
exception when duplicate_object then null; end $$;

do $$ begin
  create type public.goal_movement_type as enum ('deposit', 'withdraw');
exception when duplicate_object then null; end $$;

-- ---------------------------------------------------------------------
-- 2. Função utilitária: manter updated_at
-- ---------------------------------------------------------------------
create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at := now();
  return new;
end;
$$;

-- ---------------------------------------------------------------------
-- 3. Limpeza (ordem reversa por causa das FKs)
-- ---------------------------------------------------------------------
drop table if exists public.goal_movements cascade;
drop table if exists public.goals          cascade;
drop table if exists public.transactions   cascade;
drop table if exists public.categories     cascade;
drop table if exists public.accounts       cascade;
drop table if exists public.profiles       cascade;

-- =====================================================================
-- 4. PROFILES — 1:1 com auth.users
-- =====================================================================
create table public.profiles (
  id          uuid primary key references auth.users (id) on delete cascade,
  email       text,
  full_name   text,
  avatar_url  text,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);

create trigger trg_profiles_updated_at
  before update on public.profiles
  for each row execute function public.set_updated_at();

-- =====================================================================
-- 5. ACCOUNTS (aba Contas) — carteira, conta corrente, poupança, investimento
-- =====================================================================
create table public.accounts (
  id              uuid primary key default gen_random_uuid(),
  user_id         uuid not null references auth.users (id) on delete cascade,
  name            text not null,
  type            public.account_type not null default 'wallet',
  initial_balance numeric(14,2) not null default 0,
  color           text,
  icon            text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create index idx_accounts_user on public.accounts (user_id);

create trigger trg_accounts_updated_at
  before update on public.accounts
  for each row execute function public.set_updated_at();

-- =====================================================================
-- 6. CATEGORIES (aba Categorias) — receita, despesa, transferência
-- =====================================================================
create table public.categories (
  id          uuid primary key default gen_random_uuid(),
  user_id     uuid not null references auth.users (id) on delete cascade,
  name        text not null,
  kind        public.category_kind not null,
  color       text,
  icon        text,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now(),
  unique (user_id, kind, name)
);

create index idx_categories_user on public.categories (user_id);

create trigger trg_categories_updated_at
  before update on public.categories
  for each row execute function public.set_updated_at();

-- =====================================================================
-- 7. TRANSACTIONS (aba Transações) — núcleo do fluxo de caixa
-- =====================================================================
-- group_id agrupa lançamentos de uma mesma recorrência/parcelamento.
-- category/account ficam denormalizados (label) + FK opcional, espelhando
-- o app (categorias e contas embutidas nem sempre são linha em tabela).
create table public.transactions (
  id                 uuid primary key default gen_random_uuid(),
  user_id            uuid not null references auth.users (id) on delete cascade,
  group_id           uuid not null default gen_random_uuid(),
  description        text not null,
  category           text not null,
  category_id        uuid references public.categories (id) on delete set null,
  amount             numeric(14,2) not null check (amount >= 0),
  type               public.transaction_type not null,
  due_date           date not null,
  recurrence         public.recurrence_type not null default 'single',
  installment_number integer check (installment_number is null or installment_number > 0),
  installment_total  integer check (installment_total  is null or installment_total  > 0),
  status             public.transaction_status not null default 'open',
  settled_at         date,
  settled_amount     numeric(14,2),
  account_id         uuid references public.accounts (id) on delete set null,
  account            text,
  notes              text,
  created_at         timestamptz not null default now(),
  updated_at         timestamptz not null default now()
);

create index idx_transactions_user        on public.transactions (user_id);
create index idx_transactions_user_due     on public.transactions (user_id, due_date);
create index idx_transactions_group        on public.transactions (group_id);
create index idx_transactions_user_status  on public.transactions (user_id, status);

create trigger trg_transactions_updated_at
  before update on public.transactions
  for each row execute function public.set_updated_at();

-- =====================================================================
-- 8. GOALS + GOAL_MOVEMENTS (aba Metas)
-- =====================================================================
create table public.goals (
  id            uuid primary key default gen_random_uuid(),
  user_id       uuid not null references auth.users (id) on delete cascade,
  name          text not null,
  target_amount numeric(14,2) not null check (target_amount > 0),
  deadline      date,
  color         text,
  icon          text,
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now()
);

create index idx_goals_user on public.goals (user_id);

create trigger trg_goals_updated_at
  before update on public.goals
  for each row execute function public.set_updated_at();

create table public.goal_movements (
  id          uuid primary key default gen_random_uuid(),
  goal_id     uuid not null references public.goals (id) on delete cascade,
  user_id     uuid not null references auth.users (id) on delete cascade,
  type        public.goal_movement_type not null,
  amount      numeric(14,2) not null check (amount > 0),
  date        date not null default current_date,
  note        text,
  created_at  timestamptz not null default now()
);

create index idx_goal_movements_goal on public.goal_movements (goal_id);
create index idx_goal_movements_user on public.goal_movements (user_id);

-- =====================================================================
-- 9. Row Level Security — cada usuário só acessa o que é seu
-- =====================================================================
alter table public.profiles       enable row level security;
alter table public.accounts       enable row level security;
alter table public.categories     enable row level security;
alter table public.transactions   enable row level security;
alter table public.goals          enable row level security;
alter table public.goal_movements enable row level security;

-- profiles: o dono lê/edita o próprio perfil
create policy "profiles_select_own" on public.profiles
  for select using (auth.uid() = id);
create policy "profiles_update_own" on public.profiles
  for update using (auth.uid() = id) with check (auth.uid() = id);

-- accounts
create policy "accounts_select_own" on public.accounts
  for select using (auth.uid() = user_id);
create policy "accounts_insert_own" on public.accounts
  for insert with check (auth.uid() = user_id);
create policy "accounts_update_own" on public.accounts
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "accounts_delete_own" on public.accounts
  for delete using (auth.uid() = user_id);

-- categories
create policy "categories_select_own" on public.categories
  for select using (auth.uid() = user_id);
create policy "categories_insert_own" on public.categories
  for insert with check (auth.uid() = user_id);
create policy "categories_update_own" on public.categories
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "categories_delete_own" on public.categories
  for delete using (auth.uid() = user_id);

-- transactions
create policy "transactions_select_own" on public.transactions
  for select using (auth.uid() = user_id);
create policy "transactions_insert_own" on public.transactions
  for insert with check (auth.uid() = user_id);
create policy "transactions_update_own" on public.transactions
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "transactions_delete_own" on public.transactions
  for delete using (auth.uid() = user_id);

-- goals
create policy "goals_select_own" on public.goals
  for select using (auth.uid() = user_id);
create policy "goals_insert_own" on public.goals
  for insert with check (auth.uid() = user_id);
create policy "goals_update_own" on public.goals
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "goals_delete_own" on public.goals
  for delete using (auth.uid() = user_id);

-- goal_movements
create policy "goal_movements_select_own" on public.goal_movements
  for select using (auth.uid() = user_id);
create policy "goal_movements_insert_own" on public.goal_movements
  for insert with check (auth.uid() = user_id);
create policy "goal_movements_update_own" on public.goal_movements
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "goal_movements_delete_own" on public.goal_movements
  for delete using (auth.uid() = user_id);

-- =====================================================================
-- 10. Cadastro de novo usuário — cria SOMENTE o profile (sem dados de exemplo)
-- =====================================================================
-- Esta é a parte central do pedido: ao registrar, o usuário NÃO recebe
-- carteira, categoria nem transação inicial. Apenas o perfil é criado.
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email, full_name)
  values (
    new.id,
    new.email,
    nullif(new.raw_user_meta_data ->> 'full_name', '')
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- =====================================================================
-- FIM — estrutura criada, RLS ativo, conta nova começa zerada.
-- =====================================================================
