-- =====================================================================
-- Transferências entre contas
-- =====================================================================
-- Adiciona 'transfer' como tipo de transação (não é receita nem despesa,
-- apenas move saldo entre duas contas do usuário) e a coluna que guarda
-- a conta de destino, espelhando account_id/account (origem).
-- =====================================================================

alter type public.transaction_type add value if not exists 'transfer';

alter table public.transactions
  add column if not exists to_account_id uuid references public.accounts (id) on delete set null,
  add column if not exists to_account text;
