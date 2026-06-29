// Camada de acesso a dados (Supabase) para o Caesar Finance.
// Converte entre os tipos camelCase do app e as colunas snake_case do banco,
// carrega tudo de uma vez e sincroniza cada coleção por diferença (upsert/delete).
import { supabase } from './supabase';
import type { Account, CustomCategory, Transaction, Goal, GoalMovement } from '../App';

// Aceita apenas UUID válido; qualquer id legado (ex.: 'wallet') vira null para
// não violar as chaves estrangeiras account_id / category_id.
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const asUuid = (value?: string | null): string | null => (value && UUID_RE.test(value) ? value : null);

type Row = Record<string, unknown>;
const str = (v: unknown) => (v == null ? undefined : String(v));
const num = (v: unknown) => (v == null ? undefined : Number(v));

// ----------------------------- Mappers ------------------------------------
function accountToRow(a: Account, userId: string): Row {
  return {
    id: a.id,
    user_id: userId,
    name: a.name,
    type: a.type,
    initial_balance: a.initialBalance,
    color: a.color ?? null,
    icon: a.icon ?? null,
  };
}
function rowToAccount(r: Row): Account {
  return {
    id: String(r.id),
    name: String(r.name),
    type: r.type as Account['type'],
    initialBalance: Number(r.initial_balance ?? 0),
    color: str(r.color),
    icon: r.icon as Account['icon'],
  };
}

function categoryToRow(c: CustomCategory, userId: string): Row {
  return {
    id: c.id,
    user_id: userId,
    name: c.name,
    kind: c.kind,
    color: c.color ?? null,
    icon: c.icon ?? null,
  };
}
function rowToCategory(r: Row): CustomCategory {
  return {
    id: String(r.id),
    name: String(r.name),
    kind: r.kind as CustomCategory['kind'],
    color: str(r.color),
    icon: r.icon as CustomCategory['icon'],
  };
}

function transactionToRow(t: Transaction, userId: string): Row {
  return {
    id: t.id,
    user_id: userId,
    group_id: t.groupId,
    description: t.description,
    category: t.category,
    category_id: null,
    amount: t.amount,
    type: t.type,
    due_date: t.dueDate,
    recurrence: t.recurrence,
    installment_number: t.installmentNumber ?? null,
    installment_total: t.installmentTotal ?? null,
    status: t.status,
    settled_at: t.settledAt ?? null,
    settled_amount: t.settledAmount ?? null,
    account_id: asUuid(t.accountId),
    account: t.account ?? null,
    notes: t.notes ?? null,
  };
}
function rowToTransaction(r: Row): Transaction {
  return {
    id: String(r.id),
    groupId: String(r.group_id),
    description: String(r.description),
    category: String(r.category),
    amount: Number(r.amount ?? 0),
    type: r.type as Transaction['type'],
    dueDate: String(r.due_date),
    recurrence: r.recurrence as Transaction['recurrence'],
    installmentNumber: num(r.installment_number),
    installmentTotal: num(r.installment_total),
    status: r.status as Transaction['status'],
    settledAt: str(r.settled_at),
    settledAmount: num(r.settled_amount),
    accountId: str(r.account_id),
    account: str(r.account),
    notes: str(r.notes),
  };
}

function goalToRow(g: Goal, userId: string): Row {
  return {
    id: g.id,
    user_id: userId,
    name: g.name,
    target_amount: g.targetAmount,
    deadline: g.deadline ?? null,
    color: g.color,
    icon: g.icon,
    created_at: g.createdAt,
  };
}
function movementToRow(m: GoalMovement, goalId: string, userId: string): Row {
  return {
    id: m.id,
    goal_id: goalId,
    user_id: userId,
    type: m.type,
    amount: m.amount,
    date: m.date,
    note: m.note ?? null,
  };
}
function rowToMovement(r: Row): GoalMovement {
  return {
    id: String(r.id),
    type: r.type as GoalMovement['type'],
    amount: Number(r.amount ?? 0),
    date: String(r.date),
    note: str(r.note),
  };
}
function rowToGoal(r: Row, movements: GoalMovement[]): Goal {
  const created = String(r.created_at ?? '');
  return {
    id: String(r.id),
    name: String(r.name),
    targetAmount: Number(r.target_amount ?? 0),
    deadline: str(r.deadline),
    color: String(r.color ?? '#1B99D8'),
    icon: r.icon as Goal['icon'],
    createdAt: created.slice(0, 10) || new Date().toISOString().slice(0, 10),
    movements,
  };
}

// ------------------------------ Load --------------------------------------
export type LoadedData = {
  accounts: Account[];
  categories: CustomCategory[];
  transactions: Transaction[];
  goals: Goal[];
};

export async function loadAll(): Promise<LoadedData> {
  const [accountsRes, categoriesRes, txRes, goalsRes, movementsRes] = await Promise.all([
    supabase.from('accounts').select('*').order('created_at', { ascending: true }),
    supabase.from('categories').select('*').order('created_at', { ascending: true }),
    supabase.from('transactions').select('*').order('due_date', { ascending: true }),
    supabase.from('goals').select('*').order('created_at', { ascending: true }),
    supabase.from('goal_movements').select('*').order('date', { ascending: true }),
  ]);

  const error = accountsRes.error || categoriesRes.error || txRes.error || goalsRes.error || movementsRes.error;
  if (error) throw error;

  const movementsByGoal = new Map<string, GoalMovement[]>();
  (movementsRes.data ?? []).forEach((row) => {
    const goalId = String((row as Row).goal_id);
    const list = movementsByGoal.get(goalId) ?? [];
    list.push(rowToMovement(row as Row));
    movementsByGoal.set(goalId, list);
  });

  return {
    accounts: (accountsRes.data ?? []).map((r) => rowToAccount(r as Row)),
    categories: (categoriesRes.data ?? []).map((r) => rowToCategory(r as Row)),
    transactions: (txRes.data ?? []).map((r) => rowToTransaction(r as Row)),
    goals: (goalsRes.data ?? []).map((r) => rowToGoal(r as Row, movementsByGoal.get(String((r as Row).id)) ?? [])),
  };
}

// ------------------------------ Sync --------------------------------------
// Aproveita a imutabilidade do app: itens inalterados mantêm a mesma
// referência entre `prev` e `next`, então só sincronizamos o que mudou.
function diff<T extends { id: string }>(prev: T[], next: T[]) {
  const prevById = new Map(prev.map((item) => [item.id, item] as const));
  const nextIds = new Set(next.map((item) => item.id));
  const toUpsert = next.filter((item) => prevById.get(item.id) !== item);
  const toDelete = prev.filter((item) => !nextIds.has(item.id)).map((item) => item.id);
  return { toUpsert, toDelete };
}

export async function syncAccounts(userId: string, prev: Account[], next: Account[]) {
  const { toUpsert, toDelete } = diff(prev, next);
  if (toDelete.length) {
    const { error } = await supabase.from('accounts').delete().in('id', toDelete);
    if (error) throw error;
  }
  if (toUpsert.length) {
    const { error } = await supabase.from('accounts').upsert(toUpsert.map((a) => accountToRow(a, userId)));
    if (error) throw error;
  }
}

export async function syncCategories(userId: string, prev: CustomCategory[], next: CustomCategory[]) {
  const { toUpsert, toDelete } = diff(prev, next);
  if (toDelete.length) {
    const { error } = await supabase.from('categories').delete().in('id', toDelete);
    if (error) throw error;
  }
  if (toUpsert.length) {
    const { error } = await supabase.from('categories').upsert(toUpsert.map((c) => categoryToRow(c, userId)));
    if (error) throw error;
  }
}

export async function syncTransactions(userId: string, prev: Transaction[], next: Transaction[]) {
  const { toUpsert, toDelete } = diff(prev, next);
  if (toDelete.length) {
    const { error } = await supabase.from('transactions').delete().in('id', toDelete);
    if (error) throw error;
  }
  if (toUpsert.length) {
    const { error } = await supabase.from('transactions').upsert(toUpsert.map((t) => transactionToRow(t, userId)));
    if (error) throw error;
  }
}

export async function syncGoals(userId: string, prev: Goal[], next: Goal[]) {
  const { toUpsert, toDelete } = diff(prev, next);
  if (toDelete.length) {
    const { error } = await supabase.from('goals').delete().in('id', toDelete);
    if (error) throw error;
  }
  for (const goal of toUpsert) {
    const { error: goalError } = await supabase.from('goals').upsert(goalToRow(goal, userId));
    if (goalError) throw goalError;
    // Movimentos são poucos por meta: substituímos o conjunto inteiro.
    const { error: delError } = await supabase.from('goal_movements').delete().eq('goal_id', goal.id);
    if (delError) throw delError;
    if (goal.movements.length) {
      const { error: insError } = await supabase
        .from('goal_movements')
        .insert(goal.movements.map((m) => movementToRow(m, goal.id, userId)));
      if (insError) throw insError;
    }
  }
}
