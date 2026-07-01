// Camada de acesso a dados via API server-side para o Caesar Finance.
// Converte entre os tipos camelCase do app e as colunas snake_case do banco,
// carrega tudo de uma vez e sincroniza cada coleção por diferença (upsert/delete).
import type { Account, CustomCategory, Transaction, Goal, GoalMovement } from '../App';

// Aceita apenas UUID válido; qualquer id legado (ex.: 'wallet') vira null para
// não violar as chaves estrangeiras account_id / category_id.
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const asUuid = (value?: string | null): string | null => (value && UUID_RE.test(value) ? value : null);

type Row = Record<string, unknown>;
const str = (v: unknown) => (v == null ? undefined : String(v));
const num = (v: unknown) => (v == null ? undefined : Number(v));

// ----------------------------- Mappers ------------------------------------
function accountToRow(a: Account): Row {
  return {
    id: a.id,
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

function categoryToRow(c: CustomCategory): Row {
  return {
    id: c.id,
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

function transactionToRow(t: Transaction): Row {
  return {
    id: t.id,
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

function goalToRow(g: Goal): Row {
  return {
    id: g.id,
    name: g.name,
    target_amount: g.targetAmount,
    deadline: g.deadline ?? null,
    color: g.color,
    icon: g.icon,
    created_at: g.createdAt,
  };
}
function movementToRow(m: GoalMovement, goalId: string): Row {
  return {
    id: m.id,
    goal_id: goalId,
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
  profileName?: string;
  profilePhone?: string;
  profileAvatarUrl?: string | null;
  accounts: Account[];
  categories: CustomCategory[];
  transactions: Transaction[];
  goals: Goal[];
};

type ApiData = {
  profile: { full_name?: unknown; email?: unknown; phone?: unknown; avatar_url?: unknown } | null;
  accounts: Row[];
  categories: Row[];
  transactions: Row[];
  goals: Row[];
  goal_movements: Row[];
};

async function apiRequest<T>(method: 'GET' | 'POST', body?: Record<string, unknown>): Promise<T> {
  const response = await fetch('/api/data', {
    method,
    credentials: 'same-origin',
    headers: body ? { 'Content-Type': 'application/json' } : undefined,
    body: body ? JSON.stringify(body) : undefined,
  });
  const payload = await response.json().catch(() => ({})) as { error?: string } & T;
  if (!response.ok) throw new Error(payload.error || 'Não foi possível acessar os dados.');
  return payload;
}

export async function loadAll(): Promise<LoadedData> {
  const data = await apiRequest<ApiData>('GET');

  const movementsByGoal = new Map<string, GoalMovement[]>();
  data.goal_movements.forEach((row) => {
    const goalId = String(row.goal_id);
    const list = movementsByGoal.get(goalId) ?? [];
    list.push(rowToMovement(row));
    movementsByGoal.set(goalId, list);
  });

  return {
    profileName: str(data.profile?.full_name),
    profilePhone: str(data.profile?.phone),
    profileAvatarUrl: str(data.profile?.avatar_url) ?? null,
    accounts: data.accounts.map(rowToAccount),
    categories: data.categories.map(rowToCategory),
    transactions: data.transactions.map(rowToTransaction),
    goals: data.goals.map((row) => rowToGoal(row, movementsByGoal.get(String(row.id)) ?? [])),
  };
}

export async function updateProfile(profile: { fullName: string; phone: string; avatarUrl?: string | null }) {
  const data = await apiRequest<{ profile: { full_name?: unknown; phone?: unknown; avatar_url?: unknown } }>('POST', {
    resource: 'profile',
    profile,
  });
  return {
    fullName: str(data.profile.full_name) ?? profile.fullName,
    phone: str(data.profile.phone) ?? '',
    avatarUrl: str(data.profile.avatar_url) ?? null,
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

export async function syncAccounts(_userId: string, prev: Account[], next: Account[]) {
  const { toUpsert, toDelete } = diff(prev, next);
  if (!toDelete.length && !toUpsert.length) return;
  await apiRequest('POST', { resource: 'accounts', deleteIds: toDelete, upsert: toUpsert.map(accountToRow) });
}

export async function syncCategories(_userId: string, prev: CustomCategory[], next: CustomCategory[]) {
  const { toUpsert, toDelete } = diff(prev, next);
  if (!toDelete.length && !toUpsert.length) return;
  await apiRequest('POST', { resource: 'categories', deleteIds: toDelete, upsert: toUpsert.map(categoryToRow) });
}

export async function syncTransactions(_userId: string, prev: Transaction[], next: Transaction[]) {
  const { toUpsert, toDelete } = diff(prev, next);
  if (!toDelete.length && !toUpsert.length) return;
  await apiRequest('POST', { resource: 'transactions', deleteIds: toDelete, upsert: toUpsert.map(transactionToRow) });
}

export async function searchUserByFriendId(friendId: string): Promise<{ id: string; publicFriendId: string; name: string; email: string; avatarUrl: string | null } | null> {
  const res = await fetch(`/api/users?friendId=${encodeURIComponent(friendId)}`, { credentials: 'same-origin' });
  if (!res.ok) return null;
  const data = await res.json() as { user: { id: string; publicFriendId: string; name: string; email: string; avatarUrl: string | null } | null };
  return data.user ?? null;
}

export async function syncGoals(_userId: string, prev: Goal[], next: Goal[]) {
  const { toUpsert, toDelete } = diff(prev, next);
  if (!toDelete.length && !toUpsert.length) return;
  await apiRequest('POST', {
    resource: 'goals',
    deleteIds: toDelete,
    upsert: toUpsert.map(goalToRow),
    replaceMovements: toUpsert.map((goal) => ({
      goalId: goal.id,
      rows: goal.movements.map((movement) => movementToRow(movement, goal.id)),
    })),
  });
}
