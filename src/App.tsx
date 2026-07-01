import { Fragment, memo, useEffect, useMemo, useRef, useState, useTransition, type ChangeEvent, type FormEvent, type ReactNode, type RefObject } from 'react';
import { createPortal } from 'react-dom';
import { readSheet, type Row } from 'read-excel-file/browser';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import writeXlsxFile, { type Cell, type SheetData } from 'write-excel-file/browser';
import { Area, AreaChart, Bar, BarChart, CartesianGrid, LabelList, Legend, Line, LineChart, ReferenceLine, ResponsiveContainer, Tooltip as RechartsTooltip, XAxis } from 'recharts';
import type { Session } from './lib/auth';
import rubyLogoWhite from './assets/rubylife-white.png';
import rubyLogoColor from './assets/rubylife-color.png';
import rubyDiamond from './assets/Diamante.png';
import { authClient as supabase } from './lib/auth';
import { loadAll, searchUserByFriendId, syncAccounts, syncCategories, syncGoals, syncTransactions, updateProfile } from './lib/db';
import { DEFAULT_PUSH_PREFERENCES, PUSH_PREFERENCE_LABELS, enablePushNotifications, getPushStatus, savePushPreferences, sendPushNotification, type PushPreferences, type PushStatus } from './lib/push';
import { ShoppingListsPage } from './ShoppingListsFeature';
import {
  AlertCircle,
  BadgeDollarSign,

  ArrowLeftRight,
  ArrowDownLeft,
  ArrowUpRight,
  Banknote,
  Bell,
  BookOpen,
  BriefcaseBusiness,
  CalendarDays,
  Camera,
  CheckCircle2,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CircleDollarSign,
  Car,
  Check,
  Clock3,
  Coins,
  Copy,
  CreditCard,
  Eye,
  EyeOff,
  FileText,
  FileSpreadsheet,
  Gamepad2,
  Gift,
  GraduationCap,
  HandHeart,
  HeartPulse,
  Home,
  HelpCircle,
  House,
  LayoutGrid,
  ListFilter,
  Landmark,
  LogOut,
  Mail,
  Search,
  Minus,
  MoreVertical,
  MoreHorizontal,
  Moon,
  Music,
  Dumbbell,
  Plane,
  Sun,
  PiggyBank,
  Phone,
  PawPrint,
  Pencil,
  Plus,
  ReceiptText,
  RotateCcw,
  Repeat2,
  Save,
  Share,
  ShoppingBag,
  ShoppingCart,
  Target,
  Tags,
  TrendingDown,
  TrendingUp,
  Trash2,
  Utensils,
  Users,
  Upload,
  UserPlus,
  UserRound,
  UserX,
  WalletCards,
  Wrench,
  X,
} from 'lucide-react';

type TransactionType = 'income' | 'expense';
type AppPage = 'dashboard' | 'transactions' | 'categories' | 'goals' | 'reports' | 'accounts' | 'shopping' | 'friends' | 'profile' | 'help' | 'notifications';
const PAGE_LABELS: Record<AppPage, string> = {
  dashboard: 'Visão Geral',
  transactions: 'Transações',
  categories: 'Categorias',
  goals: 'Metas',
  reports: 'Relatórios',
  accounts: 'Contas',
  shopping: 'Listas de compras',
  friends: 'Amigos',
  profile: 'Meu perfil',
  help: 'Ajuda',
  notifications: 'Notificações',
};
export type FriendshipInvitationStatus = 'pending' | 'accepted' | 'declined' | 'cancelled';
export type SharedTransactionStatus = 'pending' | 'approved' | 'declined' | 'cancelled';
type GoalMovementType = 'deposit' | 'withdraw';
type RecurrenceType = 'single' | 'fixed' | 'installment';
type TransactionStatus = 'open' | 'settled';
type TransactionTypeFilter = 'all' | TransactionType;
type TransactionSummaryMode = 'income' | 'expense' | 'balance';
type AccountType = 'wallet' | 'checking' | 'savings' | 'investment';
type AccountTypeFilter = 'all' | AccountType;
type AccountFilters = {
  search: string;
  type: AccountTypeFilter;
};
type CategoryFilters = {
  search: string;
  type: 'all' | CategoryKind;
};
type GoalFilters = {
  search: string;
  status: 'all' | 'active' | 'reached';
};
type AccountIconKey = 'wallet' | 'card' | 'bank' | 'savings' | 'investment' | 'cash';
type CategoryKind = 'income' | 'expense' | 'transfer';
type CategoryIconKey = 'salary' | 'money' | 'work' | 'shopping' | 'refund' | 'investment' | 'home' | 'card' | 'gift' | 'food' | 'car' | 'health' | 'education' | 'leisure' | 'repeat' | 'family' | 'pets' | 'donation' | 'wallet' | 'transfer' | 'tag' | 'travel' | 'fitness' | 'music' | 'maintenance';

export type Account = {
  id: string;
  name: string;
  type: AccountType;
  initialBalance: number;
  color?: string;
  icon?: AccountIconKey;
};

export type CustomCategory = {
  id: string;
  name: string;
  kind: CategoryKind;
  color?: string;
  icon?: CategoryIconKey;
  deleted?: boolean;
};

type CategoryItem = {
  id: string;
  name: string;
  kind: CategoryKind;
  color: string;
  icon: CategoryIconKey;
};

type ReportFilters = {
  date: string;
  category: string;
  type: TransactionTypeFilter;
};

export type Transaction = {
  id: string;
  groupId: string;
  description: string;
  category: string;
  amount: number;
  type: TransactionType;
  dueDate: string;
  recurrence: RecurrenceType;
  installmentNumber?: number;
  installmentTotal?: number;
  status: TransactionStatus;
  settledAt?: string;
  settledAmount?: number;
  accountId?: string;
  account?: string;
  notes?: string;
  sharedRequestId?: string;
  sharedCreatedBy?: string;
  sharedCreatedByName?: string;
};

export type GoalMovement = {
  id: string;
  type: GoalMovementType;
  amount: number;
  date: string;
  note?: string;
};

export type Goal = {
  id: string;
  name: string;
  targetAmount: number;
  deadline?: string;
  color: string;
  icon: CategoryIconKey;
  createdAt: string;
  movements: GoalMovement[];
};

export type FriendUser = {
  id: string;
  publicFriendId: string;
  name: string;
  email: string;
  avatarUrl?: string | null;
};

export type FriendshipInvitation = {
  id: string;
  requesterId: string;
  receiverId: string;
  status: FriendshipInvitationStatus;
  createdAt: string;
  respondedAt?: string;
};

export type SharedTransactionRequest = {
  id: string;
  creatorId: string;
  receiverId: string;
  type: TransactionType;
  description: string;
  amount: number;
  dueDate: string;
  category: string;
  note?: string;
  declineReason?: string;
  status: SharedTransactionStatus;
  createdAt: string;
  respondedAt?: string;
  approvedTransactionId?: string;
};

type SharedTransactionForm = {
  type: TransactionType;
  friendId: string;
  description: string;
  amount: string;
  dueDate: string;
  category: string;
  note: string;
};

type AppNotification = {
  id: string;
  userId: string;
  type: 'friend_invite' | 'friend_accepted' | 'shared_created' | 'shared_approved' | 'shared_declined' | 'shared_cancelled';
  message: string;
  createdAt: string;
  read?: boolean;
};


type BalanceDotProps = {
  cx?: number | string;
  cy?: number | string;
  payload?: { Saldo?: number };
};

type BalanceLabelProps = {
  x?: number | string;
  y?: number | string;
  value?: unknown;
};

const TRANSACTION_DESCRIPTION_MAX_LENGTH = 24;

type LaunchForm = {
  description: string;
  category: string;
  accountId: string;
  amount: string;
  dueDate: string;
  recurrence: RecurrenceType;
  installments: string;
  fixedUntil: string;
};

type ImportPreviewRow = {
  rowNumber: number;
  type: TransactionType | null;
  description: string;
  amount: number;
  dueDate: string;
  category: string;
  transactions: Transaction[];
  error?: string;
};

function normalizePublicFriendId(value: string) {
  return value.trim().replace(/^#/, '');
}

function publicFriendIdForUser(userId: string) {
  const uuidHex = userId.replace(/-/g, '');
  if (/^[0-9a-f]{32}$/i.test(uuidHex)) return (BigInt(`0x${uuidHex}`) % 1000000n).toString().padStart(6, '0');

  const encoded = new TextEncoder().encode(userId || '0');
  const hex = Array.from(encoded, (byte) => byte.toString(16).padStart(2, '0')).join('');
  return (BigInt(`0x${hex || '0'}`) % 1000000n).toString().padStart(6, '0');
}

function friendshipPairKey(a: string, b: string) {
  return [a, b].sort().join(':');
}

function invitationConnectsUsers(invitation: FriendshipInvitation, a: string, b: string) {
  return friendshipPairKey(invitation.requesterId, invitation.receiverId) === friendshipPairKey(a, b);
}

export function areUsersFriends(currentUserId: string, targetUserId: string, invitations: FriendshipInvitation[]) {
  return invitations.some((invitation) => invitation.status === 'accepted' && invitationConnectsUsers(invitation, currentUserId, targetUserId));
}

export function canShareWithUser(currentUserId: string, targetUserId: string, invitations: FriendshipInvitation[]) {
  return areUsersFriends(currentUserId, targetUserId, invitations);
}

function findBlockingFriendInvitation(currentUserId: string, target: FriendUser, invitations: FriendshipInvitation[]) {
  if (!currentUserId || currentUserId === target.id) return 'Você não pode enviar convite para si mesmo.';
  if (areUsersFriends(currentUserId, target.id, invitations)) return 'Este usuário já está conectado como amigo.';
  const pending = invitations.find((invitation) => invitation.status === 'pending' && invitationConnectsUsers(invitation, currentUserId, target.id));
  if (pending) return 'Já existe um convite pendente entre vocês.';
  return null;
}

const PUBLIC_USERS_STORAGE_KEY = 'rubylife-public-users';
const FRIEND_INVITATIONS_STORAGE_KEY = 'rubylife-friend-invitations';
const APP_NOTIFICATIONS_STORAGE_KEY = 'rubylife-notifications';

function loadPublicUsers(): FriendUser[] {
  try {
    const raw = localStorage.getItem(PUBLIC_USERS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as FriendUser[];
    if (!Array.isArray(parsed)) return [];
    const migrated = parsed
      .filter((user) => user.id)
      .map((user) => ({ ...user, publicFriendId: publicFriendIdForUser(user.id) }));
    return [...new Map(migrated.map((user) => [user.id, user])).values()];
  } catch {
    return [];
  }
}

function storePublicUsers(users: FriendUser[]) {
  try {
    localStorage.setItem(PUBLIC_USERS_STORAGE_KEY, JSON.stringify(users));
  } catch {
    // armazenamento local indisponível
  }
}

function upsertPublicUser(user: FriendUser) {
  const users = loadPublicUsers();
  const next = [user, ...users.filter((item) => item.id !== user.id && normalizePublicFriendId(item.publicFriendId) !== normalizePublicFriendId(user.publicFriendId))];
  storePublicUsers(next);
  return next;
}

function loadStoredFriendInvitations(currentUserId?: string): FriendshipInvitation[] {
  try {
    const raw = localStorage.getItem(FRIEND_INVITATIONS_STORAGE_KEY);
    const globalInvitations = raw ? JSON.parse(raw) as FriendshipInvitation[] : [];
    if (Array.isArray(globalInvitations) && globalInvitations.length) return globalInvitations;
    if (!currentUserId) return [];
    const legacy = localStorage.getItem(`rubylife-friend-invitations-${currentUserId}`);
    const parsedLegacy = legacy ? JSON.parse(legacy) as FriendshipInvitation[] : [];
    return Array.isArray(parsedLegacy) ? parsedLegacy : [];
  } catch {
    return [];
  }
}

function storeFriendInvitations(invitations: FriendshipInvitation[]) {
  try {
    localStorage.setItem(FRIEND_INVITATIONS_STORAGE_KEY, JSON.stringify(invitations));
  } catch {
    // armazenamento local indisponível
  }
}

function loadNotifications(): AppNotification[] {
  try {
    const raw = localStorage.getItem(APP_NOTIFICATIONS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as AppNotification[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function storeNotifications(notifications: AppNotification[]) {
  try {
    localStorage.setItem(APP_NOTIFICATIONS_STORAGE_KEY, JSON.stringify(notifications));
  } catch {
    // armazenamento local indisponível
  }
}
const SHARED_TRANSACTIONS_STORAGE_KEY = 'rubylife-shared-transactions';

function loadStoredSharedTransactions(): SharedTransactionRequest[] {
  try {
    const raw = localStorage.getItem(SHARED_TRANSACTIONS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as SharedTransactionRequest[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function storeSharedTransactions(requests: SharedTransactionRequest[]) {
  try {
    localStorage.setItem(SHARED_TRANSACTIONS_STORAGE_KEY, JSON.stringify(requests));
  } catch {
    // armazenamento local indisponível
  }
}

function sharedTransactionLabel(item: Transaction) {
  if (!item.sharedCreatedByName) return null;
  return `${item.type === 'income' ? 'Receita' : 'Despesa'} compartilhada por ${item.sharedCreatedByName}`;
}

const SHOPPING_LISTS_STORAGE_KEY = 'rubylife-shopping-lists';
const ACCOUNT_TYPE_LABELS: Record<AccountType, string> = {
  wallet: 'Carteira',
  checking: 'Conta corrente',
  savings: 'Poupança',
  investment: 'Investimentos',
};
const ACCOUNT_COLORS = [
  '#1B99D8', '#0891B2', '#0F766E', '#16A34A', '#65A30D', '#F59E0B', '#EA580C', '#DC2626',
  '#E11D48', '#DB2777', '#C026D3', '#9333EA', '#7C3AED', '#4F46E5', '#2563EB', '#475569',
];
const ACCOUNT_ICON_OPTIONS: Array<{ key: AccountIconKey; label: string }> = [
  { key: 'wallet', label: 'Carteira' },
  { key: 'card', label: 'Cartão' },
  { key: 'bank', label: 'Banco' },
  { key: 'savings', label: 'Poupança' },
  { key: 'investment', label: 'Investimentos' },
  { key: 'cash', label: 'Dinheiro' },
];
const CATEGORY_KIND_LABELS: Record<CategoryKind, string> = {
  income: 'Receita',
  expense: 'Despesa',
  transfer: 'Transferência',
};
const INCOME_CATEGORIES = ['Salário', 'Extra', 'Freelance', 'Vendas', 'Reembolso', 'Investimentos', 'Aluguel', 'Cashback', 'Presente', 'Empréstimo', 'Outros'];
const EXPENSE_CATEGORIES = ['Moradia', 'Alimentação', 'Transporte', 'Saúde', 'Educação', 'Lazer', 'Compras', 'Assinaturas', 'Dívidas', 'Cartão', 'Família', 'Pets', 'Impostos', 'Trabalho', 'Doações', 'Saques', 'Outros'];
const TRANSFER_CATEGORIES = ['Entre contas', 'Para carteira', 'Para poupança', 'Para investimentos'];
const CATEGORY_GROUPS = [
  { key: 'income', label: 'Receitas', categories: INCOME_CATEGORIES },
  { key: 'expense', label: 'Despesas', categories: EXPENSE_CATEGORIES },
  { key: 'transfer', label: 'Transferência', categories: TRANSFER_CATEGORIES },
] as const;
const MONTH_OPTIONS = Array.from({ length: 12 }, (_, month) => ({
  value: month,
  label: new Intl.DateTimeFormat('pt-BR', { month: 'long' }).format(new Date(2026, month, 1)),
}));

const initialForm: LaunchForm = {
  description: '',
  category: 'Outros',
  accountId: '',
  amount: '',
  dueDate: new Date().toISOString().slice(0, 10),
  recurrence: 'single',
  installments: '12',
  fixedUntil: new Date(new Date().getFullYear(), 11, 1).toISOString().slice(0, 7),
};

// Dados vêm 100% do Supabase. Conta nova começa vazia (sem carteira,
// categoria ou transação de exemplo). Veja src/lib/db.ts.

function goalSaved(goal: Goal) {
  return goal.movements.reduce((sum, movement) => sum + (movement.type === 'deposit' ? movement.amount : -movement.amount), 0);
}

function addMonths(date: string, months: number) {
  const [year = 2026, month = 1, day = 1] = date.split('-').map(Number);
  const next = new Date(year, month - 1 + months, day);
  const wantedMonth = (month - 1 + months) % 12;
  if (next.getMonth() !== (wantedMonth + 12) % 12) next.setDate(0);
  return next.toISOString().slice(0, 10);
}

function monthKey(date: Date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function formatDate(date: string) {
  return new Intl.DateTimeFormat('pt-BR', { timeZone: 'UTC' }).format(new Date(`${date}T00:00:00Z`));
}

function formatDateGroup(date: string) {
  return new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: 'short', year: 'numeric', timeZone: 'UTC' })
    .format(new Date(`${date}T00:00:00Z`))
    .replace('.', '');
}

function displayTransactionDescription(description: string) {
  return description
    .trim()
    .replace(new RegExp(' [0-9]+ */ *[0-9]+$'), '')
    .slice(0, TRANSACTION_DESCRIPTION_MAX_LENGTH)
    .trimEnd();
}

function formatCurrency(value: number) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
}

function recurrenceLabel(item: Transaction) {
  if (item.recurrence === 'fixed') return 'Fixa';
  if (item.recurrence === 'installment') {
    return item.installmentNumber && item.installmentTotal ? `Parcela ${item.installmentNumber}/${item.installmentTotal}` : 'Parcelada';
  }
  return 'Única';
}
const REPORT_COLUMNS = ['Tipo', 'Descrição', 'Categoria', 'Vencimento', 'Valor previsto', 'Valor real', 'Status', 'Conta'];

function slugifyFileName(value: string) {
  return value
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'relatorio';
}

function reportTotals(items: Transaction[]) {
  const income = items.filter((item) => item.type === 'income').reduce((sum, item) => sum + item.amount, 0);
  const expense = items.filter((item) => item.type === 'expense').reduce((sum, item) => sum + item.amount, 0);
  const settledIncome = items.filter((item) => item.type === 'income' && item.status === 'settled').reduce((sum, item) => sum + (item.settledAmount ?? item.amount), 0);
  const settledExpense = items.filter((item) => item.type === 'expense' && item.status === 'settled').reduce((sum, item) => sum + (item.settledAmount ?? item.amount), 0);
  const pendingIncome = items.filter((item) => item.type === 'income' && item.status === 'open').reduce((sum, item) => sum + item.amount, 0);
  const pendingExpense = items.filter((item) => item.type === 'expense' && item.status === 'open').reduce((sum, item) => sum + item.amount, 0);
  return { income, expense, balance: income - expense, settledIncome, settledExpense, pendingIncome, pendingExpense };
}

function exportReportPdf(items: Transaction[], title: string) {
  const doc = new jsPDF({ orientation: 'landscape' });
  const totals = reportTotals(items);

  // Title
  doc.setFontSize(16);
  doc.setTextColor(27, 153, 216);
  doc.text(title, 14, 16);

  // Summary table
  doc.setFontSize(10);
  autoTable(doc, {
    head: [['', 'Receitas', 'Despesas', 'Saldo']],
    body: [
      ['Previsto', formatCurrency(totals.income), formatCurrency(totals.expense), formatCurrency(totals.balance)],
      ['Realizado', formatCurrency(totals.settledIncome), formatCurrency(totals.settledExpense), formatCurrency(totals.settledIncome - totals.settledExpense)],
      ['Pendente', formatCurrency(totals.pendingIncome), formatCurrency(totals.pendingExpense), formatCurrency(totals.pendingIncome - totals.pendingExpense)],
    ],
    startY: 22,
    theme: 'grid',
    styles: { fontSize: 9, cellPadding: 4, halign: 'center' },
    headStyles: { fillColor: [27, 153, 216], textColor: 255, fontStyle: 'bold' },
    columnStyles: { 0: { halign: 'left', fontStyle: 'bold' } },
    alternateRowStyles: { fillColor: [244, 249, 252] },
  });

  // Transactions table
  const summaryEnd = (doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? 50;
  doc.setFontSize(11);
  doc.setTextColor(50);
  doc.text(`Transações (${items.length})`, 14, summaryEnd + 10);

  autoTable(doc, {
    head: [REPORT_COLUMNS],
    body: items.map((item) => [
      item.type === 'income' ? 'Receita' : 'Despesa',
      item.description,
      item.category,
      formatDate(item.dueDate),
      formatCurrency(item.amount),
      item.status === 'settled' ? formatCurrency(item.settledAmount ?? item.amount) : '—',
      item.status === 'settled' ? (item.type === 'income' ? 'Recebido' : 'Pago') : 'Aberto',
      item.status === 'settled' ? (item.account ?? '') : '—',
    ]),
    startY: summaryEnd + 14,
    styles: { fontSize: 8, cellPadding: 2 },
    headStyles: { fillColor: [27, 153, 216], textColor: 255 },
    alternateRowStyles: { fillColor: [244, 249, 252] },
  });
  doc.save(`${slugifyFileName(title)}.pdf`);
}

const excelHeader = (value: string): Cell => ({ value, fontWeight: 'bold', backgroundColor: '#1B99D8', textColor: '#FFFFFF' });
const excelCurrency = (value: number): Cell => ({ value, type: Number, format: '"R$" #,##0.00' });

async function exportReportExcel(items: Transaction[], title: string) {
  const totals = reportTotals(items);
  const summaryData: SheetData = [
    ['', excelHeader('Receitas'), excelHeader('Despesas'), excelHeader('Saldo')],
    [excelHeader('Previsto'), excelCurrency(totals.income), excelCurrency(totals.expense), excelCurrency(totals.balance)],
    [excelHeader('Realizado'), excelCurrency(totals.settledIncome), excelCurrency(totals.settledExpense), excelCurrency(totals.settledIncome - totals.settledExpense)],
    [excelHeader('Pendente'), excelCurrency(totals.pendingIncome), excelCurrency(totals.pendingExpense), excelCurrency(totals.pendingIncome - totals.pendingExpense)],
  ];
  const transactionData: SheetData = [
    REPORT_COLUMNS.map(excelHeader),
    ...items.map((item) => [
      item.type === 'income' ? 'Receita' : 'Despesa', item.description, item.category, formatDate(item.dueDate),
      excelCurrency(item.amount), item.status === 'settled' ? excelCurrency(item.settledAmount ?? item.amount) : '',
      item.status === 'settled' ? (item.type === 'income' ? 'Recebido' : 'Pago') : 'Aberto',
      item.status === 'settled' ? (item.account ?? '') : '',
    ]),
  ];
  await writeXlsxFile([
    { data: summaryData, sheet: 'Resumo', columns: [14, 18, 18, 18].map((width) => ({ width })) },
    { data: transactionData, sheet: 'Transações', columns: [10, 28, 18, 12, 15, 15, 12, 16].map((width) => ({ width })), stickyRowsCount: 1 },
  ]).toFile(`${slugifyFileName(title)}.xlsx`);
}

async function downloadImportTemplate() {
  const headers = ['Tipo', 'Descrição', 'Valor', 'Vencimento', 'Categoria', 'Recorrência', 'Parcelas'];
  const data: SheetData = [
    headers.map(excelHeader),
    ['Receita', 'Salário', excelCurrency(3500), '05/01/2026', 'Salário', 'Fixa', ''],
    ['Despesa', 'Aluguel', excelCurrency(1200), '10/01/2026', 'Moradia', 'Fixa', ''],
    ['Despesa', 'Notebook', excelCurrency(4800), '15/01/2026', 'Compras', 'Parcelada', 12],
    ['Receita', 'Freelance', excelCurrency(800), '20/01/2026', 'Freelance', 'Única', ''],
  ];
  await writeXlsxFile(data, {
    sheet: 'Modelo', columns: [10, 26, 12, 14, 18, 14, 10].map((width) => ({ width })), stickyRowsCount: 1,
  }).toFile('modelo-importacao-transacoes.xlsx');
}

function accountIconFor(account: Account): AccountIconKey {
  if (account.icon) return account.icon;
  if (account.type === 'wallet') return 'wallet';
  if (account.type === 'savings') return 'savings';
  if (account.type === 'investment') return 'investment';
  return 'card';
}

function AccountIconGraphic({ icon, size = 17 }: { icon: AccountIconKey; size?: number }) {
  if (icon === 'wallet') return <WalletCards size={size} />;
  if (icon === 'card') return <CreditCard size={size} />;
  if (icon === 'bank') return <Landmark size={size} />;
  if (icon === 'savings') return <PiggyBank size={size} />;
  if (icon === 'investment') return <TrendingUp size={size} />;
  return <Banknote size={size} />;
}

function defaultCategoryIcon(name: string, kind: CategoryKind): CategoryIconKey {
  const icons: Record<string, CategoryIconKey> = {
    'Salário': 'salary', 'Extra': 'money', 'Freelance': 'work', 'Vendas': 'shopping', 'Reembolso': 'refund', 'Investimentos': 'investment',
    'Aluguel': 'home', 'Cashback': 'card', 'Presente': 'gift', 'Empréstimo': 'money', 'Moradia': 'home', 'Alimentação': 'food',
    'Transporte': 'car', 'Saúde': 'health', 'Educação': 'education', 'Lazer': 'leisure', 'Compras': 'shopping', 'Assinaturas': 'repeat',
    'Dívidas': 'money', 'Cartão': 'card', 'Família': 'family', 'Pets': 'pets', 'Impostos': 'money', 'Trabalho': 'work', 'Doações': 'donation',
    'Saques': 'wallet', 'Entre contas': 'transfer', 'Para carteira': 'wallet', 'Para poupança': 'money', 'Para investimentos': 'investment',
  };
  return icons[name] ?? (kind === 'transfer' ? 'transfer' : 'tag');
}
function CategoryIconGraphic({ icon, size = 17 }: { icon: CategoryIconKey; size?: number }) {
  if (icon === 'salary') return <BadgeDollarSign size={size} />;
  if (icon === 'money') return <Coins size={size} />;
  if (icon === 'work') return <BriefcaseBusiness size={size} />;
  if (icon === 'shopping') return <ShoppingBag size={size} />;
  if (icon === 'refund') return <RotateCcw size={size} />;
  if (icon === 'investment') return <TrendingUp size={size} />;
  if (icon === 'home') return <House size={size} />;
  if (icon === 'card') return <CreditCard size={size} />;
  if (icon === 'gift') return <Gift size={size} />;
  if (icon === 'food') return <Utensils size={size} />;
  if (icon === 'car') return <Car size={size} />;
  if (icon === 'health') return <HeartPulse size={size} />;
  if (icon === 'education') return <GraduationCap size={size} />;
  if (icon === 'leisure') return <Gamepad2 size={size} />;
  if (icon === 'repeat') return <Repeat2 size={size} />;
  if (icon === 'family') return <Users size={size} />;
  if (icon === 'pets') return <PawPrint size={size} />;
  if (icon === 'donation') return <HandHeart size={size} />;
  if (icon === 'wallet') return <WalletCards size={size} />;
  if (icon === 'transfer') return <ArrowLeftRight size={size} />;
  if (icon === 'travel') return <Plane size={size} />;
  if (icon === 'fitness') return <Dumbbell size={size} />;
  if (icon === 'music') return <Music size={size} />;
  if (icon === 'maintenance') return <Wrench size={size} />;
  return <Tags size={size} />;
}

type HsvColor = { h: number; s: number; v: number };

function hexToHsv(hex: string): HsvColor {
  const normalized = hex.replace('#', '').padEnd(6, '0').slice(0, 6);
  const red = Number.parseInt(normalized.slice(0, 2), 16) / 255;
  const green = Number.parseInt(normalized.slice(2, 4), 16) / 255;
  const blue = Number.parseInt(normalized.slice(4, 6), 16) / 255;
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const delta = max - min;
  let hue = 0;
  if (delta) {
    if (max === red) hue = 60 * (((green - blue) / delta) % 6);
    else if (max === green) hue = 60 * ((blue - red) / delta + 2);
    else hue = 60 * ((red - green) / delta + 4);
  }
  return { h: hue < 0 ? hue + 360 : hue, s: max ? (delta / max) * 100 : 0, v: max * 100 };
}

function hsvToRgb({ h, s, v }: HsvColor) {
  const saturation = s / 100;
  const value = v / 100;
  const chroma = value * saturation;
  const section = h / 60;
  const x = chroma * (1 - Math.abs((section % 2) - 1));
  const offset = value - chroma;
  let red = 0;
  let green = 0;
  let blue = 0;
  if (section < 1) [red, green] = [chroma, x];
  else if (section < 2) [red, green] = [x, chroma];
  else if (section < 3) [green, blue] = [chroma, x];
  else if (section < 4) [green, blue] = [x, chroma];
  else if (section < 5) [red, blue] = [x, chroma];
  else[red, blue] = [chroma, x];
  return {
    r: Math.round((red + offset) * 255),
    g: Math.round((green + offset) * 255),
    b: Math.round((blue + offset) * 255),
  };
}

function hsvToHex(hsv: HsvColor) {
  const { r, g, b } = hsvToRgb(hsv);
  return `#${[r, g, b].map((channel) => channel.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

function ColorSpectrumSheet({ value, onChange, onClose }: { value: string; onChange: (color: string) => void; onClose: () => void }) {
  const [hsv, setHsv] = useState<HsvColor>(() => hexToHsv(value));
  const rgb = hsvToRgb(hsv);

  useEffect(() => {
    function handleEscape(event: KeyboardEvent) {
      if (event.key === 'Escape') onClose();
    }
    document.addEventListener('keydown', handleEscape);
    return () => document.removeEventListener('keydown', handleEscape);
  }, [onClose]);

  function applyColor(next: HsvColor) {
    setHsv(next);
    onChange(hsvToHex(next));
  }

  function updateSpectrum(event: React.PointerEvent<HTMLDivElement>) {
    const bounds = event.currentTarget.getBoundingClientRect();
    const saturation = Math.min(100, Math.max(0, ((event.clientX - bounds.left) / bounds.width) * 100));
    const brightness = 100 - Math.min(100, Math.max(0, ((event.clientY - bounds.top) / bounds.height) * 100));
    applyColor({ ...hsv, s: saturation, v: brightness });
  }

  return createPortal(
    <div className="category-icon-picker-layer" onClick={onClose}>
      <div className="category-icon-picker-dialog color-spectrum-dialog" role="dialog" aria-modal="true" aria-label="Escolher cor" onClick={(event) => event.stopPropagation()}>
        <div className="category-icon-picker-head">
          <strong>Escolher cor</strong>
          <button type="button" onClick={onClose} aria-label="Fechar seletor de cor"><X size={18} /></button>
        </div>
        <div className="color-spectrum-body">
          <div
            className="color-spectrum-area"
            style={{ backgroundColor: `hsl(${hsv.h} 100% 50%)` }}
            role="slider"
            aria-label="Saturação e brilho"
            aria-valuetext={hsvToHex(hsv)}
            onPointerDown={(event) => { event.currentTarget.setPointerCapture(event.pointerId); updateSpectrum(event); }}
            onPointerMove={(event) => { if (event.currentTarget.hasPointerCapture(event.pointerId)) updateSpectrum(event); }}
          >
            <span className="color-spectrum-cursor" style={{ left: `${hsv.s}%`, top: `${100 - hsv.v}%` }} />
          </div>
          <div className="color-hue-row">
            <span className="color-preview" style={{ backgroundColor: hsvToHex(hsv) }} />
            <input type="range" min="0" max="359" value={Math.round(hsv.h)} onChange={(event) => applyColor({ ...hsv, h: Number(event.target.value) })} aria-label="Tonalidade" />
          </div>
          <div className="color-value-row">
            <span><strong>{rgb.r}</strong><small>R</small></span>
            <span><strong>{rgb.g}</strong><small>G</small></span>
            <span><strong>{rgb.b}</strong><small>B</small></span>
            <span className="color-hex-value"><strong>{hsvToHex(hsv)}</strong><small>HEX</small></span>
          </div>
        </div>
      </div>
    </div>,
    document.body,
  );
}
function parseAmount(value: string) {
  const normalized = value.replace(/\./g, '').replace(',', '.').replace(/[\u0300-\u036f]/g, '');
  return Number(normalized || 0);
}

function formatCurrencyInput(value: string | number) {
  if (value === undefined || value === null || value === '') return '';
  let raw = typeof value === 'number' ? value.toFixed(2).replace('.', ',') : String(value);
  if (raw.includes('.') && !raw.includes(',')) {
    const lastDot = raw.lastIndexOf('.');
    if (raw.length - lastDot <= 3) {
      raw = raw.slice(0, lastDot) + ',' + raw.slice(lastDot + 1);
    }
  }
  const clean = raw.replace(/[\u0300-\u036f]/g, '');
  if (!clean) return '';
  const parts = clean.split(',');
  let intPart = parts[0].replace(/\./g, '').replace(/^0+(?=\d)/, '');
  if (intPart === '' && parts.length > 1) intPart = '0';
  const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  if (parts.length > 1) {
    const decPart = parts.slice(1).join('').slice(0, 2);
    return `${formattedInt},${decPart}`;
  }
  return formattedInt;
}

function normalizeHeader(value: unknown) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function normalizeImportedDate(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, '0')}-${String(value.getUTCDate()).padStart(2, '0')}`;
  }

  if (typeof value === 'number' && value > 0) {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
  }

  const text = String(value ?? '').trim();
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) return isValidCalendarDate(text) ? text : '';
  const brazilianMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!brazilianMatch) return '';
  const normalized = `${brazilianMatch[3]}-${brazilianMatch[2]?.padStart(2, '0')}-${brazilianMatch[1]?.padStart(2, '0')}`;
  return isValidCalendarDate(normalized) ? normalized : '';
}

function isValidCalendarDate(value: string) {
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(Date.UTC(year ?? 0, (month ?? 1) - 1, day ?? 1));
  return date.getUTCFullYear() === year && date.getUTCMonth() === (month ?? 1) - 1 && date.getUTCDate() === day;
}

function parseImportedAmount(value: unknown) {
  if (typeof value === 'number') return Math.abs(value);
  // Strip currency symbols, spaces and diacritics \u2014 keep only digits, comma, dot and minus
  const text = String(value ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\d.,-]/g, '')
    .trim();
  if (!text) return 0;
  // Brazilian format (1.234,56): comma is decimal, dot is thousands separator
  // Plain format (1234.56): dot is decimal
  const normalized = text.includes(',') ? text.replace(/\./g, '').replace(',', '.') : text;
  return Math.abs(Number(normalized) || 0);
}

function parseImportedType(value: unknown): TransactionType | null {
  const normalized = normalizeHeader(value);
  if (['receita', 'entrada', 'income'].includes(normalized)) return 'income';
  if (['despesa', 'saida', 'expense'].includes(normalized)) return 'expense';
  return null;
}

function parseImportedRecurrence(value: unknown): RecurrenceType | null {
  const normalized = normalizeHeader(value);
  if (!normalized || ['unica', 'unico', 'single'].includes(normalized)) return 'single';
  if (['fixa', 'fixo', 'fixed'].includes(normalized)) return 'fixed';
  if (['parcelada', 'parcelado', 'installment'].includes(normalized)) return 'installment';
  return null;
}

const MAX_IMPORT_FILE_SIZE = 5 * 1024 * 1024;
const MAX_IMPORT_UNCOMPRESSED_SIZE = 25 * 1024 * 1024;

async function isSafeXlsxContainer(file: File) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) return false;
  const validMagic = (bytes[2] === 0x03 && bytes[3] === 0x04)
    || (bytes[2] === 0x05 && bytes[3] === 0x06)
    || (bytes[2] === 0x07 && bytes[3] === 0x08);
  if (!validMagic) return false;

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder();
  let entries = 0;
  let totalUncompressed = 0;
  let hasContentTypes = false;
  let hasWorkbook = false;

  for (let offset = 0; offset + 46 <= bytes.length;) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      offset += 1;
      continue;
    }
    const uncompressedSize = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const entryEnd = offset + 46 + fileNameLength + extraLength + commentLength;
    if (entryEnd > bytes.length || uncompressedSize === 0xffffffff) return false;

    const name = decoder.decode(bytes.slice(offset + 46, offset + 46 + fileNameLength)).replace(/\\/g, '/');
    if (name.startsWith('/') || name.split('/').includes('..')) return false;
    hasContentTypes ||= name === '[Content_Types].xml';
    hasWorkbook ||= name === 'xl/workbook.xml';
    totalUncompressed += uncompressedSize;
    entries += 1;
    if (entries > 1000 || totalUncompressed > MAX_IMPORT_UNCOMPRESSED_SIZE) return false;
    offset = entryEnd;
  }

  return entries > 0 && hasContentTypes && hasWorkbook;
}

async function parseExcelTransactions(file: File): Promise<ImportPreviewRow[]> {
  const rows: Row[] = await readSheet(file);
  if (rows.length < 2) throw new Error('A planilha precisa ter um cabeçalho e pelo menos uma transação.');
  if (rows.length > 5001) throw new Error('A planilha pode conter no máximo 5.000 transações.');
  if ((rows[0]?.length ?? 0) > 50) throw new Error('A planilha possui colunas demais.');

  const headers = rows[0]?.map(normalizeHeader) ?? [];
  const columnAliases = {
    type: ['tipo', 'natureza'],
    description: ['descricao', 'historico'],
    amount: ['valor', 'montante'],
    dueDate: ['vencimento', 'data', 'data de vencimento'],
    category: ['categoria'],
    recurrence: ['recorrencia'],
    installments: ['parcelas', 'quantidade de parcelas'],
  };
  const findColumn = (aliases: string[]) => headers.findIndex((header) => aliases.includes(header));
  const columns = {
    type: findColumn(columnAliases.type),
    description: findColumn(columnAliases.description),
    amount: findColumn(columnAliases.amount),
    dueDate: findColumn(columnAliases.dueDate),
    category: findColumn(columnAliases.category),
    recurrence: findColumn(columnAliases.recurrence),
    installments: findColumn(columnAliases.installments),
  };
  const requiredColumns = [
    ['Tipo', columns.type],
    ['Descrição', columns.description],
    ['Valor', columns.amount],
    ['Vencimento', columns.dueDate],
    ['Categoria', columns.category],
  ] as const;
  const missingColumns = requiredColumns.filter(([, index]) => index < 0).map(([label]) => label);
  if (missingColumns.length) throw new Error(`Colunas obrigatórias não encontradas: ${missingColumns.join(', ')}.`);

  return rows.slice(1).flatMap((row, index) => {
    if (row.every((cell) => cell === null || String(cell).trim() === '')) return [];

    const type = parseImportedType(row[columns.type]);
    const description = String(row[columns.description] ?? '').trim().slice(0, TRANSACTION_DESCRIPTION_MAX_LENGTH);
    const amount = parseImportedAmount(row[columns.amount]);
    const dueDate = normalizeImportedDate(row[columns.dueDate]);
    const category = String(row[columns.category] ?? '').trim();
    const recurrence = columns.recurrence >= 0 ? parseImportedRecurrence(row[columns.recurrence]) : 'single';
    const installments = columns.installments >= 0 ? String(row[columns.installments] ?? '1') : '1';
    const errors: string[] = [];

    if (!type) errors.push('tipo inválido');
    if (!description) errors.push('descrição ausente');
    if (amount <= 0) errors.push('valor inválido');
    if (!dueDate) errors.push('vencimento inválido');
    if (!category) errors.push('categoria ausente');
    if (!recurrence) errors.push('recorrência inválida');
    if (recurrence === 'installment' && Number(installments) < 1) errors.push('parcelas inválidas');

    const transactions = errors.length || !type || !recurrence ? [] : generateTransactions({
      description,
      amount: String(amount).replace('.', ','),
      dueDate,
      category,
      accountId: '',
      recurrence,
      installments,
      fixedUntil: '',
    }, type);

    return [{
      rowNumber: index + 2,
      type,
      description,
      amount,
      dueDate,
      category,
      transactions,
      error: errors.length ? errors.join(', ') : undefined,
    }];
  });
}

function generateTransactions(form: LaunchForm, type: TransactionType): Transaction[] {
  const amount = parseAmount(form.amount);
  const groupId = crypto.randomUUID();
  const base = {
    groupId,
    description: form.description.trim().slice(0, TRANSACTION_DESCRIPTION_MAX_LENGTH),
    category: form.category,
    accountId: form.accountId || undefined,
    amount,
    type,
    recurrence: form.recurrence,
    status: 'open' as const,
  };

  if (form.recurrence === 'installment') {
    const total = Math.max(1, Number(form.installments || 1));
    return Array.from({ length: total }, (_, index) => ({
      ...base,
      id: crypto.randomUUID(),
      dueDate: addMonths(form.dueDate, index),
      installmentNumber: index + 1,
      installmentTotal: total,
    }));
  }

  if (form.recurrence === 'fixed') {
    const startYear = Number(form.dueDate.slice(0, 4));
    const startMonth = Number(form.dueDate.slice(5, 7));
    const fallbackUntil = `${startYear}-12`;
    const until = form.fixedUntil || fallbackUntil;
    const untilYear = Number(until.slice(0, 4));
    const untilMonth = Number(until.slice(5, 7));
    const monthCount = Math.max(1, (untilYear - startYear) * 12 + untilMonth - startMonth + 1);

    return Array.from({ length: monthCount }, (_, index) => ({
      ...base,
      id: crypto.randomUUID(),
      dueDate: addMonths(form.dueDate, index),
    }));
  }

  return [{ ...base, id: crypto.randomUUID(), dueDate: form.dueDate }];
}

function SplashScreen() {
  return (
    <div className="splash-screen" role="status" aria-label="Carregando RubyLife">
      <img className="splash-logo" src={rubyLogoWhite} alt="RubyLife" />
      <span className="splash-spinner" aria-hidden="true" />
    </div>
  );
}

function translateAuthError(message: string): string {
  const m = message.toLowerCase();
  if (m.includes('invalid login credentials')) return 'E-mail ou senha inválidos.';
  if (m.includes('email not confirmed')) return 'Confirme seu e-mail antes de entrar.';
  if (m.includes('already registered') || m.includes('already been registered')) return 'Este e-mail já está cadastrado.';
  if (m.includes('password should be at least')) return 'A senha deve ter pelo menos 6 caracteres.';
  if (m.includes('unable to validate email') || m.includes('invalid email')) return 'E-mail inválido.';
  return message;
}

const REMEMBER_EMAIL_KEY = 'rubylife-saved-email';

function LoginScreen() {
  const [mode, setMode] = useState<'login' | 'recovery' | 'register'>('login');
  const [fullName, setFullName] = useState('');
  const [email, setEmail] = useState(() => localStorage.getItem(REMEMBER_EMAIL_KEY) ?? '');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);
  const [rememberMe, setRememberMe] = useState(() => Boolean(localStorage.getItem(REMEMBER_EMAIL_KEY)));

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setMessage(null);

    if (mode === 'recovery') {
      if (!email) return;
      setLoading(true);
      const { error } = await supabase.auth.resetPasswordForEmail(email, { redirectTo: window.location.origin });
      setLoading(false);
      setMessage(error
        ? { text: translateAuthError(error.message), type: 'error' }
        : { text: 'Link de recuperação enviado para o e-mail informado!', type: 'success' });
      return;
    }

    if (mode === 'register') {
      if (fullName.trim().length < 2) {
        setMessage({ text: 'Informe seu nome completo.', type: 'error' });
        return;
      }
      if (password.length < 8) {
        setMessage({ text: 'A senha deve ter pelo menos 8 caracteres.', type: 'error' });
        return;
      }
      if (password !== confirmPassword) {
        setMessage({ text: 'As senhas não coincidem. Verifique e tente novamente.', type: 'error' });
        return;
      }
      setLoading(true);
      const { data, error } = await supabase.auth.signUp({ email, password, fullName: fullName.trim() });
      setLoading(false);
      if (error) {
        setMessage({ text: translateAuthError(error.message), type: 'error' });
        return;
      }
      // Se a confirmação de e-mail estiver ligada, ainda não há sessão.
      if (!data.session) {
        setMessage({ text: 'Conta criada! Verifique seu e-mail para confirmar o acesso.', type: 'success' });
      }
      // Com confirmação desligada, o onAuthStateChange assume e entra direto.
      return;
    }

    setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (error) {
      setMessage({ text: translateAuthError(error.message), type: 'error' });
      return;
    }
    if (rememberMe) localStorage.setItem(REMEMBER_EMAIL_KEY, email);
    else localStorage.removeItem(REMEMBER_EMAIL_KEY);
    // Sucesso: a sessão muda e a App troca de tela automaticamente.
  }

  function switchMode(newMode: 'login' | 'recovery' | 'register') {
    setMode(newMode);
    setMessage(null);
  }

  return (
    <main className="toodledo-login-shell" aria-label="Login RubyLife">
      <div className="toodledo-login-container">
        <div className="toodledo-brand-side">
          <div className="toodledo-logo-wrapper">
            <img className="toodledo-logo-image toodledo-logo-image--light" src={rubyLogoWhite} alt="RubyLife" />
            <img className="toodledo-logo-image toodledo-logo-image--dark" src={rubyLogoColor} alt="RubyLife" />
          </div>
        </div>

        <div className="toodledo-divider" aria-hidden="true" />

        <div className="toodledo-form-side">
          <form className="toodledo-form" onSubmit={handleSubmit}>
            <h1 className="toodledo-form-title">
              {mode === 'login' && 'Entrar'}
              {mode === 'recovery' && 'Recuperar Senha'}
              {mode === 'register' && 'Criar Conta'}
            </h1>

            <div className="toodledo-inputs">
              {mode === 'register' && (
                <input
                  value={fullName}
                  onChange={(event) => setFullName(event.target.value)}
                  type="text"
                  placeholder="Nome completo"
                  autoComplete="name"
                  minLength={2}
                  maxLength={160}
                  required
                  className="toodledo-input"
                />
              )}
              <input
                value={email}
                onChange={(event) => setEmail(event.target.value)}
                type="email"
                placeholder="E-mail"
                autoComplete="email"
                required
                className="toodledo-input"
              />
              {mode !== 'recovery' && (
                <div className="toodledo-password-field">
                  <input
                    value={password}
                    onChange={(event) => setPassword(event.target.value)}
                    type={showPassword ? 'text' : 'password'}
                    placeholder={mode === 'register' ? 'Criar senha' : 'Senha'}
                    autoComplete={mode === 'register' ? 'new-password' : 'current-password'}
                    required
                    minLength={mode === 'register' ? 8 : undefined}
                    className="toodledo-input"
                  />
                  <button
                    type="button"
                    className="toodledo-password-toggle"
                    onClick={() => setShowPassword((visible) => !visible)}
                    aria-label={showPassword ? 'Ocultar senha' : 'Mostrar senha'}
                    title={showPassword ? 'Ocultar senha' : 'Mostrar senha'}
                  >
                    {showPassword ? <EyeOff size={20} /> : <Eye size={20} />}
                  </button>
                </div>
              )}
              {mode === 'register' && (
                <input
                  value={confirmPassword}
                  onChange={(event) => setConfirmPassword(event.target.value)}
                  type="password"
                  placeholder="Confirmar senha"
                  autoComplete="new-password"
                  required
                  className="toodledo-input"
                />
              )}
            </div>

            <div className="toodledo-secondary-actions">
              {mode === 'login' ? (
                <>
                  <button className="toodledo-link-btn" type="button" onClick={() => switchMode('recovery')}>
                    Esqueceu a senha?
                  </button>
                  <button className="toodledo-link-btn" type="button" onClick={() => switchMode('register')}>
                    Criar conta
                  </button>
                </>
              ) : (
                <button className="toodledo-link-btn" type="button" onClick={() => switchMode('login')}>
                  ← Voltar para o login
                </button>
              )}
            </div>

            <div className="toodledo-main-action">
              <button className="toodledo-submit" type="submit" disabled={loading}>
                {loading ? 'Aguarde…' : (
                  <>
                    {mode === 'login' && 'Entrar'}
                    {mode === 'recovery' && 'Enviar Link'}
                    {mode === 'register' && 'Cadastrar'}
                  </>
                )}
              </button>
            </div>

            {mode === 'login' ? (
              <label className="toodledo-remember">
                <input
                  type="checkbox"
                  checked={rememberMe}
                  onChange={(e) => setRememberMe(e.target.checked)}
                />
                Salvar dados de acesso
              </label>
            ) : null}

            {message ? (
              <p className={`toodledo-recovery-msg ${message.type === 'error' ? 'toodledo-msg-error' : ''}`} role="status">
                {message.text}
              </p>
            ) : null}
          </form>
        </div>
      </div>
    </main>
  );
}
function UpdatePasswordScreen({ onDone }: { onDone: () => void }) {
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setMessage(null);
    if (password.length < 8) {
      setMessage({ text: 'A senha deve ter pelo menos 8 caracteres.', type: 'error' });
      return;
    }
    if (password !== confirmPassword) {
      setMessage({ text: 'As senhas não coincidem. Verifique e tente novamente.', type: 'error' });
      return;
    }
    setLoading(true);
    const { error } = await supabase.auth.updateUser({ password });
    setLoading(false);
    if (error) {
      setMessage({ text: translateAuthError(error.message), type: 'error' });
      return;
    }
    setMessage({ text: 'Senha atualizada com sucesso!', type: 'success' });
    window.setTimeout(onDone, 900);
  }

  return (
    <main className="toodledo-login-shell" aria-label="Definir nova senha">
      <div className="toodledo-login-container">
        <div className="toodledo-brand-side">
          <div className="toodledo-logo-wrapper">
            <img className="toodledo-logo-image toodledo-logo-image--light" src={rubyLogoWhite} alt="RubyLife" />
            <img className="toodledo-logo-image toodledo-logo-image--dark" src={rubyLogoColor} alt="RubyLife" />
          </div>
        </div>

        <div className="toodledo-divider" aria-hidden="true" />

        <div className="toodledo-form-side">
          <form className="toodledo-form" onSubmit={handleSubmit}>
            <h1 className="toodledo-form-title">Definir nova senha</h1>

            <div className="toodledo-inputs">
              <input
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                type="password"
                placeholder="Nova senha"
                autoComplete="new-password"
                minLength={8}
                required
                className="toodledo-input"
              />
              <input
                value={confirmPassword}
                onChange={(event) => setConfirmPassword(event.target.value)}
                type="password"
                placeholder="Confirmar nova senha"
                autoComplete="new-password"
                required
                className="toodledo-input"
              />
            </div>

            <div className="toodledo-main-action">
              <button className="toodledo-submit" type="submit" disabled={loading}>
                {loading ? 'Aguarde…' : 'Salvar nova senha'}
              </button>
            </div>


            {message ? (
              <p className={`toodledo-recovery-msg ${message.type === 'error' ? 'toodledo-msg-error' : ''}`} role="status">
                {message.text}
              </p>
            ) : null}
          </form>
        </div>
      </div>
    </main>
  );
}

export function App() {
  const [session, setSession] = useState<Session | null>(null);
  const [authReady, setAuthReady] = useState(false);
  const [recoveryMode, setRecoveryMode] = useState(false);
  const [dataLoading, setDataLoading] = useState(false);
  const [syncError, setSyncError] = useState<string | null>(null);
  const [storedProfileName, setStoredProfileName] = useState('');
  const [storedProfilePhone, setStoredProfilePhone] = useState('');
  const [storedAvatarUrl, setStoredAvatarUrl] = useState<string | null>(null);
  const userId = session?.user.id;
  const emailFallback = session?.user.email?.split('@')[0]?.replace(/[._-]+/g, ' ').trim();
  const profileName =
    storedProfileName.trim() ||
    ((session?.user.user_metadata?.full_name ?? session?.user.user_metadata?.name) as string | undefined)?.trim() ||
    emailFallback ||
    'Usuário';
  const [showSplash, setShowSplash] = useState(
    () => typeof window !== 'undefined' && window.matchMedia('(max-width: 760px)').matches,
  );
  const [activePage, setActivePage] = useState<AppPage>('dashboard');
  const [pendingPage, setPendingPage] = useState<AppPage>('dashboard');
  const [previousPage, setPreviousPage] = useState<AppPage>('dashboard');
  const [, startNavTransition] = useTransition();
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    if (typeof window === 'undefined') return 'light';
    return localStorage.getItem('rubylife-theme') === 'dark' ? 'dark' : 'light';
  });
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    try { localStorage.setItem('rubylife-theme', theme); } catch { /* storage bloqueado */ }
  }, [theme]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [customCategories, setCustomCategories] = useState<CustomCategory[]>([]);
  const [goals, setGoals] = useState<Goal[]>([]);
  const [friendInvitations, setFriendInvitations] = useState<FriendshipInvitation[]>([]);
  const [publicUsers, setPublicUsers] = useState<FriendUser[]>([]);
  const [notifications, setNotifications] = useState<AppNotification[]>([]);
  const [friendsTab, setFriendsTab] = useState<'search' | 'received'>('search');
  const [sharedTransactions, setSharedTransactions] = useState<SharedTransactionRequest[]>([]);
  const [referenceDate, setReferenceDate] = useState(() => new Date(2026, 5, 1));
  const [launchOpen, setLaunchOpen] = useState(false);
  const [launchType, setLaunchType] = useState<TransactionType>('expense');
  const [shoppingCreateSignal, setShoppingCreateSignal] = useState(0);
  const [friendSearchSignal, setFriendSearchSignal] = useState(0);
  const [friendBackSignal, setFriendBackSignal] = useState(0);
  const [activeFriendThreadName, setActiveFriendThreadName] = useState<string | null>(null);
  const [activeShoppingListName, setActiveShoppingListName] = useState<string | null>(null);
  const [shoppingBackSignal, setShoppingBackSignal] = useState(0);
  const [importOpen, setImportOpen] = useState(false);
  const [accountOpen, setAccountOpen] = useState(false);
  const [categoryOpen, setCategoryOpen] = useState(false);
  const [editingCategory, setEditingCategory] = useState<CategoryItem | null>(null);
  const [deletingCategory, setDeletingCategory] = useState<CategoryItem | null>(null);
  const [settleTarget, setSettleTarget] = useState<Transaction | null>(null);
  const [editingTransaction, setEditingTransaction] = useState<Transaction | null>(null);
  const [deletingTransaction, setDeletingTransaction] = useState<Transaction | null>(null);
  const [longPressTransaction, setLongPressTransaction] = useState<Transaction | null>(null);
  const [goalOpen, setGoalOpen] = useState(false);
  const [editingGoal, setEditingGoal] = useState<Goal | null>(null);
  const [deletingGoal, setDeletingGoal] = useState<Goal | null>(null);
  const [goalMovementTarget, setGoalMovementTarget] = useState<{ goal: Goal; type: GoalMovementType } | null>(null);
  const [editingAccount, setEditingAccount] = useState<Account | null>(null);
  const [selectedTxIds, setSelectedTxIds] = useState<Set<string>>(() => new Set());
  const [confirmDialog, setConfirmDialog] = useState<ConfirmConfig | null>(null);
  const [categoryFilter, setCategoryFilter] = useState('all');
  const [typeFilter, setTypeFilter] = useState<TransactionTypeFilter>('all');
  const [transactionSummaryMode, setTransactionSummaryMode] = useState<TransactionSummaryMode>('balance');
  const [dateFilter, setDateFilter] = useState('');
  const [filterOpen, setFilterOpen] = useState(false);
  const [draftFilters, setDraftFilters] = useState<ReportFilters>({ date: '', category: 'all', type: 'all' });
  const filterControlRef = useRef<HTMLDivElement>(null);
  const [accountFilters, setAccountFilters] = useState<AccountFilters>({ search: '', type: 'all' });
  const [draftAccountFilters, setDraftAccountFilters] = useState<AccountFilters>({ search: '', type: 'all' });
  const [accountFilterOpen, setAccountFilterOpen] = useState(false);
  const accountFilterControlRef = useRef<HTMLDivElement>(null);
  const [categoryPageFilters, setCategoryPageFilters] = useState<CategoryFilters>({ search: '', type: 'all' });
  const [draftCategoryPageFilters, setDraftCategoryPageFilters] = useState<CategoryFilters>({ search: '', type: 'all' });
  const [categoryPageFilterOpen, setCategoryPageFilterOpen] = useState(false);
  const categoryPageFilterControlRef = useRef<HTMLDivElement>(null);
  const [goalFilters, setGoalFilters] = useState<GoalFilters>({ search: '', status: 'all' });
  const [draftGoalFilters, setDraftGoalFilters] = useState<GoalFilters>({ search: '', status: 'all' });
  const [goalFilterOpen, setGoalFilterOpen] = useState(false);
  const goalFilterControlRef = useRef<HTMLDivElement>(null);
  const mainContentRef = useRef<HTMLElement>(null);

  useEffect(() => {
    if (!showSplash) return;
    const timer = window.setTimeout(() => setShowSplash(false), 3500);
    return () => window.clearTimeout(timer);
  }, [showSplash]);

  // Sessão do Supabase: lê a atual e escuta login/logout.
  useEffect(() => {
    let settled = false;
    // Garante que o app saia do splash mesmo se getSession falhar/demorar
    // (ex.: Safari mobile com storage bloqueado).
    const fallback = window.setTimeout(() => { if (!settled) setAuthReady(true); }, 4000);
    supabase.auth.getSession()
      .then(({ data }) => { setSession(data.session); })
      .catch((error) => { console.error('Falha ao obter sessão:', error); })
      .finally(() => { settled = true; window.clearTimeout(fallback); setAuthReady(true); });
    const { data: sub } = supabase.auth.onAuthStateChange((event, nextSession) => {
      if (event === 'PASSWORD_RECOVERY') setRecoveryMode(true);
      setSession(nextSession);
      setAuthReady(true);
    });
    return () => { window.clearTimeout(fallback); sub.subscription.unsubscribe(); };
  }, []);

  // Carrega os dados do usuário quando há sessão; limpa ao sair.
  useEffect(() => {
    if (!userId) {
      setStoredProfileName('');
      setStoredProfilePhone('');
      setStoredAvatarUrl(null);
      setAccounts([]);
      setCustomCategories([]);
      setTransactions([]);
      setGoals([]);
      setFriendInvitations([]);
      setPublicUsers([]);
      setNotifications([]);
      setSharedTransactions([]);
      return;
    }
    let cancelled = false;
    setDataLoading(true);
    setSyncError(null);
    loadAll()
      .then((data) => {
        if (cancelled) return;
        setStoredProfileName(data.profileName ?? '');
        setStoredProfilePhone(data.profilePhone ?? '');
        setStoredAvatarUrl(data.profileAvatarUrl ?? null);
        setAccounts(data.accounts);
        setCustomCategories(data.categories);
        setTransactions(data.transactions);
        setGoals(data.goals);
        setFriendInvitations(loadStoredFriendInvitations(userId));
        setPublicUsers(loadPublicUsers());
        setNotifications(loadNotifications());
        setSharedTransactions(loadStoredSharedTransactions());
      })
      .catch((error) => {
        console.error('Falha ao carregar dados do Supabase:', error);
        if (!cancelled) setSyncError('Não foi possível carregar seus dados. Recarregue a página.');
      })
      .finally(() => {
        if (!cancelled) setDataLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [userId]);

  function reportSyncError(scope: string, error: unknown) {
    console.error(`Falha ao salvar ${scope} no Supabase:`, error);
    setSyncError(`Não foi possível salvar (${scope}). Verifique sua conexão.`);
  }

  function scrollContentToTop() {
    const main = mainContentRef.current;
    if (main) {
      main.scrollTop = 0;
      main.scrollLeft = 0;
      main.scrollTo({ top: 0, left: 0, behavior: 'auto' });
    }

    document.documentElement.scrollTop = 0;
    document.body.scrollTop = 0;
    window.scrollTo({ top: 0, left: 0, behavior: 'auto' });

    window.requestAnimationFrame(() => {
      mainContentRef.current?.scrollTo({ top: 0, left: 0, behavior: 'auto' });
      window.scrollTo({ top: 0, left: 0, behavior: 'auto' });
    });
  }

  function handleNavigate(page: AppPage) {
    // Update tab bar highlight immediately — no waiting for the heavy re-render
    setPendingPage(page);
    startNavTransition(() => {
      setActivePage((current) => {
        if (current !== 'notifications' && current !== 'help' && current !== 'profile') {
          setPreviousPage(current);
        }
        if (current === page) scrollContentToTop();
        return page;
      });
      setPendingPage(page);
    });
  }

  function navigateFromUrl(rawUrl: string) {
    const url = new URL(rawUrl, window.location.origin);
    const path = url.pathname.replace(/\/$/, '');
    if (path === '/amigos' || url.searchParams.get('page') === 'friends') {
      setActivePage('friends');
      setFriendsTab(url.searchParams.get('aba') === 'solicitacoes' || url.searchParams.get('tab') === 'received' ? 'received' : 'search');
    } else if (path === '/listas-compras' || path.startsWith('/listas-compras/') || url.searchParams.get('page') === 'shopping') {
      setActivePage('shopping');
    } else if (path === '/transacoes-compartilhadas' || url.searchParams.get('page') === 'transactions') {
      setActivePage('friends');
    } else if (path === '/perfil') {
      setActivePage('profile');
    }
  }

  useEffect(() => {
    navigateFromUrl(window.location.href);
    function onServiceWorkerMessage(event: MessageEvent) {
      if (event.data?.type === 'RUBYLIFE_PUSH_NAVIGATE' && event.data.url) navigateFromUrl(event.data.url);
    }
    navigator.serviceWorker?.addEventListener('message', onServiceWorkerMessage);
    return () => navigator.serviceWorker?.removeEventListener('message', onServiceWorkerMessage);
  }, []);

  useEffect(() => {
    requestAnimationFrame(() => scrollContentToTop());
  }, [activePage]);

  useEffect(() => {
    if (!filterOpen) return;

    function closeOnOutsideClick(event: PointerEvent) {
      const target = event.target as Element | null;
      if (!filterControlRef.current?.contains(target) && !target?.closest('.filter-popover') && !target?.closest('.mobile-tx-quick-btn--filter')) {
        setFilterOpen(false);
      }
    }

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === 'Escape') setFilterOpen(false);
    }

    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [filterOpen]);

  useEffect(() => {
    if (!accountFilterOpen) return;

    function closeOnOutsideClick(event: PointerEvent) {
      const target = event.target as Element;
      if (!accountFilterControlRef.current?.contains(target) && !target.closest('[data-account-filter-trigger]')) {
        setAccountFilterOpen(false);
      }
    }

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === 'Escape') setAccountFilterOpen(false);
    }

    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [accountFilterOpen]);

  useEffect(() => {
    if (!categoryPageFilterOpen) return;

    function closeOnOutsideClick(event: PointerEvent) {
      const target = event.target as Element;
      if (!categoryPageFilterControlRef.current?.contains(target) && !target.closest('[data-category-filter-trigger]')) {
        setCategoryPageFilterOpen(false);
      }
    }

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === 'Escape') setCategoryPageFilterOpen(false);
    }

    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [categoryPageFilterOpen]);

  useEffect(() => {
    if (!goalFilterOpen) return;

    function closeOnOutsideClick(event: PointerEvent) {
      const target = event.target as Element;
      if (!goalFilterControlRef.current?.contains(target) && !target.closest('[data-goal-filter-trigger]')) {
        setGoalFilterOpen(false);
      }
    }

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === 'Escape') setGoalFilterOpen(false);
    }

    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [goalFilterOpen]);

  // Atualização otimista: aplica no estado na hora e sincroniza com o
  // Supabase em segundo plano (só o que mudou).
  function commitTransactions(next: Transaction[]) {
    const prev = transactions;
    setTransactions(next);
    if (userId) void syncTransactions(userId, prev, next).catch((error) => reportSyncError('transações', error));
  }

  function commitAccounts(next: Account[]) {
    const prev = accounts;
    setAccounts(next);
    if (userId) void syncAccounts(userId, prev, next).catch((error) => reportSyncError('contas', error));
  }

  function commitCustomCategories(next: CustomCategory[]) {
    const prev = customCategories;
    setCustomCategories(next);
    if (userId) void syncCategories(userId, prev, next).catch((error) => reportSyncError('categorias', error));
  }

  function commitGoals(next: Goal[]) {
    const prev = goals;
    setGoals(next);
    if (userId) void syncGoals(userId, prev, next).catch((error) => reportSyncError('metas', error));
  }

  function commitFriendInvitations(next: FriendshipInvitation[]) {
    setFriendInvitations(next);
    storeFriendInvitations(next);
  }

  function commitNotifications(next: AppNotification[]) {
    setNotifications(next);
    storeNotifications(next);
  }

  function commitSharedTransactions(next: SharedTransactionRequest[]) {
    setSharedTransactions(next);
    storeSharedTransactions(next);
  }

  function notifyUser(input: Parameters<typeof sendPushNotification>[0]) {
    void sendPushNotification(input).catch((error) => {
      console.warn('Falha ao enviar push notification:', error);
    });
  }

  const currentFriendUser = useMemo<FriendUser>(() => ({
    id: userId ?? 'current-user',
    publicFriendId: userId ? publicFriendIdForUser(userId) : '000000',
    name: profileName,
    email: session?.user.email ?? '',
    avatarUrl: storedAvatarUrl,
  }), [profileName, session?.user.email, storedAvatarUrl, userId]);

  useEffect(() => {
    if (!userId) return;
    setPublicUsers(upsertPublicUser(currentFriendUser));
  }, [currentFriendUser, userId]);
  useEffect(() => {
    if (!userId || dataLoading) return;
    const cleanKey = `rubylife-demo-cleaned-v2-${userId}`;
    if (localStorage.getItem(cleanKey)) return;

    for (const key of Object.keys(localStorage)) {
      if (key.startsWith('rubylife-demo-social-seed-v1')) localStorage.removeItem(key);
    }

    const cleanUsers = loadPublicUsers().filter((u) => !u.id.startsWith('demo-'));
    storePublicUsers(cleanUsers);
    setPublicUsers(cleanUsers);

    const cleanInvitations = loadStoredFriendInvitations(userId).filter((inv) => !inv.id.startsWith('demo-'));
    storeFriendInvitations(cleanInvitations);
    setFriendInvitations(cleanInvitations);

    const cleanShared = loadStoredSharedTransactions().filter((tx) => !tx.id.startsWith('demo-'));
    storeSharedTransactions(cleanShared);
    setSharedTransactions(cleanShared);

    const cleanNotifications = loadNotifications().filter((n) => !n.id.startsWith('demo-'));
    storeNotifications(cleanNotifications);
    setNotifications(cleanNotifications);

    try {
      const raw = localStorage.getItem(SHOPPING_LISTS_STORAGE_KEY);
      if (raw) {
        const lists = JSON.parse(raw) as Array<{ id: string }>;
        const cleaned = lists.filter((l) => !l.id.startsWith('demo-'));
        localStorage.setItem(SHOPPING_LISTS_STORAGE_KEY, JSON.stringify(cleaned));
        window.dispatchEvent(new CustomEvent('rubylife-shopping-lists-change'));
      }
    } catch {
      // armazenamento local indisponível
    }

    localStorage.setItem(cleanKey, '1');
  }, [currentFriendUser, dataLoading, userId]);

  const friendDirectory = useMemo<FriendUser[]>(() => {
    const byId = new Map<string, FriendUser>();
    [currentFriendUser, ...publicUsers].forEach((user) => { if (user.id) byId.set(user.id, user); });
    return [...byId.values()];
  }, [currentFriendUser, publicUsers]);

  function sendFriendInvitation(target: FriendUser) {
    if (!userId) return 'Sessão inválida para enviar convite.';
    const blockingMessage = findBlockingFriendInvitation(userId, target, friendInvitations);
    if (blockingMessage) return blockingMessage;
    const invitation: FriendshipInvitation = {
      id: crypto.randomUUID(),
      requesterId: userId,
      receiverId: target.id,
      status: 'pending',
      createdAt: new Date().toISOString(),
    };
    commitFriendInvitations([...friendInvitations, invitation]);
    commitNotifications([
      ...notifications,
      {
        id: crypto.randomUUID(),
        userId: target.id,
        type: 'friend_invite',
        message: `${currentFriendUser.name} enviou um convite de amizade para você.`,
        createdAt: new Date().toISOString(),
      },
    ]);
    notifyUser({ userId: target.id, title: 'Novo convite de amizade', body: currentFriendUser.name + ' enviou um convite de amizade para você.', url: '/amigos?aba=solicitacoes', type: 'friend_invite' });
    return null;
  }

  function acceptFriendInvitation(invitationId: string) {
    const invitation = friendInvitations.find((item) => item.id === invitationId);
    if (!invitation || invitation.receiverId !== currentFriendUser.id || invitation.status !== 'pending') return;
    commitFriendInvitations(friendInvitations.map((item) => (
      item.id === invitationId ? { ...item, status: 'accepted', respondedAt: new Date().toISOString() } : item
    )));
    commitNotifications([
      ...notifications,
      {
        id: crypto.randomUUID(),
        userId: invitation.requesterId,
        type: 'friend_accepted',
        message: `${currentFriendUser.name} aceitou seu convite de amizade.`,
        createdAt: new Date().toISOString(),
      },
    ]);
    notifyUser({ userId: invitation.requesterId, title: 'Convite aceito', body: currentFriendUser.name + ' aceitou seu convite de amizade.', url: '/amigos', type: 'friend_accepted' });
  }

  function declineFriendInvitation(invitationId: string) {
    commitFriendInvitations(friendInvitations.map((invitation) => (
      invitation.id === invitationId && invitation.receiverId === currentFriendUser.id && invitation.status === 'pending'
        ? { ...invitation, status: 'declined', respondedAt: new Date().toISOString() }
        : invitation
    )));
  }

  function removeFriend(friendId: string) {
    commitFriendInvitations(friendInvitations.filter((invitation) => !(
      invitation.status === 'accepted' && invitationConnectsUsers(invitation, currentFriendUser.id, friendId)
    )));
  }

  function openFriendRequestsFromNotification() {
    setFriendsTab('received');
    setActivePage('friends');
    commitNotifications(notifications.map((notification) => (
      notification.userId === currentFriendUser.id && notification.type === 'friend_invite' ? { ...notification, read: true } : notification
    )));
  }
  const acceptedFriends = useMemo(() => {
    const byId = new Map(friendDirectory.map((friend) => [friend.id, friend] as const));
    return friendInvitations
      .filter((invitation) => invitation.status === 'accepted' && (invitation.requesterId === currentFriendUser.id || invitation.receiverId === currentFriendUser.id))
      .map((invitation) => byId.get(invitation.requesterId === currentFriendUser.id ? invitation.receiverId : invitation.requesterId))
      .filter((friend): friend is FriendUser => Boolean(friend));
  }, [currentFriendUser.id, friendDirectory, friendInvitations]);

  function createSharedTransaction(form: SharedTransactionForm) {
    if (!userId) return 'Sessão inválida para criar transação compartilhada.';
    const friend = acceptedFriends.find((candidate) => candidate.id === form.friendId);
    if (!friend || !canShareWithUser(userId, friend.id, friendInvitations)) return 'Apenas amigos conectados podem receber transações compartilhadas.';
    const amount = parseAmount(form.amount);
    if (!form.description.trim() || amount <= 0 || !form.dueDate || !form.category) return 'Preencha descrição, valor, vencimento e categoria.';
    const description = form.description.trim().slice(0, TRANSACTION_DESCRIPTION_MAX_LENGTH);
    const requestId = crypto.randomUUID();
    commitSharedTransactions([
      ...sharedTransactions,
      {
        id: requestId,
        creatorId: userId,
        receiverId: friend.id,
        type: form.type,
        description,
        amount,
        dueDate: form.dueDate,
        category: form.category,
        note: form.note.trim() || undefined,
        status: 'pending',
        createdAt: new Date().toISOString(),
      },
    ]);
    commitNotifications([
      ...notifications,
      {
        id: crypto.randomUUID(),
        userId: friend.id,
        type: 'shared_created',
        message: `${currentFriendUser.name} criou uma ${form.type === 'income' ? 'receita' : 'despesa'} para você aprovar: ${description}.`,
        createdAt: new Date().toISOString(),
      },
    ]);
    notifyUser({
      userId: friend.id,
      title: form.type === 'income' ? 'Nova receita para aprovar' : 'Nova despesa para aprovar',
      body: `${currentFriendUser.name} criou uma ${form.type === 'income' ? 'receita' : 'despesa'} de ${formatCurrency(amount)} para você aprovar.`,
      url: '/transacoes-compartilhadas?aba=pendentes',
      type: form.type === 'income' ? 'shared_income_created' : 'shared_expense_created',
      data: { requestId },
    });
    return null;
  }
  function approveSharedTransaction(requestId: string) {
    if (!userId) return;
    const request = sharedTransactions.find((item) => item.id === requestId);
    if (!request || request.receiverId !== userId || request.status !== 'pending') return;
    const creator = friendDirectory.find((item) => item.id === request.creatorId);
    const newTransactionId = crypto.randomUUID();
    const newTransaction: Transaction = {
      id: newTransactionId,
      groupId: crypto.randomUUID(),
      description: request.description,
      category: request.category,
      amount: request.amount,
      type: request.type,
      dueDate: request.dueDate,
      recurrence: 'single',
      status: 'open',
      notes: request.note,
      sharedRequestId: request.id,
      sharedCreatedBy: request.creatorId,
      sharedCreatedByName: creator?.name ?? 'Amigo',
    };
    commitTransactions([...transactions, newTransaction]);
    commitSharedTransactions(sharedTransactions.map((item) => (
      item.id === requestId
        ? { ...item, status: 'approved', respondedAt: new Date().toISOString(), approvedTransactionId: newTransactionId }
        : item
    )));
    commitNotifications([
      ...notifications,
      {
        id: crypto.randomUUID(),
        userId: request.creatorId,
        type: 'shared_approved',
        message: `${currentFriendUser.name} aprovou a transação: ${request.description}.`,
        createdAt: new Date().toISOString(),
      },
    ]);
    notifyUser({ userId: request.creatorId, title: 'Transação aprovada', body: currentFriendUser.name + ' aprovou a transação compartilhada.', url: '/transacoes-compartilhadas?aba=pendentes', type: 'shared_approved', data: { requestId } });
  }

  function declineSharedTransaction(requestId: string, reason: string) {
    if (!userId) return;
    const declineReason = reason.trim();
    if (!declineReason) return;
    const request = sharedTransactions.find((item) => item.id === requestId);
    if (!request || request.receiverId !== userId || request.status !== 'pending') return;
    commitSharedTransactions(sharedTransactions.map((item) => (
      item.id === requestId ? { ...item, status: 'declined', declineReason, respondedAt: new Date().toISOString() } : item
    )));
    commitNotifications([
      ...notifications,
      {
        id: crypto.randomUUID(),
        userId: request.creatorId,
        type: 'shared_declined',
        message: `${currentFriendUser.name} recusou a transação: ${request.description}. Motivo: ${declineReason}`,
        createdAt: new Date().toISOString(),
      },
    ]);
    notifyUser({ userId: request.creatorId, title: 'Transação recusada', body: `${currentFriendUser.name} recusou a transação compartilhada. Motivo: ${declineReason}`, url: '/transacoes-compartilhadas?aba=pendentes', type: 'shared_declined', data: { requestId } });
  }

  function cancelSharedTransaction(requestId: string) {
    if (!userId) return;
    const request = sharedTransactions.find((item) => item.id === requestId);
    if (!request || request.creatorId !== userId || request.status !== 'pending') return;
    commitSharedTransactions(sharedTransactions.map((item) => (
      item.id === requestId ? { ...item, status: 'cancelled', respondedAt: new Date().toISOString() } : item
    )));
    commitNotifications([
      ...notifications,
      {
        id: crypto.randomUUID(),
        userId: request.receiverId,
        type: 'shared_cancelled',
        message: `${currentFriendUser.name} cancelou a transação: ${request.description}.`,
        createdAt: new Date().toISOString(),
      },
    ]);
  }

  const currentMonth = monthKey(referenceDate);
  const categoryItems = useMemo<CategoryItem[]>(() => (
    customCategories.filter((category) => !category.deleted).map((category, index) => ({
      id: category.id,
      name: category.name,
      kind: category.kind,
      color: category.color ?? ACCOUNT_COLORS[index % ACCOUNT_COLORS.length] ?? '#1B99D8',
      icon: category.icon ?? 'tag',
    }))
  ), [customCategories]);
  const categoryGroups = useMemo(() => CATEGORY_GROUPS.map((group) => ({
    ...group,
    categories: categoryItems.filter((category) => category.kind === group.key).map((category) => category.name),
  })), [categoryItems]);
  const incomeCategories = categoryGroups.find((group) => group.key === 'income')?.categories ?? [];
  const expenseCategories = categoryGroups.find((group) => group.key === 'expense')?.categories ?? [];
  const transactionCategoryOptions = [...new Set([...incomeCategories, ...expenseCategories])];
  const categoryLookup = useMemo(() => {
    const map = new Map<string, CategoryItem>();
    categoryItems.forEach((category) => map.set(`${category.kind}:${category.name}`, category));
    return map;
  }, [categoryItems]);
  const monthItems = useMemo(() =>
    transactions
      .filter((item) => (dateFilter ? item.dueDate === dateFilter : item.dueDate.slice(0, 7) === currentMonth))
      .filter((item) => categoryFilter === 'all' || item.category === categoryFilter)
      .filter((item) => typeFilter === 'all' || item.type === typeFilter)
      .sort((a, b) => a.dueDate.localeCompare(b.dueDate)),
    [transactions, dateFilter, currentMonth, categoryFilter, typeFilter],
  );

  const selectedInView = useMemo(() => monthItems.filter((item) => selectedTxIds.has(item.id)), [monthItems, selectedTxIds]);
  const selectedOpenInView = useMemo(() => selectedInView.filter((item) => item.status === 'open'), [selectedInView]);
  const selectedSettledInView = useMemo(() => selectedInView.filter((item) => item.status === 'settled'), [selectedInView]);
  const selectedOpenWithoutAccount = useMemo(() => selectedOpenInView.filter((item) => !accounts.some((account) => account.id === item.accountId || account.name === item.account)), [selectedOpenInView, accounts]);
  const allSelected = monthItems.length > 0 && selectedInView.length === monthItems.length;

  function toggleSelectTx(id: string) {
    setSelectedTxIds((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  function toggleSelectAllTx() {
    setSelectedTxIds((current) => {
      if (monthItems.length > 0 && monthItems.every((item) => current.has(item.id))) {
        const next = new Set(current);
        monthItems.forEach((item) => next.delete(item.id));
        return next;
      }
      const next = new Set(current);
      monthItems.forEach((item) => next.add(item.id));
      return next;
    });
  }

  function bulkSettleSelectedTransactions() {
    const selectedIds = new Set(selectedOpenInView.map((item) => item.id));
    const settledAt = new Date().toISOString().slice(0, 10);

    commitTransactions(transactions.map((item) => {
      if (!selectedIds.has(item.id) || item.status === 'settled') return item;
      const account = accounts.find((candidate) => candidate.id === item.accountId || candidate.name === item.account);
      if (!account) return item;
      return {
        ...item,
        status: 'settled',
        settledAt,
        settledAmount: item.amount,
        accountId: account.id,
        account: account.name,
      };
    }));
    setSelectedTxIds(new Set());
  }

  function markTransactionAsPending(targetId: string) {
    commitTransactions(transactions.map((item) => {
      if (item.id !== targetId) return item;
      const { settledAt: _settledAt, settledAmount: _settledAmount, ...rest } = item;
      return { ...rest, status: 'open' };
    }));
  }

  function bulkMarkSelectedTransactionsAsPending() {
    const selectedIds = new Set(selectedSettledInView.map((item) => item.id));
    commitTransactions(transactions.map((item) => {
      if (!selectedIds.has(item.id)) return item;
      const { settledAt: _settledAt, settledAmount: _settledAmount, ...rest } = item;
      return { ...rest, status: 'open' };
    }));
    setSelectedTxIds(new Set());
  }


  const activeFilterCount = Number(categoryFilter !== 'all') + Number(typeFilter !== 'all') + Number(Boolean(dateFilter));

  const summary = useMemo(() => {
    const income = monthItems.filter((item) => item.type === 'income').reduce((sum, item) => sum + item.amount, 0);
    const expense = monthItems.filter((item) => item.type === 'expense').reduce((sum, item) => sum + item.amount, 0);
    const pendingIncome = monthItems.filter((item) => item.type === 'income' && item.status === 'open').reduce((sum, item) => sum + item.amount, 0);
    const pendingExpense = monthItems.filter((item) => item.type === 'expense' && item.status === 'open').reduce((sum, item) => sum + item.amount, 0);
    return { income, expense, balance: income - expense, pendingIncome, pendingExpense };
  }, [monthItems]);

  const mobileCategoryBreakdown = useMemo(() => {
    const expenseMap = new Map<string, { total: number; color?: string }>();
    const allMap = new Map<string, { total: number; color?: string }>();
    let expenseMax = 0;
    let allMax = 0;
    for (const item of monthItems) {
      const meta = categoryLookup.get(`${item.type}:${item.category}`);
      const color = meta?.color ?? (item.type === 'income' ? '#10b981' : '#ef4444');
      const currentAll = allMap.get(item.category) ?? { total: 0, color };
      currentAll.total += item.amount;
      if (currentAll.total > allMax) allMax = currentAll.total;
      allMap.set(item.category, currentAll);

      if (item.type === 'expense') {
        const currentExp = expenseMap.get(item.category) ?? { total: 0, color };
        currentExp.total += item.amount;
        if (currentExp.total > expenseMax) expenseMax = currentExp.total;
        expenseMap.set(item.category, currentExp);
      }
    }
    const expenseList = Array.from(expenseMap.entries()).map(([name, val]) => ({ name, ...val })).sort((a, b) => b.total - a.total).slice(0, 5);
    if (expenseList.length > 0) return { list: expenseList, max: expenseMax || 1 };
    const allList = Array.from(allMap.entries()).map(([name, val]) => ({ name, ...val })).sort((a, b) => b.total - a.total).slice(0, 5);
    return { list: allList, max: allMax || 1 };
  }, [monthItems, categoryLookup]);


  const accountBalances = useMemo(() => {
    const balances = Object.fromEntries(accounts.map((account) => [account.id, account.initialBalance])) as Record<string, number>;
    transactions.filter((item) => item.status === 'settled').forEach((item) => {
      const account = accounts.find((candidate) => candidate.id === item.accountId || candidate.name === item.account);
      if (!account) return;
      const settledAmount = item.settledAmount ?? item.amount;
      balances[account.id] = (balances[account.id] ?? 0) + (item.type === 'income' ? settledAmount : -settledAmount);
    });
    return balances;
  }, [accounts, transactions]);

  function openFilters() {
    setDraftFilters({
      date: dateFilter,
      category: categoryFilter,
      type: typeFilter,
    });
    setFilterOpen((open) => !open);
  }

  function applyFilters() {
    setDateFilter(draftFilters.date);
    setCategoryFilter(draftFilters.category);
    setTypeFilter(draftFilters.type);
    setFilterOpen(false);
  }

  function openAccountFilters() {
    setDraftAccountFilters(accountFilters);
    setAccountFilterOpen((open) => !open);
  }

  function applyAccountFilters() {
    setAccountFilters(draftAccountFilters);
    setAccountFilterOpen(false);
  }

  function openCategoryPageFilters() {
    setDraftCategoryPageFilters(categoryPageFilters);
    setCategoryPageFilterOpen((open) => !open);
  }

  function applyCategoryPageFilters() {
    setCategoryPageFilters(draftCategoryPageFilters);
    setCategoryPageFilterOpen(false);
  }

  function openGoalFilters() {
    setDraftGoalFilters(goalFilters);
    setGoalFilterOpen((open) => !open);
  }

  function applyGoalFilters() {
    setGoalFilters(draftGoalFilters);
    setGoalFilterOpen(false);
  }

  if (showSplash || !authReady) return <SplashScreen />;

  if (recoveryMode) return <UpdatePasswordScreen onDone={() => setRecoveryMode(false)} />;

  if (!session) return <LoginScreen />;

  if (dataLoading) return <SplashScreen />;

  return (
    <div className="app-layout">
      {syncError && (
        <div role="alert" style={{ position: 'fixed', top: 0, left: 0, right: 0, zIndex: 9999, background: '#DC2626', color: '#fff', padding: '8px 16px', textAlign: 'center', fontSize: 14, display: 'flex', justifyContent: 'center', alignItems: 'center', gap: 12 }}>
          <span>{syncError}</span>
          <button type="button" onClick={() => setSyncError(null)} style={{ background: 'transparent', border: '1px solid rgba(255,255,255,0.6)', color: '#fff', borderRadius: 6, padding: '2px 8px', cursor: 'pointer' }}>Fechar</button>
        </div>
      )}
      <Topbar activePage={activePage} userName={profileName} userAvatarUrl={storedAvatarUrl} theme={theme} notificationCount={notifications.filter((notification) => notification.userId === currentFriendUser.id && !notification.read).length} friendDetailName={activePage === 'friends' ? activeFriendThreadName : null} onFriendDetailBack={() => setFriendBackSignal((value) => value + 1)} shoppingDetailName={activePage === 'shopping' ? activeShoppingListName : null} onShoppingDetailBack={() => setShoppingBackSignal((value) => value + 1)} onOpenFriendRequests={openFriendRequestsFromNotification} onOpenFriendSearch={() => { handleNavigate('friends'); setFriendSearchSignal((value) => value + 1); }} onOpenShoppingCreate={() => { handleNavigate('shopping'); setShoppingCreateSignal((value) => value + 1); }} onOpenGoalCreate={() => { setEditingGoal(null); setGoalOpen(true); }} onGoBack={() => handleNavigate(previousPage)} onToggleTheme={() => setTheme((current) => (current === 'dark' ? 'light' : 'dark'))} onNavigate={handleNavigate} onLogout={() => supabase.auth.signOut()} onOpenTransactionFilters={openFilters} onImportTransactions={() => setImportOpen(true)} onOpenAccountFilters={openAccountFilters} accountFilterOpen={accountFilterOpen} accountActiveFilterCount={Number(Boolean(accountFilters.search.trim())) + Number(accountFilters.type !== 'all')} onOpenCategoryFilters={openCategoryPageFilters} categoryFilterOpen={categoryPageFilterOpen} categoryActiveFilterCount={Number(Boolean(categoryPageFilters.search.trim())) + Number(categoryPageFilters.type !== 'all')} onOpenGoalFilters={openGoalFilters} goalFilterOpen={goalFilterOpen} goalActiveFilterCount={Number(Boolean(goalFilters.search.trim())) + Number(goalFilters.status !== 'all')} />
      <div className="content-layout">
        <Sidebar activePage={pendingPage} onNavigate={handleNavigate} />
        <MobileTabBar
          activePage={pendingPage}
          onNavigate={handleNavigate}
          onNewIncome={() => { setLaunchType('income'); setLaunchOpen(true); }}
          onNewExpense={() => { setLaunchType('expense'); setLaunchOpen(true); }}
          onNewGoal={() => { setEditingGoal(null); setGoalOpen(true); }}
          onNewCategory={() => { setEditingCategory(null); setCategoryOpen(true); }}
          onNewSharedTransaction={() => { handleNavigate('friends'); }}
          onNewShoppingList={() => { handleNavigate('shopping'); setShoppingCreateSignal((value) => value + 1); }}
        />
        <main ref={mainContentRef} className={`main-content main-content--${activePage}`}>
          {activePage === 'dashboard' ? (
            <DashboardPage
              transactions={transactions}
              referenceDate={referenceDate}
              onChangeDate={setReferenceDate}
              onNavigate={handleNavigate}
            />
          ) : activePage === 'transactions' ? (
            <>
              <div className="tx-desktop-view">
                <section className="page-header page-header-split">
                  <div className="page-header-left">
                    <h1 className="page-title">Transações</h1>
                  </div>
                  <div className="page-header-center">
                    <MonthNavigator date={referenceDate} onChange={(next) => { setReferenceDate(next); setDateFilter(''); }} />
                  </div>
                  <div className="page-header-actions">
                    <div className="filter-control" ref={filterControlRef}>
                      <button type="button" className={`filter-trigger btn-icon-only${filterOpen ? ' active' : ''}`} onClick={openFilters} aria-expanded={filterOpen} aria-haspopup="dialog" title="Filtros">
                        <ListFilter size={16} />
                        {activeFilterCount > 0 ? <span className="filter-badge">{activeFilterCount}</span> : null}
                      </button>

                      {filterOpen ? (
                        <div className="filter-popover" role="dialog" aria-label="Filtros das transações">
                          <div className="filter-popover-header">
                            <strong>Filtrar transações</strong>
                            <button type="button" className="filter-popover-close" onClick={() => setFilterOpen(false)} aria-label="Fechar filtros"><X size={18} /></button>
                          </div>
                          <div className="filter-grid filter-grid--stacked">
                            <label className="filter-field">
                              <span>Data fixa</span>
                              <input type="date" value={draftFilters.date} onChange={(event) => setDraftFilters((current) => ({ ...current, date: event.target.value }))} />
                            </label>
                            <label className="filter-field">
                              <span>Categoria</span>
                              <select value={draftFilters.category} onChange={(event) => setDraftFilters((current) => ({ ...current, category: event.target.value }))}>
                                <option value="all">Todas</option>
                                {transactionCategoryOptions.map((category) => <option key={category} value={category}>{category}</option>)}
                              </select>
                            </label>
                            <label className="filter-field">
                              <span>Tipo</span>
                              <select value={draftFilters.type} onChange={(event) => setDraftFilters((current) => ({ ...current, type: event.target.value as TransactionTypeFilter }))}>
                                <option value="all">Todos</option>
                                <option value="income">Receita</option>
                                <option value="expense">Despesa</option>
                              </select>
                            </label>
                          </div>
                          <div className="filter-popover-actions">
                            <button type="button" className="filter-clear" onClick={() => setDraftFilters({ date: '', category: 'all', type: 'all' })}>
                              <RotateCcw size={14} /> Limpar
                            </button>
                            <button type="button" className="filter-apply" onClick={applyFilters}>Aplicar filtros</button>
                          </div>
                        </div>
                      ) : null}
                    </div>
                    <button type="button" className="page-secondary-action btn-icon-only" onClick={() => setImportOpen(true)} title="Importar Excel">
                      <Upload size={16} />
                    </button>
                    <button type="button" className="page-primary-action" onClick={() => setLaunchOpen(true)}>
                      <Plus size={16} />
                      Nova transação
                    </button>
                  </div>
                </section>

                <div className="transaction-summary-grid">
                  <TransactionSummaryCard summary={summary} activeMode={transactionSummaryMode} onModeChange={setTransactionSummaryMode} />
                  <TransactionSummaryStats summary={summary} />
                </div>

                {selectedInView.length > 0 ? (
                  <div className="bulk-bar">
                    <span className="bulk-bar-count"><strong>{selectedInView.length}</strong> {selectedInView.length === 1 ? 'selecionada' : 'selecionadas'}</span>
                    <div className="bulk-bar-actions">
                      <button type="button" className="bulk-bar-button bulk-bar-button--clear" onClick={() => setSelectedTxIds(new Set())}><X size={15} /> Limpar seleção</button>
                      <button type="button" className="bulk-bar-button bulk-bar-button--success" disabled={selectedOpenInView.length === 0 || selectedOpenWithoutAccount.length > 0} title={selectedOpenWithoutAccount.length ? 'Defina o banco ou conta antes de efetivar.' : undefined} onClick={() => setConfirmDialog({
                        title: 'Efetivado',
                        message: `Deseja efetivar ${selectedOpenInView.length} ${selectedOpenInView.length === 1 ? 'transação em aberto' : 'transações em aberto'}?`,
                        confirmLabel: 'Efetivado',
                        onConfirm: bulkSettleSelectedTransactions,
                      })}><CheckCircle2 size={15} /> Efetivado</button>
                      <button type="button" className="bulk-bar-button" disabled={selectedSettledInView.length === 0} onClick={() => setConfirmDialog({
                        title: 'Desefetivar',
                        message: `Deseja desefetivar ${selectedSettledInView.length} ${selectedSettledInView.length === 1 ? 'transação efetivada' : 'transações efetivadas'}?`,
                        confirmLabel: 'Desefetivar',
                        onConfirm: bulkMarkSelectedTransactionsAsPending,
                      })}><RotateCcw size={15} /> Desefetivar</button>
                      <button type="button" className="bulk-bar-button bulk-bar-button--danger" onClick={() => setConfirmDialog({
                        title: 'Excluir transações',
                        message: `Deseja excluir ${selectedInView.length} ${selectedInView.length === 1 ? 'transação selecionada' : 'transações selecionadas'}?`,
                        confirmLabel: 'Excluir',
                        onConfirm: () => { commitTransactions(transactions.filter((item) => !selectedTxIds.has(item.id))); setSelectedTxIds(new Set()); },
                      })}><Trash2 size={15} /> Excluir selecionadas</button>
                    </div>
                  </div>
                ) : null}

                <section className="resource-panel">
                  <div className="documents-report launches-report">
                    <div className="documents-report-head launches-report-head">
                      <span className="row-check-col"><input type="checkbox" className="row-check" checked={allSelected} onChange={toggleSelectAllTx} aria-label="Selecionar todas" /></span>
                      <span>Descrição</span>
                      <span>Categoria</span>
                      <span>Tipo</span>
                      <span>Recorrência</span>
                      <span>Vencimento</span>
                      <span>Valor previsto</span>
                      <span>Valor real</span>
                      <span>Conta</span>
                      <span>Status</span>
                      <span>Ações</span>
                    </div>
                    <div className="documents-report-body">
                      {monthItems.length === 0 ? (
                        <div className="resource-empty-state">
                          <ReceiptText size={30} />
                          <strong>Nenhuma transação encontrada</strong>
                          <p>Use o botão de nova transação para registrar receitas, despesas, parcelas ou valores fixos.</p>
                        </div>
                      ) : monthItems.map((item, index) => {
                        const showDateGroup = index === 0 || monthItems[index - 1]?.dueDate !== item.dueDate;

                        return (
                          <Fragment key={item.id}>
                            {showDateGroup ? <div className="mobile-date-separator">{formatDateGroup(item.dueDate)}</div> : null}
                            <SwipeableTransactionRow
                              item={item}
                              isSelected={selectedTxIds.has(item.id)}
                              selectionMode={selectedTxIds.size > 0}
                              onToggleSelect={() => toggleSelectTx(item.id)}
                              onEdit={() => setEditingTransaction(item)}
                              onSettle={() => setSettleTarget(item)}
                              onDelete={() => setDeletingTransaction(item)}
                              onMarkPending={() => markTransactionAsPending(item.id)}
                              onLongPress={() => setLongPressTransaction(item)}
                              categoryMeta={categoryLookup.get(`${item.type}:${item.category}`)}
                            />
                          </Fragment>
                        );
                      })}
                    </div>
                  </div>
                </section>
              </div>

              <div className="tx-mobile-view">
                <section className="mobile-tx-top-area">
                  <div className="mobile-tx-month-bar">
                    <div className="mobile-tx-month-center">
                      <MonthNavigator date={referenceDate} onChange={(next) => { setReferenceDate(next); setDateFilter(''); }} />
                    </div>
                  </div>

                  <div className="mobile-tx-summary-carousel">
                    {/* Card 1: Receitas */}
                    <div className="mobile-tx-summary-card mobile-tx-summary-card--income">
                      <div className="mobile-tx-card-row">
                        <div className="mobile-tx-card-left">
                          <span className="mobile-tx-card-icon mobile-tx-card-icon--blue"><ArrowUpRight size={22} /></span>
                          <span className="mobile-tx-card-title">Receitas</span>
                        </div>
                        <div className="mobile-tx-card-right">
                          <div className="mobile-tx-card-value">{formatCurrency(summary.income)}</div>
                        </div>
                      </div>
                      <div className="mobile-tx-card-footer">
                        <span>Recebido: <strong>{formatCurrency(summary.income - (summary.pendingIncome ?? 0))}</strong></span>
                        <span className="mobile-tx-card-subval">A receber: {formatCurrency(summary.pendingIncome ?? 0)}</span>
                      </div>
                    </div>

                    {/* Card 2: Despesas */}
                    <div className="mobile-tx-summary-card mobile-tx-summary-card--expense">
                      <div className="mobile-tx-card-row">
                        <div className="mobile-tx-card-left">
                          <span className="mobile-tx-card-icon mobile-tx-card-icon--red"><ArrowDownLeft size={22} /></span>
                          <span className="mobile-tx-card-title">Despesas</span>
                        </div>
                        <div className="mobile-tx-card-right">
                          <div className="mobile-tx-card-value">{formatCurrency(summary.expense)}</div>
                        </div>
                      </div>
                      <div className="mobile-tx-card-footer">
                        <span>Pago: <strong>{formatCurrency(summary.expense - (summary.pendingExpense ?? 0))}</strong></span>
                        <span className="mobile-tx-card-subval">A pagar: {formatCurrency(summary.pendingExpense ?? 0)}</span>
                      </div>
                    </div>

                    {/* Card 3: Saldo */}
                    <div className="mobile-tx-summary-card mobile-tx-summary-card--balance">
                      <div className="mobile-tx-card-row">
                        <div className="mobile-tx-card-left">
                          <span className="mobile-tx-card-icon mobile-tx-card-icon--gray"><CircleDollarSign size={22} /></span>
                          <span className="mobile-tx-card-title">Saldo</span>
                        </div>
                        <div className="mobile-tx-card-right">
                          <div className="mobile-tx-card-value">{formatCurrency(summary.balance)}</div>
                        </div>
                      </div>
                      <div className="mobile-tx-card-footer">
                        <span>Realizado: <strong>{formatCurrency((summary.income - (summary.pendingIncome ?? 0)) - (summary.expense - (summary.pendingExpense ?? 0)))}</strong></span>
                        <span className="mobile-tx-card-trend">
                          {summary.balance >= 0 ? <TrendingUp size={14} className="trend-up" /> : <TrendingDown size={14} className="trend-down" />}
                        </span>
                      </div>
                    </div>

                    {/* Card 4: Categorias */}
                    <div className="mobile-tx-summary-card mobile-tx-summary-card--chart">
                      <div className="mobile-tx-card-head">
                        <span className="mobile-tx-card-title">Categorias</span>
                      </div>
                      {mobileCategoryBreakdown.list.length === 0 ? (
                        <div className="mobile-tx-chart-empty">Sem transações no período</div>
                      ) : (
                        <div className="mobile-tx-col-chart">
                          {mobileCategoryBreakdown.list.map((cat) => {
                            const heightPct = Math.max(16, Math.round((cat.total / mobileCategoryBreakdown.max) * 100));
                            return (
                              <div key={cat.name} className="mobile-tx-col-item">
                                <span className="mobile-tx-col-val">{formatCurrency(cat.total)}</span>
                                <div className="mobile-tx-col-track">
                                  <div className="mobile-tx-col-fill" style={{ height: `${heightPct}%`, backgroundColor: cat.color }} />
                                </div>
                                <span className="mobile-tx-col-label" title={cat.name}>{cat.name}</span>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  </div>
                </section>

                {filterOpen ? (
                  <div className="filter-popover filter-popover--mobile" role="dialog" aria-label="Filtros das transações">
                    <div className="filter-popover-header">
                      <strong>Filtrar transações</strong>
                      <button type="button" className="filter-popover-close" onClick={() => setFilterOpen(false)} aria-label="Fechar filtros"><X size={18} /></button>
                    </div>
                    <div className="filter-grid filter-grid--stacked">
                      <label className="filter-field">
                        <span>Data fixa</span>
                        <input type="date" value={draftFilters.date} onChange={(event) => setDraftFilters((current) => ({ ...current, date: event.target.value }))} />
                      </label>
                      <label className="filter-field">
                        <span>Categoria</span>
                        <select value={draftFilters.category} onChange={(event) => setDraftFilters((current) => ({ ...current, category: event.target.value }))}>
                          <option value="all">Todas</option>
                          {transactionCategoryOptions.map((category) => <option key={category} value={category}>{category}</option>)}
                        </select>
                      </label>
                      <label className="filter-field">
                        <span>Tipo</span>
                        <select value={draftFilters.type} onChange={(event) => setDraftFilters((current) => ({ ...current, type: event.target.value as TransactionTypeFilter }))}>
                          <option value="all">Todos</option>
                          <option value="income">Receita</option>
                          <option value="expense">Despesa</option>
                        </select>
                      </label>
                    </div>
                    <div className="filter-popover-actions">
                      <button type="button" className="filter-clear" onClick={() => setDraftFilters({ date: '', category: 'all', type: 'all' })}>
                        <RotateCcw size={14} /> Limpar
                      </button>
                      <button type="button" className="filter-apply" onClick={applyFilters}>Aplicar filtros</button>
                    </div>
                  </div>
                ) : null}

                <section className="mobile-tx-bottom-sheet">
                  <div className="mobile-tx-sheet-pull-indicator">
                    <div className="pull-bar" />
                  </div>

                  <div className="mobile-tx-sheet-header">
                    <h2 className="mobile-tx-sheet-title">Lançamentos do mês</h2>
                  </div>

                  {selectedInView.length > 0 ? (
                    <div className="bulk-bar mobile-bulk-bar">
                      <span className="bulk-bar-count"><strong>{selectedInView.length}</strong> {selectedInView.length === 1 ? 'selecionada' : 'selecionadas'}</span>
                      <div className="bulk-bar-actions">
                        <button type="button" className="bulk-bar-button bulk-bar-button--clear" onClick={() => setSelectedTxIds(new Set())}><X size={15} /></button>
                        <button type="button" className="bulk-bar-button bulk-bar-button--success" disabled={selectedOpenInView.length === 0 || selectedOpenWithoutAccount.length > 0} title={selectedOpenWithoutAccount.length ? 'Defina o banco ou conta antes de efetivar.' : undefined} onClick={() => setConfirmDialog({
                          title: 'Efetivado',
                          message: `Deseja efetivar ${selectedOpenInView.length} ${selectedOpenInView.length === 1 ? 'transação em aberto' : 'transações em aberto'}?`,
                          confirmLabel: 'Efetivado',
                          onConfirm: bulkSettleSelectedTransactions,
                        })}><CheckCircle2 size={15} /></button>
                        <button type="button" className="bulk-bar-button" disabled={selectedSettledInView.length === 0} onClick={() => setConfirmDialog({
                          title: 'Desefetivar',
                          message: `Deseja desefetivar ${selectedSettledInView.length} ${selectedSettledInView.length === 1 ? 'transação efetivada' : 'transações efetivadas'}?`,
                          confirmLabel: 'Desefetivar',
                          onConfirm: bulkMarkSelectedTransactionsAsPending,
                        })}><RotateCcw size={15} /></button>
                        <button type="button" className="bulk-bar-button bulk-bar-button--danger" onClick={() => setConfirmDialog({
                          title: 'Excluir transações',
                          message: `Deseja excluir ${selectedInView.length} ${selectedInView.length === 1 ? 'transação selecionada' : 'transações selecionadas'}?`,
                          confirmLabel: 'Excluir',
                          onConfirm: () => { commitTransactions(transactions.filter((item) => !selectedTxIds.has(item.id))); setSelectedTxIds(new Set()); },
                        })}><Trash2 size={15} /></button>
                      </div>
                    </div>
                  ) : null}

                  <div className="mobile-tx-sheet-list">
                    {monthItems.length === 0 ? (
                      <div className="resource-empty-state mobile-tx-empty">
                        <ReceiptText size={36} />
                        <strong>Nenhuma transação no período</strong>
                        <p>Toque nos botões de Receita, Despesa ou no + para registrar.</p>
                      </div>
                    ) : (
                      monthItems.map((item, index) => {
                        const showDateGroup = index === 0 || monthItems[index - 1]?.dueDate !== item.dueDate;

                        return (
                          <Fragment key={item.id}>
                            {showDateGroup ? <div className="mobile-date-separator">{formatDateGroup(item.dueDate)}</div> : null}
                            <SwipeableTransactionRow
                              item={item}
                              isSelected={selectedTxIds.has(item.id)}
                              selectionMode={selectedTxIds.size > 0}
                              onToggleSelect={() => toggleSelectTx(item.id)}
                              onEdit={() => setEditingTransaction(item)}
                              onSettle={() => setSettleTarget(item)}
                              onDelete={() => setDeletingTransaction(item)}
                              onMarkPending={() => markTransactionAsPending(item.id)}
                              onLongPress={() => setLongPressTransaction(item)}
                              categoryMeta={categoryLookup.get(`${item.type}:${item.category}`)}
                            />
                          </Fragment>
                        );
                      })
                    )}
                  </div>
                </section>
              </div>
            </>
          ) : activePage === 'categories' ? (
            <CategoriesPage
              items={categoryItems}
              filters={categoryPageFilters}
              draftFilters={draftCategoryPageFilters}
              filterOpen={categoryPageFilterOpen}
              filterControlRef={categoryPageFilterControlRef}
              onOpenFilters={openCategoryPageFilters}
              onDraftFiltersChange={setDraftCategoryPageFilters}
              onApplyFilters={applyCategoryPageFilters}
              onNew={() => { setEditingCategory(null); setCategoryOpen(true); }}
              onEdit={(category) => { setEditingCategory(category); setCategoryOpen(true); }}
              onDelete={setDeletingCategory}
              onBulkDelete={(ids) => setConfirmDialog({
                title: 'Excluir categorias',
                message: `Deseja excluir ${ids.length} ${ids.length === 1 ? 'categoria selecionada' : 'categorias selecionadas'}?`,
                onConfirm: () => {
                  const idSet = new Set(ids);
                  commitCustomCategories(customCategories.filter((item) => !idSet.has(item.id)));
                },
              })}
            />
          ) : activePage === 'goals' ? (
            <GoalsPage
              goals={goals}
              filters={goalFilters}
              draftFilters={draftGoalFilters}
              filterOpen={goalFilterOpen}
              filterControlRef={goalFilterControlRef}
              onOpenFilters={openGoalFilters}
              onDraftFiltersChange={setDraftGoalFilters}
              onApplyFilters={applyGoalFilters}
              onNew={() => { setEditingGoal(null); setGoalOpen(true); }}
              onEdit={(goal) => { setEditingGoal(goal); setGoalOpen(true); }}
              onDelete={setDeletingGoal}
              onDeposit={(goal) => setGoalMovementTarget({ goal, type: 'deposit' })}
              onWithdraw={(goal) => setGoalMovementTarget({ goal, type: 'withdraw' })}
            />
          ) : activePage === 'reports' ? (
            <ReportsPage transactions={transactions} categoryLookup={categoryLookup} referenceDate={referenceDate} onChangeDate={setReferenceDate} />
          ) : activePage === 'shopping' ? (
            <ShoppingListsPage currentUser={currentFriendUser} friends={acceptedFriends} openCreateSignal={shoppingCreateSignal} backSignal={shoppingBackSignal} onDetailChange={setActiveShoppingListName} onListItemAdded={({ listId, listName, itemName, participantIds }) => { participantIds.forEach((participantId) => notifyUser({ userId: participantId, title: 'Lista de compras atualizada', body: `${currentFriendUser.name} adicionou ${itemName} à lista ${listName}.`, url: `/listas-compras/${listId}`, type: 'shopping_list_updated', data: { listId } })); }} />
          ) : activePage === 'friends' ? (
            <FriendsPage
              currentUser={currentFriendUser}
              users={friendDirectory}
              invitations={friendInvitations}
              acceptedFriends={acceptedFriends}
              sharedTransactions={sharedTransactions}
              incomeCategories={incomeCategories}
              expenseCategories={expenseCategories}
              activeTab={friendsTab}
              openSearchSignal={friendSearchSignal}
              backSignal={friendBackSignal}
              onThreadChange={setActiveFriendThreadName}
              onTabChange={setFriendsTab}
              onSendInvitation={sendFriendInvitation}
              onAcceptInvitation={(invitationId) => { acceptFriendInvitation(invitationId); setFriendsTab('search'); }}
              onDeclineInvitation={(invitationId) => declineFriendInvitation(invitationId)}
              onRemoveFriend={(friend) => setConfirmDialog({
                title: 'Remover amigo',
                message: 'Tem certeza que deseja remover essa amizade?',
                confirmLabel: 'Remover',
                onConfirm: () => removeFriend(friend.id),
              })}
              onCreateShared={createSharedTransaction}
              onApproveShared={approveSharedTransaction}
              onDeclineShared={declineSharedTransaction}
              onCancelShared={cancelSharedTransaction}
            />
          ) : activePage === 'profile' ? (
            <ProfilePage
              fullName={storedProfileName || profileName}
              email={session.user.email ?? ''}
              publicFriendId={currentFriendUser.publicFriendId}
              phone={storedProfilePhone}
              avatarUrl={storedAvatarUrl}
              onSave={async (profile) => {
                const saved = await updateProfile(profile);
                setStoredProfileName(saved.fullName);
                setStoredProfilePhone(saved.phone);
                setStoredAvatarUrl(saved.avatarUrl ?? null);
              }}
              categories={categoryItems}
              onNewCategory={() => { setEditingCategory(null); setCategoryOpen(true); }}
              onEditCategory={(category) => { setEditingCategory(category); setCategoryOpen(true); }}
            />
          ) : activePage === 'help' ? (
            <HelpPage />
          ) : activePage === 'notifications' ? (
            <NotificationsPage
              notifications={notifications}
              onClearNotifications={() => commitNotifications([])}
            />
          ) : (
            <AccountsPage
              accounts={accounts}
              balances={accountBalances}
              filters={accountFilters}
              draftFilters={draftAccountFilters}
              filterOpen={accountFilterOpen}
              filterControlRef={accountFilterControlRef}
              onOpenFilters={openAccountFilters}
              onDraftFiltersChange={setDraftAccountFilters}
              onApplyFilters={applyAccountFilters}
              onNew={() => { setEditingAccount(null); setAccountOpen(true); }}
              onEdit={(account) => { setEditingAccount(account); setAccountOpen(true); }}
              onDelete={(account) => setConfirmDialog({
                title: 'Excluir conta',
                message: <>Deseja excluir a conta <strong>{account.name}</strong>? As transações já lançadas serão mantidas.</>,
                onConfirm: () => commitAccounts(accounts.filter((item) => item.id !== account.id)),
              })}
              onBulkDelete={(ids) => setConfirmDialog({
                title: 'Excluir contas',
                message: `Deseja excluir ${ids.length} ${ids.length === 1 ? 'conta selecionada' : 'contas selecionadas'}?`,
                onConfirm: () => commitAccounts(accounts.filter((item) => !ids.includes(item.id))),
              })}
            />
          )}
        </main>
      </div>

      {launchOpen && (
        <LaunchModal
          accounts={accounts}
          incomeCategories={incomeCategories}
          expenseCategories={expenseCategories}
          initialType={launchType}
          onClose={() => setLaunchOpen(false)}
          onCreate={(items) => {
            commitTransactions([...transactions, ...items]);
            setLaunchOpen(false);
          }}
        />
      )}

      {importOpen && (
        <ImportTransactionsModal
          onClose={() => setImportOpen(false)}
          existingCategories={customCategories}
          onImport={(items, newCats) => {
            if (newCats.length) commitCustomCategories([...customCategories, ...newCats]);
            commitTransactions([...transactions, ...items]);
            const firstDate = items[0]?.dueDate;
            if (firstDate) setReferenceDate(new Date(Number(firstDate.slice(0, 4)), Number(firstDate.slice(5, 7)) - 1, 1));
            setImportOpen(false);
          }}
        />
      )}

      {(accountOpen || editingAccount) && (
        <AccountModal
          accounts={accounts}
          account={editingAccount}
          onClose={() => { setAccountOpen(false); setEditingAccount(null); }}
          onSave={(account) => {
            const exists = accounts.some((item) => item.id === account.id);
            commitAccounts(exists ? accounts.map((item) => (item.id === account.id ? account : item)) : [...accounts, account]);
            setAccountOpen(false);
            setEditingAccount(null);
          }}
        />
      )}

      {categoryOpen && (
        <CategoryModal
          items={categoryItems}
          category={editingCategory}
          onClose={() => { setCategoryOpen(false); setEditingCategory(null); }}
          onSave={(category) => {
            const exists = customCategories.some((item) => item.id === category.id);
            commitCustomCategories(exists ? customCategories.map((item) => item.id === category.id ? category : item) : [...customCategories, category]);
            setCategoryOpen(false);
            setEditingCategory(null);
          }}
        />
      )}

      {deletingCategory ? (
        <DeleteCategoryModal
          category={deletingCategory}
          onClose={() => setDeletingCategory(null)}
          onConfirm={() => {
            commitCustomCategories(customCategories.filter((item) => item.id !== deletingCategory.id));
            setDeletingCategory(null);
          }}
        />
      ) : null}

      {settleTarget && (
        <SettleModal
          item={settleTarget}
          accounts={accounts}
          onClose={() => setSettleTarget(null)}
          onSave={(accountId, settledAt, settledAmount) => {
            const account = accounts.find((candidate) => candidate.id === accountId);
            if (!account) return;
            commitTransactions(transactions.map((item) => (
              item.id === settleTarget.id ? { ...item, status: 'settled', accountId: account.id, account: account.name, settledAt, settledAmount } : item
            )));
            setSettleTarget(null);
          }}
        />
      )}

      {editingTransaction && (
        <EditTransactionModal
          item={editingTransaction}
          accounts={accounts}
          incomeCategories={incomeCategories}
          expenseCategories={expenseCategories}
          onClose={() => setEditingTransaction(null)}
          onSave={(updated) => {
            commitTransactions(transactions.map((item) => (item.id === updated.id ? updated : item)));
            setEditingTransaction(null);
          }}
        />
      )}

      {deletingTransaction && (
        <DeleteTransactionModal
          item={deletingTransaction}
          onClose={() => setDeletingTransaction(null)}
          relatedCount={transactions.filter((item) => item.groupId === deletingTransaction.groupId).length}
          onConfirmOne={() => {
            commitTransactions(transactions.filter((item) => item.id !== deletingTransaction.id));
            setDeletingTransaction(null);
          }}
          onConfirmAll={() => {
            commitTransactions(transactions.filter((item) => item.groupId !== deletingTransaction.groupId));
            setDeletingTransaction(null);
          }}
        />
      )}

      {longPressTransaction && (
        <LongPressOptionsModal
          item={longPressTransaction}
          onClose={() => setLongPressTransaction(null)}
          onEdit={() => setEditingTransaction(longPressTransaction)}
          onSettle={() => setSettleTarget(longPressTransaction)}
          onDelete={() => setDeletingTransaction(longPressTransaction)}
          onMarkPending={() => markTransactionAsPending(longPressTransaction.id)}
          onSelect={() => toggleSelectTx(longPressTransaction.id)}
        />
      )}

      {goalOpen && (
        <GoalModal
          goal={editingGoal}
          onClose={() => { setGoalOpen(false); setEditingGoal(null); }}
          onSave={(goal) => {
            const exists = goals.some((item) => item.id === goal.id);
            commitGoals(exists ? goals.map((item) => (item.id === goal.id ? goal : item)) : [...goals, goal]);
            setGoalOpen(false);
            setEditingGoal(null);
          }}
        />
      )}

      {goalMovementTarget && (
        <GoalMovementModal
          goal={goalMovementTarget.goal}
          movementType={goalMovementTarget.type}
          onClose={() => setGoalMovementTarget(null)}
          onSave={(movement) => {
            commitGoals(goals.map((item) => (
              item.id === goalMovementTarget.goal.id ? { ...item, movements: [...item.movements, movement] } : item
            )));
            setGoalMovementTarget(null);
          }}
        />
      )}

      {deletingGoal && (
        <DeleteGoalModal
          goal={deletingGoal}
          onClose={() => setDeletingGoal(null)}
          onConfirm={() => {
            commitGoals(goals.filter((item) => item.id !== deletingGoal.id));
            setDeletingGoal(null);
          }}
        />
      )}

      {confirmDialog && <ConfirmDialog config={confirmDialog} onClose={() => setConfirmDialog(null)} />}
    </div>
  );
}

type RowAction = { key: string; label: string; icon: ReactNode; onClick: () => void; danger?: boolean };
type ConfirmConfig = { title: string; message: ReactNode; confirmLabel?: string; onConfirm: () => void };

function ConfirmDialog({ config, onClose }: { config: ConfirmConfig; onClose: () => void }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card confirm-modal" role="dialog" aria-modal="true" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">{config.title}</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body confirmation-message">
          <span className="confirmation-icon"><Trash2 size={20} /></span>
          <p>{config.message}</p>
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="button" className="button-danger" onClick={() => { config.onConfirm(); onClose(); }}><Trash2 size={15} /> {config.confirmLabel ?? 'Excluir'}</button>
        </div>
      </div>
    </div>
  );
}

function RowActions({ actions }: { actions: RowAction[] }) {
  const [open, setOpen] = useState(false);
  const [pos, setPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });
  const wrapRef = useRef<HTMLDivElement>(null);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    function onDown(event: PointerEvent) {
      const target = event.target as Node;
      if (!wrapRef.current?.contains(target) && !menuRef.current?.contains(target)) setOpen(false);
    }
    function onEsc(event: KeyboardEvent) { if (event.key === 'Escape') setOpen(false); }
    function onScroll() { setOpen(false); }
    document.addEventListener('pointerdown', onDown);
    document.addEventListener('keydown', onEsc);
    window.addEventListener('scroll', onScroll, true);
    return () => {
      document.removeEventListener('pointerdown', onDown);
      document.removeEventListener('keydown', onEsc);
      window.removeEventListener('scroll', onScroll, true);
    };
  }, [open]);

  function toggle() {
    if (!open && buttonRef.current) {
      const rect = buttonRef.current.getBoundingClientRect();
      const menuWidth = 196;
      const menuHeight = actions.length * 40 + 12;
      const top = rect.bottom + 6 + menuHeight > window.innerHeight ? rect.top - 6 - menuHeight : rect.bottom + 6;
      const left = Math.min(window.innerWidth - menuWidth - 12, Math.max(12, rect.right - menuWidth));
      setPos({ top: Math.max(12, top), left });
    }
    setOpen((value) => !value);
  }

  return (
    <div className="row-actions" ref={wrapRef}>
      <button type="button" ref={buttonRef} className={`row-actions-trigger${open ? ' active' : ''}`} aria-label="Ações" aria-haspopup="menu" aria-expanded={open} onClick={toggle}>
        <MoreVertical size={16} />
      </button>
      {open
        ? createPortal(
          <div className="row-actions-menu" role="menu" ref={menuRef} style={{ top: pos.top, left: pos.left }}>
            {actions.map((action) => (
              <button key={action.key} type="button" role="menuitem" className={`row-actions-item${action.danger ? ' row-actions-item--danger' : ''}`} onClick={() => { setOpen(false); action.onClick(); }}>
                {action.icon}
                {action.label}
              </button>
            ))}
          </div>,
          document.body,
        )
        : null}
    </div>
  );
}

function LongPressOptionsModal({ item, onClose, onEdit, onSettle, onDelete, onMarkPending, onSelect }: { item: Transaction; onClose: () => void; onEdit: () => void; onSettle: () => void; onDelete: () => void; onMarkPending: () => void; onSelect: () => void }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card longpress-modal" role="dialog" aria-modal="true" onClick={(event) => event.stopPropagation()} style={{ width: 'min(100%, 360px)', padding: '20px' }}>
        <div className="longpress-modal-header">
          <h2 className="modal-title modal-title-only" style={{ fontSize: '18px' }}>Opções da transação</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="longpress-modal-title">
          {displayTransactionDescription(item.description)}
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
          <button type="button" className="button-secondary longpress-btn--blue" onClick={() => { onClose(); onSelect(); }}>
            <CheckCircle2 size={18} /> Selecionar
          </button>
          {item.status === 'open' ? (
            <button type="button" className="button-secondary longpress-btn--green" onClick={() => { onClose(); onSettle(); }}>
              <CheckCircle2 size={18} /> Efetivado
            </button>
          ) : (
            <button type="button" className="button-secondary longpress-btn--blue" onClick={() => { onClose(); onMarkPending(); }}>
              <RotateCcw size={18} /> Desefetivar
            </button>
          )}
          <button type="button" className="button-secondary longpress-btn--neutral" onClick={() => { onClose(); onEdit(); }}>
            <Pencil size={18} /> Editar
          </button>
          <button type="button" className="button-danger longpress-btn--neutral" onClick={() => { onClose(); onDelete(); }}>
            <Trash2 size={18} /> Excluir
          </button>
        </div>
      </div>
    </div>
  );
}

function SwipeableTransactionRow({
  item,
  isSelected,
  selectionMode,
  onToggleSelect,
  onEdit,
  onSettle,
  onDelete,
  onMarkPending,
  onLongPress,
  categoryMeta,
}: {
  item: Transaction;
  isSelected: boolean;
  selectionMode: boolean;
  onToggleSelect: () => void;
  onEdit: () => void;
  onSettle: () => void;
  onDelete: () => void;
  onMarkPending: () => void;
  onLongPress: () => void;
  categoryMeta?: { color: string; icon: CategoryIconKey };
}) {
  const [offsetX, setOffsetX] = useState(0);
  const [isSwiping, setIsSwiping] = useState(false);
  const touchStartRef = useRef<{ x: number; y: number } | null>(null);
  const longPressTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  function clearTimer() {
    if (longPressTimerRef.current) {
      clearTimeout(longPressTimerRef.current);
      longPressTimerRef.current = null;
    }
  }

  function handleTouchStart(e: React.TouchEvent | React.PointerEvent) {
    if (selectionMode) return;
    if (e.target instanceof HTMLElement && (e.target.closest('button') || e.target.closest('input') || e.target.closest('.row-actions'))) return;
    const clientX = 'touches' in e ? e.touches[0].clientX : e.clientX;
    const clientY = 'touches' in e ? e.touches[0].clientY : e.clientY;
    touchStartRef.current = { x: clientX, y: clientY };
    setIsSwiping(false);

    clearTimer();
    longPressTimerRef.current = setTimeout(() => {
      onLongPress();
      touchStartRef.current = null;
    }, 500);
  }

  function handleTouchMove(e: React.TouchEvent | React.PointerEvent) {
    if (!touchStartRef.current) return;
    const clientX = 'touches' in e ? e.touches[0].clientX : e.clientX;
    const clientY = 'touches' in e ? e.touches[0].clientY : e.clientY;
    const dx = clientX - touchStartRef.current.x;
    const dy = clientY - touchStartRef.current.y;

    if (Math.abs(dx) > 10 || Math.abs(dy) > 10) {
      clearTimer();
    }

    if (Math.abs(dx) > Math.abs(dy) && Math.abs(dx) > 15) {
      setIsSwiping(true);
      const clamped = Math.max(-110, Math.min(110, dx));
      setOffsetX(clamped);
    }
  }

  function handleTouchEnd() {
    clearTimer();
    if (!touchStartRef.current) return;
    touchStartRef.current = null;
    setIsSwiping(false);

    if (offsetX <= -70) {
      setOffsetX(0);
      onDelete();
    } else if (offsetX >= 70 && item.status === 'open') {
      setOffsetX(0);
      onSettle();
    } else {
      setOffsetX(0);
    }
  }

  return (
    <div className="swipeable-row-container" style={{ position: 'relative', overflow: 'hidden', borderRadius: '14px' }}>
      <div
        style={{
          position: 'absolute',
          inset: 0,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '0 20px',
          borderRadius: '14px',
          backgroundColor: offsetX < 0 ? '#ef4444' : offsetX > 0 ? '#10b981' : 'transparent',
          color: '#ffffff',
          fontWeight: 700,
          fontSize: '14px',
          pointerEvents: 'none',
          transition: 'background-color 0.2s ease',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', opacity: offsetX > 20 ? 1 : 0, transition: 'opacity 0.2s ease' }}>
          <CheckCircle2 size={20} />
          <span>Efetivado</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', opacity: offsetX < -20 ? 1 : 0, transition: 'opacity 0.2s ease' }}>
          <span>Excluir</span>
          <Trash2 size={20} />
        </div>
      </div>
      <article
        className={`documents-report-row launches-report-row${isSelected ? ' is-selected' : ''}`}
        style={{ transform: `translateX(${offsetX}px)`, transition: isSwiping ? 'none' : 'transform 0.25s ease', position: 'relative', zIndex: 1 }}
        onTouchStart={handleTouchStart}
        onTouchMove={handleTouchMove}
        onTouchEnd={handleTouchEnd}
        onPointerDown={handleTouchStart}
        onPointerMove={handleTouchMove}
        onPointerUp={handleTouchEnd}
        onClick={(event) => {
          if (!selectionMode) return;
          if (event.target instanceof HTMLElement && (event.target.closest('button') || event.target.closest('input') || event.target.closest('.row-actions'))) return;
          onToggleSelect();
        }}
      >
        <span className="row-check-col" onClick={(e) => e.stopPropagation()}>
          <input type="checkbox" className="row-check" checked={isSelected} onChange={onToggleSelect} aria-label={`Selecionar ${item.description}`} />
        </span>
        <div className="launch-main" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          {categoryMeta ? (
            <span className="launch-category-icon launch-leading-icon" style={{ backgroundColor: `${categoryMeta.color}18`, color: categoryMeta.color, display: 'inline-flex', padding: '6px', borderRadius: '50%', flexShrink: 0 }}>
              <CategoryIconGraphic icon={categoryMeta.icon} size={14} />
            </span>
          ) : null}
          <div style={{ display: 'flex', flexDirection: 'column' }}>
            <strong>{displayTransactionDescription(item.description)}</strong>
            {sharedTransactionLabel(item) ? <span className="shared-transaction-badge"><Share size={13} /> {sharedTransactionLabel(item)}</span> : null}
            {item.notes ? <span className="launch-notes">{item.notes}</span> : null}
          </div>
        </div>
        <span className="launch-category">
          {categoryMeta ? (
            <span className="launch-category-icon" style={{ backgroundColor: `${categoryMeta.color}18`, color: categoryMeta.color }}>
              <CategoryIconGraphic icon={categoryMeta.icon} size={14} />
            </span>
          ) : null}
          <span className="launch-category-name">{item.category}</span>
        </span>
        <span
          className={`status-icon transaction-kind transaction-kind--${item.type}`}
          style={categoryMeta ? { backgroundColor: `${categoryMeta.color}1C`, color: categoryMeta.color } : undefined}
        >
          {categoryMeta ? (
            <CategoryIconGraphic icon={categoryMeta.icon} size={20} />
          ) : (
            item.type === 'income' ? <ArrowUpRight size={16} /> : <ArrowDownLeft size={16} />
          )}
          <span className="status-icon-text">{item.type === 'income' ? 'Receita' : 'Despesa'}</span>
        </span>
        <span className="recurrence-pill launch-recurrence">{recurrenceLabel(item)}</span>
        <span className="launch-date">{formatDate(item.dueDate)}</span>
        <strong className={`launch-value launch-value--${item.type}`}>{item.type === 'income' ? '+ ' : '- '}{formatCurrency(item.amount)}</strong>
        <span className={`launch-real${item.status === 'settled' ? ` launch-real--${item.type}` : ' launch-real--pending'}`}>
          {item.status === 'settled' ? formatCurrency(item.settledAmount ?? item.amount) : '—'}
        </span>
        <span className="launch-muted">{item.status === 'settled' && item.account ? item.account : '—'}</span>
        <span className="launch-status-cell">
          <StatusIcon status={item.status} dueDate={item.dueDate} />
        </span>
        <span className="launch-actions-cell" onClick={(e) => e.stopPropagation()}>
          <RowActions actions={[
            { key: 'edit', label: 'Editar', icon: <Pencil size={15} />, onClick: onEdit },
            ...(item.status === 'open'
              ? [{ key: 'settle', label: 'Efetivado', icon: <CheckCircle2 size={15} />, onClick: onSettle }]
              : [{ key: 'pending', label: 'Desefetivar', icon: <RotateCcw size={15} />, onClick: onMarkPending }]),
            { key: 'delete', label: 'Excluir', icon: <Trash2 size={15} />, onClick: onDelete, danger: true },
          ]} />
        </span>
      </article>
    </div>
  );
}

function StatusIcon({ status, dueDate }: { status: TransactionStatus; dueDate?: string }) {
  const today = new Date().toISOString().slice(0, 10);
  const isOverdue = status === 'open' && dueDate !== undefined && dueDate < today;
  const effectiveStatus = isOverdue ? 'overdue' : status;
  const label = status === 'settled' ? 'Efetivado' : isOverdue ? 'Atrasado' : 'Pendente';
  return (
    <span className={`status-icon status-icon--${effectiveStatus}`} title={label}>
      {status === 'settled' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
      <span className="status-icon-text">{label}</span>
    </span>
  );
}

function MonthNavigator({ date, onChange }: { date: Date; onChange: (date: Date) => void }) {
  const [open, setOpen] = useState(false);
  const [pickYear, setPickYear] = useState(date.getFullYear());
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    function onDown(event: PointerEvent) { if (!ref.current?.contains(event.target as Node)) setOpen(false); }
    function onEsc(event: KeyboardEvent) { if (event.key === 'Escape') setOpen(false); }
    document.addEventListener('pointerdown', onDown);
    document.addEventListener('keydown', onEsc);
    return () => { document.removeEventListener('pointerdown', onDown); document.removeEventListener('keydown', onEsc); };
  }, [open]);

  const monthLabel = new Intl.DateTimeFormat('pt-BR', { month: 'long' }).format(date);
  const monthTitle = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);

  return (
    <div className="month-nav" ref={ref}>
      <button type="button" className="month-nav-arrow" aria-label="Mês anterior" onClick={() => onChange(new Date(date.getFullYear(), date.getMonth() - 1, 1))}><ChevronLeft size={18} /></button>
      <button type="button" className="month-nav-pill" onClick={() => { setPickYear(date.getFullYear()); setOpen((value) => !value); }} aria-haspopup="dialog" aria-expanded={open}>
        <strong>{monthTitle}</strong> {date.getFullYear()}
      </button>
      <button type="button" className="month-nav-arrow" aria-label="Próximo mês" onClick={() => onChange(new Date(date.getFullYear(), date.getMonth() + 1, 1))}><ChevronRight size={18} /></button>
      {open ? (
        <div className="month-nav-popover" role="dialog" aria-label="Selecionar mês e ano">
          <div className="month-nav-year">
            <button type="button" className="month-nav-arrow" aria-label="Ano anterior" onClick={() => setPickYear((year) => year - 1)}><ChevronLeft size={16} /></button>
            <strong>{pickYear}</strong>
            <button type="button" className="month-nav-arrow" aria-label="Próximo ano" onClick={() => setPickYear((year) => year + 1)}><ChevronRight size={16} /></button>
          </div>
          <div className="month-nav-grid">
            {MONTH_OPTIONS.map((option) => {
              const active = date.getFullYear() === pickYear && date.getMonth() === option.value;
              const short = new Intl.DateTimeFormat('pt-BR', { month: 'short' }).format(new Date(2026, option.value, 1)).replace('.', '');
              return (
                <button key={option.value} type="button" className={`month-nav-month${active ? ' active' : ''}`} onClick={() => { onChange(new Date(pickYear, option.value, 1)); setOpen(false); }}>{short}</button>
              );
            })}
          </div>
        </div>
      ) : null}
    </div>
  );
}

const DashboardPage = memo(function DashboardPage({ transactions, referenceDate, onChangeDate, onNavigate }: {
  transactions: Transaction[];
  referenceDate: Date;
  onChangeDate: (date: Date) => void;
  onNavigate: (page: AppPage) => void;
}) {
  const incomeExpenseChartRef = useRef<HTMLDivElement>(null);
  const balanceChartRef = useRef<HTMLDivElement>(null);
  const [incomeExpenseChartKey, setIncomeExpenseChartKey] = useState(0);
  const [balanceChartKey, setBalanceChartKey] = useState(0);

  useEffect(() => {
    function closeChartTooltips(event: PointerEvent) {
      const target = event.target as Node;
      if (!incomeExpenseChartRef.current?.contains(target)) setIncomeExpenseChartKey((key) => key + 1);
      if (!balanceChartRef.current?.contains(target)) setBalanceChartKey((key) => key + 1);
    }

    document.addEventListener('pointerdown', closeChartTooltips);
    return () => document.removeEventListener('pointerdown', closeChartTooltips);
  }, []);

  const monthKeyValue = monthKey(referenceDate);

  const { income, expense, pendingIncome, pendingExpense } = useMemo(() => {
    const monthTxs = transactions.filter((item) => item.dueDate.slice(0, 7) === monthKeyValue);
    return {
      income: monthTxs.filter((item) => item.type === 'income').reduce((sum, item) => sum + item.amount, 0),
      expense: monthTxs.filter((item) => item.type === 'expense').reduce((sum, item) => sum + item.amount, 0),
      pendingIncome: monthTxs.filter((item) => item.type === 'income' && item.status === 'open').reduce((sum, item) => sum + item.amount, 0),
      pendingExpense: monthTxs.filter((item) => item.type === 'expense' && item.status === 'open').reduce((sum, item) => sum + item.amount, 0),
    };
  }, [transactions, monthKeyValue]);

  const projectedBalance = income - expense;

  const monthlySeries = useMemo(() => {
    const year = referenceDate.getFullYear();
    return Array.from({ length: 12 }, (_, index) => {
      const date = new Date(year, index, 1);
      const key = monthKey(date);
      const monthIncome = transactions.filter((item) => item.type === 'income' && item.dueDate.slice(0, 7) === key).reduce((sum, item) => sum + item.amount, 0);
      const monthExpense = transactions.filter((item) => item.type === 'expense' && item.dueDate.slice(0, 7) === key).reduce((sum, item) => sum + item.amount, 0);
      return {
        month: new Intl.DateTimeFormat('pt-BR', { month: 'short' }).format(date).replace('.', ''),
        Receitas: monthIncome,
        Despesas: monthExpense,
        Saldo: monthIncome - monthExpense,
      };
    });
  }, [transactions, referenceDate]);

  const mobileChartEnd = referenceDate.getMonth() + 1;
  const mobileChartStart = Math.max(0, mobileChartEnd - 6);
  const mobileMonthlySeries = monthlySeries.slice(mobileChartStart, mobileChartEnd);

  const balanceColor = (value: number) => (value > 0 ? '#0284c7' : value < 0 ? '#dc2626' : '#475569');
  const renderBalanceDot = ({ cx, cy, payload }: BalanceDotProps) => {
    const dotX = Number(cx);
    const dotY = Number(cy);
    const balance = Number(payload?.Saldo ?? 0);

    if (Number.isNaN(dotX) || Number.isNaN(dotY)) return null;

    return <circle cx={dotX} cy={dotY} r={4} fill="#ffffff" stroke={balanceColor(balance)} strokeWidth={2} />;
  };

  const renderBalanceLabel = ({ x, y, value }: BalanceLabelProps) => {
    const labelX = Number(x);
    const labelY = Number(y);
    const balance = Number(value ?? 0);

    if (Number.isNaN(labelX) || Number.isNaN(labelY)) return <g />;

    return (
      <text x={labelX} y={labelY + (balance < 0 ? 18 : -10)} textAnchor="middle" fill={balanceColor(balance)} fontSize={11} fontWeight={800}>
        {formatCurrency(balance)}
      </text>
    );
  };

  const hasData = transactions.length > 0;
  const monthTransactions = useMemo(
    () => transactions.filter((item) => item.dueDate.slice(0, 7) === monthKeyValue),
    [transactions, monthKeyValue],
  );
  const settledTotal = useMemo(
    () => monthTransactions.filter((item) => item.status === 'settled').reduce((sum, item) => sum + item.amount, 0),
    [monthTransactions],
  );
  const openTotal = pendingIncome + pendingExpense;
  const monthTotal = income + expense;
  const settledPercent = monthTotal > 0 ? Math.round((settledTotal / monthTotal) * 100) : 0;

  // Relatório de transações do mês (listar + exportar)
  const reportItems = useMemo(() => [...monthTransactions].sort((a, b) => a.dueDate.localeCompare(b.dueDate)), [monthTransactions]);
  const exportMonthLabel = `${MONTH_OPTIONS[referenceDate.getMonth()]?.label ?? ''} de ${referenceDate.getFullYear()}`;
  const reportTitle = `Relatório de ${exportMonthLabel}`;
  const exportPdf = () => exportReportPdf(reportItems, reportTitle);
  const exportExcel = () => { void exportReportExcel(reportItems, reportTitle); };

  return (
    <>
      <section className="page-header page-header-split">
        <div className="page-header-left">
          <h1 className="page-title">Visão Geral</h1>
        </div>
        <div className="page-header-center">
          <MonthNavigator date={referenceDate} onChange={onChangeDate} />
        </div>
        <div className="page-header-actions">
          {hasData ? (
            <>
              <button type="button" className="page-secondary-action report-export-desktop" onClick={exportPdf}><FileText size={16} /> Baixar PDF</button>
              <button type="button" className="page-primary-action report-export-desktop" onClick={exportExcel}><FileSpreadsheet size={16} /> Baixar Excel</button>
              <ReportExportMenu onPdf={exportPdf} onExcel={exportExcel} />
            </>
          ) : null}
        </div>
      </section>

      <div className="dashboard-metric-grid">
        <MetricCard label="Receita prevista" value={formatCurrency(income)} tone="success" icon={<TrendingUp size={18} />} />
        <MetricCard label="Despesa prevista" value={formatCurrency(expense)} tone="danger" icon={<TrendingDown size={18} />} />
        <MetricCard label="Falta a receber" value={formatCurrency(pendingIncome)} tone="warning" icon={<Clock3 size={18} />} />
        <MetricCard label="Falta a pagar" value={formatCurrency(pendingExpense)} tone="alert" icon={<Clock3 size={18} />} />
        <article className={`dashboard-result-card dashboard-result-card--desktop dashboard-result-card--${projectedBalance >= 0 ? 'positive' : 'negative'}`}>
          <div className="dashboard-month-card-head"><span>Resultado do mês</span><strong>{projectedBalance >= 0 ? 'Positivo' : 'Negativo'}</strong></div>
          <div className="dashboard-month-balance">{formatCurrency(projectedBalance)}</div>
          <div className="dashboard-month-track" aria-label={`${settledPercent}% efetivado`}><span style={{ width: `${settledPercent}%` }} /></div>
          <div className="dashboard-month-meta"><span>{settledPercent}% efetivado</span><span>{formatCurrency(openTotal)} em aberto</span></div>
        </article>
      </div>

      <section className="dashboard-insights-grid dashboard-insights-grid--mobile">
        <article className={`dashboard-month-card dashboard-month-card--${projectedBalance >= 0 ? 'positive' : 'negative'}`}>
          <div className="dashboard-month-card-head">
            <span>Resultado do mês</span>
            <strong>{projectedBalance >= 0 ? 'Positivo' : 'Negativo'}</strong>
          </div>
          <div className="dashboard-month-balance">{formatCurrency(projectedBalance)}</div>
          <div className="dashboard-month-track" aria-label={`${settledPercent}% efetivado`}><span style={{ width: `${settledPercent}%` }} /></div>
          <div className="dashboard-month-meta"><span>{settledPercent}% efetivado</span><span>{formatCurrency(openTotal)} em aberto</span></div>
        </article>
      </section>

      {hasData ? (
        <>
          <div className="dashboard-charts dashboard-charts--single">
            <section className="resource-panel chart-panel chart-panel--wide">
              <div className="chart-head"><strong>Receitas x Despesas</strong><span className="chart-head-label--desktop">Janeiro a dezembro</span><span className="chart-head-label--mobile">Últimos 6 meses</span></div>
              <div ref={incomeExpenseChartRef} className="chart-box chart-box--wide dashboard-income-chart--full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart key={incomeExpenseChartKey} data={monthlySeries} margin={{ top: 12, right: 18, left: 8, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f6" />
                    <XAxis dataKey="month" tickLine={false} axisLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                    <RechartsTooltip formatter={(value) => formatCurrency(Number(value))} cursor={{ fill: 'rgba(148,163,184,0.08)' }} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Bar dataKey="Receitas" fill="#1B99D8" radius={[6, 6, 0, 0]} maxBarSize={30} />
                    <Bar dataKey="Despesas" fill="#dc2626" radius={[6, 6, 0, 0]} maxBarSize={30} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="chart-box chart-box--wide dashboard-income-chart--mobile">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart key={`mob-${incomeExpenseChartKey}`} data={mobileMonthlySeries} margin={{ top: 12, right: 18, left: 8, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f6" />
                    <XAxis dataKey="month" tickLine={false} axisLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                    <RechartsTooltip formatter={(value) => formatCurrency(Number(value))} cursor={{ fill: 'rgba(148,163,184,0.08)' }} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Bar dataKey="Receitas" fill="#1B99D8" radius={[6, 6, 0, 0]} maxBarSize={36} />
                    <Bar dataKey="Despesas" fill="#dc2626" radius={[6, 6, 0, 0]} maxBarSize={36} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </section>
          </div>

          <div className="dashboard-charts dashboard-charts--single">
            <section className="resource-panel chart-panel chart-panel--wide">
              <div className="chart-head"><strong>Saldo mensal</strong><span className="chart-head-label--desktop">Janeiro a dezembro</span><span className="chart-head-label--mobile">Últimos 6 meses</span></div>
              <div ref={balanceChartRef} className="chart-box chart-box--wide dashboard-balance-chart dashboard-balance-chart--line">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart key={balanceChartKey} data={monthlySeries} margin={{ top: 34, right: 30, left: 30, bottom: 8 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f6" />
                    <XAxis dataKey="month" tickLine={false} axisLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                    <ReferenceLine y={0} stroke="#cbd5e1" strokeDasharray="4 4" />
                    <RechartsTooltip formatter={(value) => formatCurrency(Number(value))} cursor={{ stroke: '#94a3b8', strokeWidth: 1 }} />
                    <Line type="monotone" dataKey="Saldo" stroke="#334155" strokeWidth={3} dot={renderBalanceDot} activeDot={{ r: 6 }}>
                      <LabelList dataKey="Saldo" content={renderBalanceLabel} />
                    </Line>
                  </LineChart>
                </ResponsiveContainer>
              </div>
              <div className="chart-box chart-box--wide dashboard-balance-chart dashboard-balance-chart--area">
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart key={`mob-area-${balanceChartKey}`} data={mobileMonthlySeries} margin={{ top: 12, right: 18, left: 8, bottom: 0 }}>
                    <defs>
                      <linearGradient id="balanceGradient" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor="#1B99D8" stopOpacity={0.25} />
                        <stop offset="95%" stopColor="#1B99D8" stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f6" />
                    <XAxis dataKey="month" tickLine={false} axisLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                    <ReferenceLine y={0} stroke="#cbd5e1" strokeDasharray="4 4" />
                    <RechartsTooltip formatter={(value) => formatCurrency(Number(value))} cursor={{ stroke: '#94a3b8', strokeWidth: 1 }} />
                    <Area type="monotone" dataKey="Saldo" stroke="#1B99D8" strokeWidth={2.5} fill="url(#balanceGradient)" dot={{ r: 4, fill: '#1B99D8', strokeWidth: 0 }} activeDot={{ r: 6 }} />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </section>
          </div>

        </>
      ) : (
        <section className="resource-panel">
          <div className="documents-report">
            <div className="resource-empty-state">
              <LayoutGrid size={30} />
              <strong>Seu painel está vazio</strong>
              <p>Registre transações para visualizar gráficos, saldos e o resumo financeiro do mês.</p>
              <button type="button" className="page-primary-action" onClick={() => onNavigate('transactions')}><Plus size={16} /> Registrar transação</button>
            </div>
          </div>
        </section>
      )}
      {hasData ? (
        <DashboardExportMenu onPdf={exportPdf} onExcel={exportExcel} />
      ) : null}
    </>
  );
});
function GoalsPage({ goals, filters, draftFilters, filterOpen, filterControlRef, onOpenFilters, onDraftFiltersChange, onApplyFilters, onNew, onEdit, onDelete, onDeposit, onWithdraw }: {
  goals: Goal[];
  filters: GoalFilters;
  draftFilters: GoalFilters;
  filterOpen: boolean;
  filterControlRef: RefObject<HTMLDivElement | null>;
  onOpenFilters: () => void;
  onDraftFiltersChange: (filters: GoalFilters) => void;
  onApplyFilters: () => void;
  onNew: () => void;
  onEdit: (goal: Goal) => void;
  onDelete: (goal: Goal) => void;
  onDeposit: (goal: Goal) => void;
  onWithdraw: (goal: Goal) => void;
}) {
  const [movementsGoal, setMovementsGoal] = useState<Goal | null>(null);
  const normalizedSearch = filters.search.trim().toLowerCase();
  const filteredGoals = goals.filter((goal) => {
    const saved = goalSaved(goal);
    const reached = goal.targetAmount > 0 && saved >= goal.targetAmount;
    const matchesStatus = filters.status === 'all'
      || (filters.status === 'reached' && reached)
      || (filters.status === 'active' && !reached);
    return matchesStatus && goal.name.toLowerCase().includes(normalizedSearch);
  });
  const activeFilterCount = Number(Boolean(filters.search.trim())) + Number(filters.status !== 'all');

  return (
    <>
      <section className="page-header page-header-split">
        <div className="page-header-left">
          <h1 className="page-title">Metas</h1>
        </div>
        <div className="page-header-center"></div>
        <div className="page-header-actions">
          <div className="filter-control goal-filter-control" ref={filterControlRef}>
            <button
              type="button"
              className={`filter-trigger btn-icon-only${filterOpen ? ' active' : ''}`}
              data-goal-filter-trigger
              onClick={onOpenFilters}
              aria-expanded={filterOpen}
              aria-haspopup="dialog"
              title="Filtros"
            >
              <ListFilter size={16} />
              {activeFilterCount > 0 ? <span className="filter-badge">{activeFilterCount}</span> : null}
            </button>
            {filterOpen ? (
              <div className="filter-popover goal-filter-popover" role="dialog" aria-label="Filtros das metas">
                <div className="filter-popover-header"><strong>Filtrar metas</strong><button type="button" className="filter-popover-close" onClick={onOpenFilters} aria-label="Fechar filtros"><X size={18} /></button></div>
                <div className="filter-grid filter-grid--stacked">
                  <label className="filter-field">
                    <span>Meta</span>
                    <input value={draftFilters.search} onChange={(event) => onDraftFiltersChange({ ...draftFilters, search: event.target.value })} placeholder="Buscar pelo nome" />
                  </label>
                  <label className="filter-field">
                    <span>Status</span>
                    <select value={draftFilters.status} onChange={(event) => onDraftFiltersChange({ ...draftFilters, status: event.target.value as GoalFilters['status'] })}>
                      <option value="all">Todas</option>
                      <option value="active">Em andamento</option>
                      <option value="reached">Concluídas</option>
                    </select>
                  </label>
                </div>
                <div className="filter-popover-actions">
                  <button type="button" className="filter-clear" onClick={() => onDraftFiltersChange({ search: '', status: 'all' })}><RotateCcw size={14} /> Limpar</button>
                  <button type="button" className="filter-apply" onClick={onApplyFilters}>Aplicar filtros</button>
                </div>
              </div>
            ) : null}
          </div>
          <button type="button" className="page-primary-action" onClick={onNew}><Plus size={16} /> Nova meta</button>
        </div>
      </section>

      {filteredGoals.length ? (
        <div className="goals-grid">
          {filteredGoals.map((goal) => {
            const saved = goalSaved(goal);
            const percent = goal.targetAmount > 0 ? Math.min(100, Math.round((saved / goal.targetAmount) * 100)) : 0;
            const reached = saved >= goal.targetAmount && goal.targetAmount > 0;
            return (
              <article key={goal.id} className="goal-card">
                <div className="goal-card-accent" style={{ background: goal.color }} />
                <div className="goal-card-body">
                  <div className="goal-card-head">
                    <span className="goal-icon" style={{ backgroundColor: `${goal.color}18`, color: goal.color }}><CategoryIconGraphic icon={goal.icon} size={20} /></span>
                    <div className="goal-title">
                      <strong>{goal.name}</strong>
                      {goal.deadline ? <span><CalendarDays size={11} /> Até {formatDate(goal.deadline)}</span> : <span>Sem prazo</span>}
                    </div>
                  </div>

                  <div className="goal-progress">
                    <div className="goal-progress-bar"><span style={{ width: `${percent}%`, background: `linear-gradient(90deg, ${goal.color}, ${goal.color}cc)` }} /></div>
                    <div className="goal-progress-meta">
                      <span className={reached ? 'goal-reached' : ''}>{reached ? '✓ Meta atingida' : `${percent}%`}</span>
                      <span>{formatCurrency(saved)} <small>de {formatCurrency(goal.targetAmount)}</small></span>
                    </div>
                  </div>

                  <div className="goal-card-buttons">
                    <button type="button" className="goal-button goal-button--deposit" onClick={() => onDeposit(goal)}><Plus size={15} /><span className="goal-btn-text">Depositar</span></button>
                    <button type="button" className="goal-button goal-button--withdraw" onClick={() => onWithdraw(goal)} disabled={saved <= 0}><Minus size={15} /><span className="goal-btn-text">Resgatar</span></button>
                    <button type="button" className="goal-button goal-history-btn" onClick={() => setMovementsGoal(goal)} title="Ver histórico">
                      <ReceiptText size={15} /><span className="goal-btn-text">Histórico{goal.movements.length > 0 ? ` (${goal.movements.length})` : ''}</span>
                    </button>
                  </div>

                  <div className="goal-card-actions">
                    <button type="button" className="category-action-button" title="Editar meta" aria-label={`Editar ${goal.name}`} onClick={() => onEdit(goal)}><Pencil size={15} /></button>
                    <button type="button" className="category-action-button category-action-button--danger" title="Excluir meta" aria-label={`Excluir ${goal.name}`} onClick={() => onDelete(goal)}><Trash2 size={15} /></button>
                  </div>
                </div>
              </article>
            );
          })}
        </div>
      ) : (
        <section className="resource-panel">
          <div className="documents-report">
            <div className="resource-empty-state">
              <Target size={30} />
              <strong>Nenhuma meta criada</strong>
              <p>Crie metas como reserva de emergência, viagem ou compra e registre depósitos e resgates ao longo do tempo.</p>
              <button type="button" className="page-primary-action" onClick={onNew}><Plus size={16} /> Criar primeira meta</button>
            </div>
          </div>
        </section>
      )}

      {movementsGoal ? (
        <div className="modal-overlay" onClick={(e) => { if (e.target === e.currentTarget) setMovementsGoal(null); }}>
          <div className="modal-card goal-movements-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                <span className="goal-icon" style={{ backgroundColor: `${movementsGoal.color}18`, color: movementsGoal.color }}><CategoryIconGraphic icon={movementsGoal.icon} size={18} /></span>
                <div>
                  <h2 className="modal-title" style={{ marginBottom: '2px' }}>Lançamentos — {movementsGoal.name}</h2>
                  <p className="goal-modal-subtitle">{movementsGoal.movements.length} {movementsGoal.movements.length === 1 ? 'lan?amento' : 'lan?amentos'} ? Guardado: {formatCurrency(goalSaved(movementsGoal))}</p>
                </div>
              </div>
              <button type="button" className="modal-close-button" onClick={() => setMovementsGoal(null)} aria-label="Fechar"><X size={18} /></button>
            </div>
            <div className="modal-body">
              <div className="goal-movements-list">
                {movementsGoal.movements.length ? (
                  [...movementsGoal.movements].reverse().map((movement) => (
                    <div key={movement.id} className="goal-movement-item">
                      <div className="goal-movement-left">
                        <span className={`goal-movement-badge goal-movement-badge--${movement.type}`}>
                          {movement.type === 'deposit' ? <ArrowDownLeft size={16} /> : <ArrowUpRight size={16} />}
                        </span>
                        <div className="goal-movement-info">
                          <strong>{movement.type === 'deposit' ? 'Depósito' : 'Resgate'}</strong>
                          <span className="goal-movement-date"><CalendarDays size={12} /> {formatDate(movement.date)}</span>
                          {movement.note ? <span className="goal-movement-note">{movement.note}</span> : null}
                        </div>
                      </div>
                      <strong className={`goal-movement-amount goal-movement-amount--${movement.type}`}>
                        {movement.type === 'deposit' ? '+' : '-'}{formatCurrency(movement.amount)}
                      </strong>
                    </div>
                  ))
                ) : (
                  <div style={{ textAlign: 'center', padding: '24px 0', color: '#94a3b8', fontSize: '14px', fontWeight: 600 }}>
                    Nenhum lançamento registrado nesta meta ainda.
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </>
  );
}

function DashboardExportMenu({ onPdf, onExcel }: { onPdf: () => void; onExcel: () => void }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    function onDown(event: PointerEvent) {
      if (!ref.current?.contains(event.target as Node)) setOpen(false);
    }
    function onEsc(event: KeyboardEvent) {
      if (event.key === 'Escape') setOpen(false);
    }
    document.addEventListener('pointerdown', onDown);
    document.addEventListener('keydown', onEsc);
    return () => {
      document.removeEventListener('pointerdown', onDown);
      document.removeEventListener('keydown', onEsc);
    };
  }, [open]);

  return (
    <div className={`dashboard-export-menu${open ? ' open' : ''}`} ref={ref}>
      <button type="button" className="dashboard-export-fab" onClick={() => setOpen((value) => !value)} aria-label="Exportar relatório" aria-haspopup="menu" aria-expanded={open} title="Exportar relatório">
        <Share size={22} />
      </button>
      {open ? (
        <div className="dashboard-export-popover" role="menu">
          <button type="button" role="menuitem" onClick={() => { setOpen(false); onPdf(); }}><FileText size={16} /> Exportar em PDF</button>
          <button type="button" role="menuitem" onClick={() => { setOpen(false); onExcel(); }}><FileSpreadsheet size={16} /> Exportar em Excel</button>
        </div>
      ) : null}
    </div>
  );
}

function ReportExportMenu({ onPdf, onExcel }: { onPdf: () => void; onExcel: () => void }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    function onDown(event: PointerEvent) { if (!ref.current?.contains(event.target as Node)) setOpen(false); }
    function onEsc(event: KeyboardEvent) { if (event.key === 'Escape') setOpen(false); }
    document.addEventListener('pointerdown', onDown);
    document.addEventListener('keydown', onEsc);
    return () => { document.removeEventListener('pointerdown', onDown); document.removeEventListener('keydown', onEsc); };
  }, [open]);

  return (
    <div className="report-export-menu" ref={ref}>
      <button type="button" className={`report-export-trigger${open ? ' active' : ''}`} aria-haspopup="menu" aria-expanded={open} onClick={() => setOpen((value) => !value)} title="Exportar relatório">
        <Share size={18} />
      </button>
      {open ? (
        <div className="report-export-popover" role="menu">
          <button type="button" role="menuitem" onClick={() => { setOpen(false); onPdf(); }}><FileText size={16} /> Exportar PDF</button>
          <button type="button" role="menuitem" onClick={() => { setOpen(false); onExcel(); }}><FileSpreadsheet size={16} /> Exportar Excel</button>
        </div>
      ) : null}
    </div>
  );
}

function ReportsPage({ transactions, categoryLookup, referenceDate, onChangeDate }: { transactions: Transaction[]; categoryLookup: Map<string, CategoryItem>; referenceDate: Date; onChangeDate: (date: Date) => void }) {
  const monthKeyValue = monthKey(referenceDate);

  const filtered = useMemo(() => {
    return transactions
      .filter((item) => item.dueDate.slice(0, 7) === monthKeyValue)
      .sort((a, b) => a.dueDate.localeCompare(b.dueDate));
  }, [transactions, monthKeyValue]);

  const title = `Relatório de ${MONTH_OPTIONS[referenceDate.getMonth()]?.label ?? ''} de ${referenceDate.getFullYear()}`;
  const totals = reportTotals(filtered);
  const groupedTransactions = useMemo(() => {
    return filtered.reduce<Array<{ date: string; items: Transaction[] }>>((groups, item) => {
      const currentGroup = groups[groups.length - 1];
      if (currentGroup?.date === item.dueDate) {
        currentGroup.items.push(item);
      } else {
        groups.push({ date: item.dueDate, items: [item] });
      }
      return groups;
    }, []);
  }, [filtered]);
  return (
    <>
      <section className="page-header page-header-split">
        <div className="page-header-left">
          <h1 className="page-title">Relatórios</h1>
        </div>
        <div className="page-header-center">
          <MonthNavigator date={referenceDate} onChange={onChangeDate} />
        </div>
        <div className="page-header-actions">
          <button type="button" className="page-secondary-action report-export-desktop" onClick={() => exportReportPdf(filtered, title)}><FileText size={16} /> Baixar PDF</button>
          <button type="button" className="page-primary-action report-export-desktop" onClick={() => exportReportExcel(filtered, title)}><FileSpreadsheet size={16} /> Baixar Excel</button>
          <ReportExportMenu onPdf={() => exportReportPdf(filtered, title)} onExcel={() => exportReportExcel(filtered, title)} />
        </div>
      </section>

      <div className="reports-summary-desktop">
        <TransactionSummaryStats summary={totals} />
      </div>

      <div className="report-summary report-summary--mobile">
        <div className="report-summary-card">
          <div className="report-summary-header">
            <span className="report-summary-icon report-summary-icon--income"><TrendingUp size={20} /></span>
            <span className="report-summary-label">Receitas</span>
          </div>
          <strong className="report-summary-value report-positive">{formatCurrency(totals.income)}</strong>
        </div>
        <div className="report-summary-card">
          <div className="report-summary-header">
            <span className="report-summary-icon report-summary-icon--expense"><TrendingDown size={20} /></span>
            <span className="report-summary-label">Despesas</span>
          </div>
          <strong className="report-summary-value report-negative">{formatCurrency(totals.expense)}</strong>
        </div>
        <div className={`report-summary-card report-summary-card--balance ${totals.balance >= 0 ? 'report-summary-card--positive' : 'report-summary-card--negative'}`}>
          <div className="report-summary-header">
            <span className="report-summary-icon report-summary-icon--balance"><CircleDollarSign size={20} /></span>
            <span className="report-summary-label">Saldo Final</span>
          </div>
          <strong className={`report-summary-value ${totals.balance >= 0 ? 'report-positive' : 'report-negative'}`}>{formatCurrency(totals.balance)}</strong>
        </div>
      </div>

      <section className="resource-panel reports-desktop-table">
        <div className="documents-report launches-report">
          <div className="documents-report-head report-launches-head">
            <span>Descrição</span>
            <span>Categoria</span>
            <span>Tipo</span>
            <span>Recorrência</span>
            <span>Vencimento</span>
            <span>Valor previsto</span>
            <span>Valor real</span>
            <span>Status</span>
          </div>
          <div className="documents-report-body">
            {filtered.length ? filtered.map((item) => {
              const meta = categoryLookup.get(`${item.type}:${item.category}`);
              return (
                <article className="documents-report-row report-launches-row" key={item.id}>
                  <div className="launch-main">
                    <strong>{displayTransactionDescription(item.description)}</strong>
                    {sharedTransactionLabel(item) ? <span className="shared-transaction-badge"><Share size={13} /> {sharedTransactionLabel(item)}</span> : null}
                    {item.notes ? <span className="launch-notes">{item.notes}</span> : null}
                  </div>
                  <span className="launch-category">
                    {meta ? (
                      <span className="launch-category-icon" style={{ backgroundColor: `${meta.color}18`, color: meta.color }}>
                        <CategoryIconGraphic icon={meta.icon} size={14} />
                      </span>
                    ) : null}
                    <span className="launch-category-name">{item.category}</span>
                  </span>
                  <span className={`status-icon launch-type launch-type--${item.type}`}>
                    {item.type === 'income' ? <ArrowUpRight size={16} /> : <ArrowDownLeft size={16} />}
                    <span className="status-icon-text">{item.type === 'income' ? 'Receita' : 'Despesa'}</span>
                  </span>
                  <span className="recurrence-pill launch-recurrence">{recurrenceLabel(item)}</span>
                  <span className="launch-date">{formatDate(item.dueDate)}</span>
                  <strong className={`launch-value launch-value--${item.type}`}>{item.type === 'income' ? '+ ' : '- '}{formatCurrency(item.amount)}</strong>
                  <span className={`launch-real${item.status === 'settled' ? ` launch-real--${item.type}` : ' launch-real--pending'}`}>
                    {item.status === 'settled' ? formatCurrency(item.settledAmount ?? item.amount) : '—'}
                  </span>
                  <span className="launch-status-cell"><StatusIcon status={item.status} /></span>
                </article>
              );
            }) : (
              <div className="resource-empty-state">
                <ReceiptText size={30} />
                <strong>Nenhuma transação encontrada</strong>
                <p>Não há lançamentos no período selecionado.</p>
              </div>
            )}
          </div>
        </div>
      </section>
      <section className="resource-panel report-cards-section">
        <div className="report-cards-list">
          {groupedTransactions.length ? groupedTransactions.map((group) => (
            <section className="report-date-group" key={group.date}>
              <div className="report-date-separator">{formatDateGroup(group.date)}</div>
              <div className="report-date-cards">
                {group.items.map((item) => {
                  const meta = categoryLookup.get(`${item.type}:${item.category}`);

                  return (
                    <article className="report-transaction-card" key={item.id}>
                      <div className={`report-card-icon report-card-icon--${item.type}`}>
                        {item.type === 'income' ? <ArrowUpRight size={20} /> : <ArrowDownLeft size={20} />}
                      </div>

                      <div className="report-card-main">
                        <strong className="report-card-title">{displayTransactionDescription(item.description)}</strong>
                        <div className="report-card-subtitle">
                          {meta ? (
                            <span className="report-card-cat-icon" style={{ backgroundColor: `${meta.color}18`, color: meta.color }}>
                              <CategoryIconGraphic icon={meta.icon} size={13} />
                            </span>
                          ) : <WalletCards size={15} className="report-card-default-icon" />}
                          <span>{item.category}</span>
                        </div>
                        <span className="report-card-type">{recurrenceLabel(item)}</span>
                      </div>

                      <div className="report-card-right">
                        <strong className={`report-card-amount report-card-amount--${item.type}`}>
                          {item.type === 'income' ? '+ ' : '- '}{formatCurrency(item.amount)}
                        </strong>
                        <span className="report-card-date">{formatDate(item.dueDate)}</span>
                      </div>
                    </article>
                  );
                })}
              </div>
            </section>
          )) : (
            <div className="resource-empty-state">
              <ReceiptText size={30} />
              <strong>Nenhuma transação encontrada</strong>
              <p>Não há transações registradas para este período.</p>
            </div>
          )}
        </div>
      </section>
    </>
  );
}

function GoalModal({ goal, onClose, onSave }: { goal: Goal | null; onClose: () => void; onSave: (goal: Goal) => void }) {
  const [name, setName] = useState(goal?.name ?? '');
  const [targetAmount, setTargetAmount] = useState(() => goal ? formatCurrencyInput(goal.targetAmount) : '');
  const [deadline, setDeadline] = useState(goal?.deadline ?? '');
  const [color, setColor] = useState(goal?.color ?? ACCOUNT_COLORS[0] ?? '#1B99D8');
  const [icon, setIcon] = useState<CategoryIconKey>(goal?.icon ?? 'money');
  const [goalIconPickerOpen, setGoalIconPickerOpen] = useState(false);
  const [goalColorPickerOpen, setGoalColorPickerOpen] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    if (!goalIconPickerOpen) return;
    function handleEscape(event: KeyboardEvent) {
      if (event.key === 'Escape') setGoalIconPickerOpen(false);
    }
    document.addEventListener('keydown', handleEscape);
    return () => document.removeEventListener('keydown', handleEscape);
  }, [goalIconPickerOpen]);

  function submit(event: React.FormEvent) {
    event.preventDefault();
    const cleanName = name.trim();
    const amount = parseAmount(targetAmount);
    if (!cleanName) { setError('Informe um nome para a meta.'); return; }
    if (amount <= 0) { setError('Informe o valor do objetivo.'); return; }
    onSave({
      id: goal?.id ?? crypto.randomUUID(),
      name: cleanName,
      targetAmount: amount,
      deadline: deadline || undefined,
      color,
      icon,
      createdAt: goal?.createdAt ?? new Date().toISOString().slice(0, 10),
      movements: goal?.movements ?? [],
    });
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">{goal ? 'Editar meta' : 'Nova meta'}</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body">
          <div className="modal-form-grid">
            <label className="form-field form-field-full">
              <span className="form-label">Nome da meta</span>
              <input className="form-input" value={name} onChange={(event) => { setName(event.target.value); setError(''); }} placeholder="Ex.: Reserva de emergência, Viagem" />
            </label>
            <label className="form-field">
              <span className="form-label">Valor objetivo</span>
              <div className="form-input-wrap"><span className="form-input-prefix">R$</span><input className="form-input" value={targetAmount} onChange={(event) => { setTargetAmount(formatCurrencyInput(event.target.value)); setError(''); }} placeholder="0,00" inputMode="decimal" /></div>
            </label>
            <label className="form-field">
              <span className="form-label form-label-optional">Prazo</span>
              <div className="form-input-wrap"><CalendarDays size={16} /><input className="form-input" type="date" value={deadline} onChange={(event) => setDeadline(event.target.value)} /></div>
            </label>
            <div className="category-selector-grid">
              <div className="form-field">
                <span className="form-label form-label-optional">Ícone</span>
                <button type="button" className="form-input category-selector-btn" aria-haspopup="dialog" aria-expanded={goalIconPickerOpen} onClick={() => setGoalIconPickerOpen(true)}>
                  <span className="category-selector-icon" style={{ backgroundColor: `${color}18`, color }}><CategoryIconGraphic icon={icon} size={16} /></span>
                  <span className="category-selector-label">Escolher ícone</span>
                </button>
              </div>
              <div className="form-field">
                <span className="form-label form-label-optional">Cor</span>
                <button type="button" className="form-input category-selector-btn" aria-haspopup="dialog" aria-expanded={goalColorPickerOpen} onClick={() => setGoalColorPickerOpen(true)}>
                  <span className="category-selector-color" style={{ backgroundColor: color }} />
                  <span className="category-selector-label">Escolher cor</span>
                </button>
              </div>
            </div>
            {error ? <div className="form-error form-field-full"><AlertCircle size={15} /> {error}</div> : null}
          </div>
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary">{goal ? 'Salvar meta' : 'Criar meta'}</button>
        </div>
      </form>
      {goalColorPickerOpen ? <ColorSpectrumSheet value={color} onChange={setColor} onClose={() => setGoalColorPickerOpen(false)} /> : null}
      {goalIconPickerOpen
        ? createPortal(
          <div className="category-icon-picker-layer" onClick={() => setGoalIconPickerOpen(false)}>
            <div className="category-icon-picker-dialog" role="dialog" aria-modal="true" aria-label="Escolher ícone" onClick={(event) => event.stopPropagation()}>
              <div className="category-icon-picker-head">
                <strong>Escolher ícone</strong>
                <button type="button" onClick={() => setGoalIconPickerOpen(false)} aria-label="Fechar seletor"><X size={18} /></button>
              </div>
              <div className="category-icon-picker-grid">
                {ALL_CATEGORY_ICONS.map((key) => (
                  <button key={key} type="button" className={icon === key ? 'active' : ''} style={{ color }} aria-label={`Selecionar ícone ${key}`} aria-pressed={icon === key} onClick={() => { setIcon(key); setGoalIconPickerOpen(false); }}>
                    <CategoryIconGraphic icon={key} size={20} />
                  </button>
                ))}
              </div>
            </div>
          </div>,
          document.body,
        )
        : null}
    </div>
  );
}

function GoalMovementModal({ goal, movementType, onClose, onSave }: { goal: Goal; movementType: GoalMovementType; onClose: () => void; onSave: (movement: GoalMovement) => void }) {
  const saved = goalSaved(goal);
  const [amount, setAmount] = useState('');
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [note, setNote] = useState('');
  const [error, setError] = useState('');
  const isDeposit = movementType === 'deposit';

  function submit(event: React.FormEvent) {
    event.preventDefault();
    const parsedAmount = parseAmount(amount);
    if (parsedAmount <= 0) { setError('Informe um valor válido.'); return; }
    if (!isDeposit && parsedAmount > saved) { setError(`Você só pode resgatar até ${formatCurrency(saved)}.`); return; }
    onSave({ id: crypto.randomUUID(), type: movementType, amount: parsedAmount, date, note: note.trim() || undefined });
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card settle-modal" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">{isDeposit ? 'Depositar na meta' : 'Resgatar da meta'}</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body">
          <div className="goal-movement-summary">
            <span className="goal-icon" style={{ backgroundColor: `${goal.color}18`, color: goal.color }}><CategoryIconGraphic icon={goal.icon} size={16} /></span>
            <div><strong>{goal.name}</strong><span>Guardado: {formatCurrency(saved)} de {formatCurrency(goal.targetAmount)}</span></div>
          </div>
          <div className="modal-form-grid">
            <label className="form-field form-field-full"><span className="form-label">Valor {isDeposit ? 'do depósito' : 'do resgate'}</span><div className="form-input-wrap"><span className="form-input-prefix">R$</span><input className="form-input" value={amount} onChange={(event) => { setAmount(formatCurrencyInput(event.target.value)); setError(''); }} placeholder="0,00" inputMode="decimal" /></div></label>
            <label className="form-field form-field-full"><span className="form-label">Data</span><div className="form-input-wrap"><CalendarDays size={16} /><input className="form-input" type="date" value={date} onChange={(event) => setDate(event.target.value)} /></div></label>
            <label className="form-field form-field-full"><span className="form-label form-label-optional">Observação</span><input className="form-input" value={note} onChange={(event) => setNote(event.target.value)} placeholder="Opcional" /></label>
            {error ? <div className="form-error form-field-full"><AlertCircle size={15} /> {error}</div> : null}
          </div>
        </div>
        <div className="modal-actions"><button type="button" className="button-secondary" onClick={onClose}>Cancelar</button><button type="submit" className="button-primary">{isDeposit ? 'Depositar' : 'Resgatar'}</button></div>
      </form>
    </div>
  );
}

function DeleteGoalModal({ goal, onClose, onConfirm }: { goal: Goal; onClose: () => void; onConfirm: () => void }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card confirm-modal" role="dialog" aria-modal="true" aria-labelledby="delete-goal-title" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header"><h2 className="modal-title modal-title-only" id="delete-goal-title">Excluir meta</h2><button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button></div>
        <div className="modal-body confirmation-message"><span className="confirmation-icon"><Trash2 size={20} /></span><p>Deseja excluir a meta <strong>{goal.name}</strong> e todo o seu histórico?</p></div>
        <div className="modal-actions"><button type="button" className="button-secondary" onClick={onClose}>Cancelar</button><button type="button" className="button-danger" onClick={onConfirm}><Trash2 size={15} /> Excluir</button></div>
      </div>
    </div>
  );
}

function NotificationsPage({
  notifications,
  onClearNotifications,
}: {
  notifications: AppNotification[];
  onClearNotifications: () => void;
}) {
  const unread = notifications.filter((n) => !n.read).length;
  const sorted = useMemo(() => [...notifications].sort((a, b) => b.createdAt.localeCompare(a.createdAt)), [notifications]);

  function relativeTime(iso: string) {
    const diff = Date.now() - new Date(iso).getTime();
    const m = Math.floor(diff / 60000);
    if (m < 1) return 'Agora mesmo';
    if (m < 60) return `${m} min atrás`;
    const h = Math.floor(m / 60);
    if (h < 24) return `${h}h atrás`;
    const d = Math.floor(h / 24);
    return `${d} ${d === 1 ? 'dia' : 'dias'} atrás`;
  }

  const notifMeta: Record<AppNotification['type'], { icon: ReactNode; color: string }> = {
    friend_invite: { icon: <UserPlus size={18} />, color: 'blue' },
    friend_accepted: { icon: <Users size={18} />, color: 'green' },
    shared_created: { icon: <ArrowUpRight size={18} />, color: 'orange' },
    shared_approved: { icon: <CheckCircle2 size={18} />, color: 'green' },
    shared_declined: { icon: <X size={18} />, color: 'red' },
    shared_cancelled: { icon: <X size={18} />, color: 'gray' },
  };

  return (
    <section className="notifications-page">
      <section className="page-header page-header-split">
        <div className="page-header-left">
          <h1 className="page-title">Notificações</h1>
          {unread > 0 ? <span className="page-subtitle">{unread} não {unread === 1 ? 'lida' : 'lidas'}</span> : null}
        </div>
        <div className="page-header-actions">
          {notifications.length > 0 ? (
            <button type="button" className="page-secondary-action" onClick={onClearNotifications}>
              Limpar histórico
            </button>
          ) : null}
        </div>
      </section>

      <DeviceNotificationsCard />

      {sorted.length > 0 ? (
        <section className="notif-feed-section">
          <h2 className="notif-feed-title">Histórico de avisos</h2>
          <div className="notif-feed">
            {sorted.map((notification) => {
              const { icon, color } = notifMeta[notification.type];
              return (
                <div key={notification.id} className={`notif-item${notification.read ? '' : ' notif-item--unread'}`}>
                  <span className={`notif-icon notif-icon--${color}`}>{icon}</span>
                  <div className="notif-body">
                    <span className="notif-message">{notification.message}</span>
                    <span className="notif-time">{relativeTime(notification.createdAt)}</span>
                  </div>
                  {!notification.read ? <span className="notif-dot" aria-label="Não lida" /> : null}
                </div>
              );
            })}
          </div>
        </section>
      ) : (
        <div className="notif-empty">
          <span className="notif-empty-icon"><Bell size={28} /></span>
          <strong>Sem notificações</strong>
          <span>Convites de amizade e atualizações de compartilhamentos aparecerão aqui.</span>
        </div>
      )}
    </section>
  );
}

function HelpPage() {
  const sections = [
    {
      title: 'Primeiros passos',
      intro: 'O RubyLife começa vazio para que você monte o controle conforme a sua realidade. Cadastre a estrutura básica e depois registre as movimentações.',
      steps: [
        'No computador, use o menu lateral. No celular, use a barra inferior e o botão central de adicionar para acessar as ações mais frequentes.',
        'Comece em Contas: cadastre onde seu dinheiro fica, informe o tipo e registre o saldo inicial correto.',
        'Depois, revise Categorias e crie as classificações de receita, despesa ou transferência que você realmente utiliza.',
        'Registre receitas e despesas em Transações. Os lançamentos ficam em aberto até o pagamento ou recebimento ser efetivado.',
        'Use a Visão Geral e os Relatórios para acompanhar o mês. Metas, listas de compras e recursos com amigos funcionam de forma independente.',
      ],
    },
    {
      title: 'Visão Geral',
      intro: 'A tela inicial reúne os números do mês selecionado e oferece atalhos para investigar cada resultado.',
      steps: [
        'Use as setas do período para avançar ou voltar um mês. Todos os indicadores da página acompanham o período escolhido.',
        'Confira receitas, despesas, resultado mensal, valores a receber e valores a pagar para entender a situação do mês.',
        'Observe os gráficos de evolução e distribuição por categoria para localizar os maiores impactos no orçamento.',
        'Toque nos cards e atalhos disponíveis para abrir a área relacionada e analisar os lançamentos que formam aquele total.',
        'Use o menu de exportação para gerar PDF ou Excel quando precisar guardar ou compartilhar o resumo.',
      ],
    },
    {
      title: 'Transações',
      intro: 'Transações registram toda entrada ou saída, desde lançamentos únicos até contas fixas e compras parceladas.',
      steps: [
        'Toque em Nova transação e escolha Receita para entradas ou Despesa para saídas.',
        'Informe descrição, categoria, valor previsto e vencimento. Se já souber onde ocorrerá a movimentação, selecione também a conta.',
        'Escolha a recorrência: única para um evento, fixa para repetição mensal ou parcelada para uma quantidade definida de parcelas.',
        'Enquanto não houver pagamento ou recebimento, a transação permanece em aberto. Use Efetivado para confirmar data, conta e valor real.',
        'Ao efetivar, o saldo atual da conta é atualizado. Use Desefetivar para desfazer a confirmação e devolver a transação ao estado em aberto.',
        'Use Editar para corrigir dados e Excluir para remover apenas uma ocorrência ou, quando oferecido, todo o grupo recorrente.',
        'Use filtros de data, categoria e tipo, ou selecione várias linhas para efetivar, desefetivar e excluir em lote.',
        'Para trazer dados de uma planilha, use Importar, baixe o modelo, confira a prévia e confirme somente as linhas válidas.',
      ],
    },
    {
      title: 'Contas',
      intro: 'Contas representam carteira, banco, poupança ou investimento e mostram a diferença entre o saldo inicial e o saldo movimentado.',
      steps: [
        'Toque em Nova conta e informe um nome que identifique claramente onde o dinheiro está.',
        'Escolha o tipo: carteira, conta corrente, poupança ou investimento.',
        'Informe o saldo inicial existente antes das movimentações controladas pelo RubyLife.',
        'Escolha cor e ícone para reconhecer a conta com rapidez e salve o cadastro.',
        'O saldo atual considera o saldo inicial mais as receitas efetivadas e menos as despesas efetivadas vinculadas à conta.',
        'Use os filtros para buscar por nome ou tipo. As ações permitem editar, excluir ou selecionar várias contas.',
      ],
    },
    {
      title: 'Categorias',
      intro: 'Categorias classificam as movimentações e alimentam filtros, totais e gráficos do sistema.',
      steps: [
        'Toque em Nova categoria e dê um nome direto, como Salário, Moradia, Mercado ou Transporte.',
        'Defina o tipo correto: Receita, Despesa ou Transferência. O tipo determina onde a categoria poderá ser usada.',
        'Escolha um ícone na folha inferior e uma cor no seletor de espectro, arrastando até encontrar a tonalidade desejada.',
        'Evite categorias duplicadas ou com nomes muito parecidos para manter os relatórios claros.',
        'Use os filtros para localizar categorias por nome e tipo. Você também pode editar ou excluir cadastros personalizados.',
      ],
    },
    {
      title: 'Metas',
      intro: 'Metas acompanham objetivos financeiros separados do fluxo comum de receitas e despesas.',
      steps: [
        'Toque no botão de adicionar da barra superior ou em Nova meta.',
        'Informe o nome, valor objetivo e, se desejar, um prazo para conclusão.',
        'Escolha um ícone e abra o seletor de cor para arrastar pelo espectro até a tonalidade desejada.',
        'Use Depositar para registrar um valor guardado. Cada depósito aumenta o progresso apresentado no card.',
        'Use Resgatar para retirar parte do valor. O sistema não permite resgatar mais do que já foi guardado.',
        'Abra o histórico da meta para consultar depósitos e resgates. Use Editar para alterar seus dados ou Excluir para remover a meta e o histórico.',
        'Use os filtros para pesquisar pelo nome e separar metas ativas das já alcançadas.',
      ],
    },
    {
      title: 'Relatórios',
      intro: 'Relatórios consolidam o mês em totais e em uma relação cronológica das movimentações.',
      steps: [
        'Escolha o mês pelas setas do período para atualizar todo o relatório.',
        'Confira os totais de receitas, despesas e saldo final antes de analisar os lançamentos detalhados.',
        'Percorra os grupos por data para conferir descrição, categoria, recorrência, vencimento e valores.',
        'Se algum total estiver incorreto, volte a Transações e revise categoria, valor, vencimento e status dos registros.',
        'Use Baixar PDF para leitura e compartilhamento ou Baixar Excel para continuar a análise em uma planilha.',
      ],
    },
    {
      title: 'Listas de compras',
      intro: 'Listas de compras organizam produtos, quantidades, preços e participantes até o encerramento da compra.',
      steps: [
        'Toque em Nova lista, informe nome, data e observação, e escolha entre uma lista particular ou compartilhada.',
        'Para compartilhar, selecione os amigos participantes. Eles poderão adicionar e editar itens enquanto a lista estiver aberta.',
        'Abra a lista e adicione produto, quantidade e preço unitário. O sistema calcula o total de cada item e o valor final.',
        'Use as ações do item para editar ou remover. O histórico de listas ajuda a sugerir nomes de produtos já utilizados.',
        'Ao concluir a compra, toque em Finalizar compra. Para encerrar sem concluir, use Cancelar lista.',
        'Listas finalizadas ou canceladas ficam somente para consulta e podem ser duplicadas para iniciar uma nova compra.',
      ],
    },
    {
      title: 'Amigos e compartilhamentos',
      intro: 'A área Amigos conecta usuários pelo ID de amizade e permite enviar lançamentos financeiros para aprovação.',
      steps: [
        'Abra Meu perfil, copie seu ID de amizade e envie-o à pessoa que deseja adicionar.',
        'Em Amigos, toque em Adicionar amigo, informe o ID recebido e envie a solicitação.',
        'Na área de solicitações, aceite ou recuse convites. Depois de aceito, o contato aparece na sua lista de amigos.',
        'Abra um amigo e toque em Lançar para criar uma transação compartilhada, informando tipo, descrição, valor, vencimento e categoria.',
        'O destinatário pode aprovar ou recusar. Quando aprovado, o lançamento entra no controle financeiro correspondente.',
        'Use os filtros da conversa para acompanhar solicitações pendentes e o histórico. Também é possível cancelar um envio ainda pendente ou remover uma amizade.',
      ],
    },
    {
      title: 'Notificações',
      intro: 'Notificações avisam sobre convites, compartilhamentos e atualizações importantes mesmo fora da tela atual.',
      steps: [
        'Abra Notificações pelo sino da barra superior ou pelo menu de ajustes.',
        'Ative todas as notificações de uma vez ou escolha individualmente quais tipos deseja receber.',
        'Quando o navegador solicitar permissão, confirme para permitir os avisos do dispositivo.',
        'No iPhone, instale o RubyLife na Tela de Início e abra o aplicativo instalado antes de ativar notificações push.',
        'Use Limpar histórico para remover os avisos exibidos na central sem alterar seus dados financeiros.',
      ],
    },
    {
      title: 'Perfil e configurações',
      intro: 'O perfil concentra sua identificação, foto, ID de amizade e preferências gerais do aplicativo.',
      steps: [
        'Em Meu perfil, atualize nome, telefone e foto. O e-mail identifica a conta e não é alterado nesse formulário.',
        'Toque na foto para escolher uma imagem, confira a prévia e use Salvar alterações para confirmar.',
        'Copie o ID de amizade quando quiser receber convites de outros usuários.',
        'Abra o menu do perfil na barra superior para alternar entre modo claro e escuro, acessar notificações ou voltar à Central de Ajuda.',
        'Use Sair somente quando desejar encerrar a sessão neste dispositivo.',
      ],
    },
  ];
  const [activeModule, setActiveModule] = useState('0');
  const activeSection = sections[Number(activeModule)] ?? sections[0];

  return (
    <section className="help-page help-center-page">
      <header className="help-center-hero">
        <div className="help-center-kicker"><BookOpen size={16} /> Manual prático do sistema</div>
        <h1>Central de Ajuda</h1>
        <p>Aprenda a usar cada área do RubyLife, do primeiro cadastro aos recursos compartilhados.</p>
      </header>

      <div className="help-center-filters help-center-filters--single">
        <label className="help-center-field">
          <span>Módulo</span>
          <select value={activeModule} onChange={(event) => setActiveModule(event.target.value)}>
            {sections.map((section, index) => <option key={section.title} value={String(index)}>{section.title}</option>)}
          </select>
        </label>
      </div>

      <article className="help-center-guide" aria-live="polite">
        <div className="help-center-guide-heading"><h2>Passo a passo</h2></div>
        <div className="help-center-guide-context">
          <strong>{activeSection.title}</strong>
          <p>{activeSection.intro}</p>
        </div>
        <ol className="help-center-steps">
          {activeSection.steps.map((item, index) => (
            <li key={item}>
              <span className="help-center-step-number">{index + 1}</span>
              <span>{item}</span>
            </li>
          ))}
        </ol>
      </article>
    </section>
  );
}
function DeviceNotificationsCard() {
  const [status, setStatus] = useState<PushStatus | null>(null);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);
  const [preferences, setPreferences] = useState<PushPreferences>(() => {
    try {
      const raw = localStorage.getItem('rubylife-push-preferences');
      return raw ? { ...DEFAULT_PUSH_PREFERENCES, ...JSON.parse(raw) } : DEFAULT_PUSH_PREFERENCES;
    } catch {
      return DEFAULT_PUSH_PREFERENCES;
    }
  });

  async function refreshStatus() {
    setStatus(await getPushStatus());
  }

  useEffect(() => {
    void refreshStatus();
  }, []);



  const allActive = PUSH_PREFERENCE_LABELS.every((item) => preferences[item.key]);

  async function toggleAllPreferences() {
    setBusy(true);
    setMessage(null);

    const nextState = !allActive;

    if (nextState && status && !status.subscribed && status.supported && status.publicKeyConfigured) {
      try {
        await enablePushNotifications();
        await refreshStatus();
      } catch (error) {
        console.warn('Browser push subscription failed:', error);
      }
    }

    const next = { ...preferences };
    PUSH_PREFERENCE_LABELS.forEach((item) => {
      next[item.key] = nextState;
    });

    setPreferences(next);
    localStorage.setItem('rubylife-push-preferences', JSON.stringify(next));
    try {
      await savePushPreferences(next);
      setMessage({ type: 'success', text: nextState ? 'Todas as notificações foram ativadas.' : 'Todas as notificações foram desativadas.' });
    } catch {
      setMessage({ type: 'error', text: 'Preferências atualizadas neste dispositivo, mas não foi possível sincronizar agora.' });
    } finally {
      setBusy(false);
    }
  }

  async function togglePreference(key: keyof PushPreferences) {
    const next = { ...preferences, [key]: !preferences[key] };
    setPreferences(next);
    localStorage.setItem('rubylife-push-preferences', JSON.stringify(next));
    try {
      await savePushPreferences(next);
      setMessage({ type: 'success', text: 'Preferências de notificação salvas.' });
    } catch {
      setMessage({ type: 'error', text: 'Preferência salva neste dispositivo, mas não foi possível sincronizar agora.' });
    }
  }

  const iosHint = status?.platform === 'ios' && !status.standalone
    ? 'Para receber notificações no iPhone, adicione o RubyLife à Tela de Início e depois toque em “Ativar notificações”.'
    : null;
  const unsupported = status && !status.supported;
  const missingKey = status?.supported && !status.publicKeyConfigured;

  return (
    <section className="push-settings-card">
      <div className="push-settings-head">
        <span className="push-settings-icon"><Bell size={20} /></span>
        <div>
          <strong>Notificações do dispositivo</strong>
          <small>Receba avisos importantes do RubyLife mesmo quando o sistema estiver fechado.</small>
        </div>
      </div>

      <div className="push-preferences-list" style={{ margin: '16px 0 20px' }}>
        <button type="button" className="push-preference-row" onClick={toggleAllPreferences} disabled={busy} style={{ width: '100%' }}>
          <span><strong>Ativar todas as notificações</strong></span>
          <span className={`ios-switch ${allActive ? 'ios-switch--on' : ''}`} aria-hidden="true"><span className="ios-switch-knob" /></span>
        </button>
      </div>

      {iosHint ? <p className="push-settings-note">{iosHint}</p> : null}
      {unsupported ? <p className="push-settings-note push-settings-note--error">Este navegador ou dispositivo não suporta notificações push para PWA.</p> : null}
      {missingKey ? <p className="push-settings-note push-settings-note--error">A chave pública VAPID ainda não está configurada no ambiente do app.</p> : null}
      {message ? <p className={`push-settings-message push-settings-message--${message.type}`}>{message.text}</p> : null}

      <div className="push-preferences-list">
        {PUSH_PREFERENCE_LABELS.map((item) => (
          <button type="button" key={item.key} className="push-preference-row" onClick={() => togglePreference(item.key)}>
            <span><strong>{item.label}</strong><small>{item.description}</small></span>
            <span className={`ios-switch ${preferences[item.key] ? 'ios-switch--on' : ''}`} aria-hidden="true"><span className="ios-switch-knob" /></span>
          </button>
        ))}
      </div>
    </section>
  );
}
function ProfilePage({ fullName, email, publicFriendId, phone, avatarUrl, onSave, categories, onNewCategory, onEditCategory }: {
  fullName: string;
  email: string;
  publicFriendId: string;
  phone: string;
  avatarUrl?: string | null;
  onSave: (profile: { fullName: string; phone: string; avatarUrl: string | null }) => Promise<void>;
  categories: CategoryItem[];
  onNewCategory: () => void;
  onEditCategory: (category: CategoryItem) => void;
}) {
  const [activeTab, setActiveTab] = useState<'profile' | 'categories'>('profile');
  const [name, setName] = useState(fullName);
  const [phoneNumber, setPhoneNumber] = useState(phone);
  const [photo, setPhoto] = useState<string | null>(avatarUrl ?? null);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [copiedFriendId, setCopiedFriendId] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setName(fullName);
    setPhoneNumber(phone);
    setPhoto(avatarUrl ?? null);
  }, [fullName, phone, avatarUrl]);

  function handlePhotoSelect(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const maxSize = 220;
        let width = img.width;
        let height = img.height;
        if (width > height) {
          if (width > maxSize) {
            height = Math.round((height * maxSize) / width);
            width = maxSize;
          }
        } else {
          if (height > maxSize) {
            width = Math.round((width * maxSize) / height);
            height = maxSize;
          }
        }
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(img, 0, 0, width, height);
          const dataUrl = canvas.toDataURL('image/jpeg', 0.88);
          setPhoto(dataUrl);
          setMessage({ type: 'success', text: 'Nova foto selecionada! Clique em Salvar alterações para armazenar no Supabase.' });
        }
      };
      img.src = e.target?.result as string;
    };
    reader.readAsDataURL(file);
  }

  async function copyFriendId() {
    try {
      await navigator.clipboard?.writeText(publicFriendId);
      setCopiedFriendId(true);
      window.setTimeout(() => setCopiedFriendId(false), 1800);
    } catch {
      setMessage({ type: 'error', text: 'Não foi possível copiar o ID automaticamente.' });
    }
  }

  async function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const normalizedName = name.trim();
    const normalizedPhone = phoneNumber.trim();
    setMessage(null);

    if (normalizedName.length < 2) {
      setMessage({ type: 'error', text: 'Informe seu nome completo.' });
      return;
    }
    if (normalizedPhone && !/^[0-9()+\-\s]+$/.test(normalizedPhone)) {
      setMessage({ type: 'error', text: 'Informe um telefone válido.' });
      return;
    }

    setSaving(true);
    try {
      await onSave({ fullName: normalizedName, phone: normalizedPhone, avatarUrl: photo });
      setMessage({ type: 'success', text: 'Perfil e foto salvos com sucesso no Supabase!' });
    } catch (error) {
      console.error('Falha ao atualizar perfil:', error);
      setMessage({ type: 'error', text: 'Não foi possível salvar seu perfil. Tente novamente.' });
    } finally {
      setSaving(false);
    }
  }

  return (
    <section className="profile-page">
      <header className="profile-page-header">
        <span className="profile-page-kicker"><UserRound size={15} /> Conta pessoal</span>
        <h1 className="page-title">Meu perfil</h1>
        <p className="page-subtitle">Personalize sua foto de perfil e mantenha seus dados cadastrais atualizados no Supabase.</p>
      </header>

      {/* Abas — visíveis apenas no mobile */}
      <div className="profile-mobile-tabs">
        <button type="button" className={activeTab === 'profile' ? 'active' : ''} onClick={() => setActiveTab('profile')}>
          <UserRound size={16} /> Perfil
        </button>
        <button type="button" className={activeTab === 'categories' ? 'active' : ''} onClick={() => setActiveTab('categories')}>
          <Tags size={16} /> Categorias
        </button>
      </div>

      {activeTab === 'categories' ? (
        <div className="profile-categories-view">
          <div className="profile-categories-header">
            <span>{categories.length} {categories.length === 1 ? 'categoria' : 'categorias'}</span>
            <button type="button" className="page-primary-action" onClick={onNewCategory}><Plus size={16} /> Nova categoria</button>
          </div>
          {(['income', 'expense'] as const).map((kind) => {
            const group = categories.filter((c) => c.kind === kind);
            if (!group.length) return null;
            return (
              <div key={kind} className="profile-category-group">
                <div className="profile-category-group-label">{kind === 'income' ? 'Receitas' : 'Despesas'}</div>
                <div className="profile-category-list">
                  {group.map((cat) => (
                    <button key={cat.id} type="button" className="profile-category-row" onClick={() => onEditCategory(cat)}>
                      <span className="profile-category-icon" style={{ background: cat.color + '22', color: cat.color }}>
                        <CategoryIconGraphic icon={cat.icon} size={16} />
                      </span>
                      <span className="profile-category-name">{cat.name}</span>
                      <ChevronRight size={16} className="profile-category-chevron" />
                    </button>
                  ))}
                </div>
              </div>
            );
          })}
          {categories.length === 0 && (
            <div className="profile-categories-empty">
              <Tags size={32} />
              <strong>Nenhuma categoria</strong>
              <p>Crie categorias para organizar suas receitas e despesas.</p>
            </div>
          )}
        </div>
      ) : null}

      {activeTab === 'profile' ? <>
        <section className="profile-friend-id-card">
          <div>
            <span>ID de amizade</span>
            <strong>{publicFriendId}</strong>
            <small>Compartilhe este ID para receber convites de amizade.</small>
          </div>
          <button type="button" className="page-secondary-action" onClick={copyFriendId}><Save size={15} /> {copiedFriendId ? 'Copiado' : 'Copiar ID'}</button>
        </section>

        <form className="profile-registration-card" onSubmit={submit}>

          <div className="profile-registration-layout">
            <div className="profile-photo-column">
              <span className="profile-photo-title">FOTO PERFIL</span>

              <div className="profile-photo-circle" onClick={() => fileInputRef.current?.click()} title="Clique para alterar foto">
                {photo ? (
                  <img src={photo} alt="Foto de perfil" className="profile-photo-preview" />
                ) : (
                  <UserRound size={68} className="profile-photo-icon" />
                )}
              </div>

              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                onChange={handlePhotoSelect}
                style={{ display: 'none' }}
              />

              <div className="profile-photo-actions">
                <button
                  type="button"
                  className="profile-photo-btn"
                  onClick={() => fileInputRef.current?.click()}
                >
                  <Camera size={14} /> {photo ? 'Alterar foto' : 'Adicionar foto'}
                </button>

                {photo ? (
                  <button
                    type="button"
                    className="profile-photo-btn profile-photo-btn--remove"
                    onClick={() => {
                      setPhoto(null);
                      if (fileInputRef.current) fileInputRef.current.value = '';
                      setMessage({ type: 'success', text: 'Foto removida. Clique em Salvar alterações para confirmar no Supabase.' });
                    }}
                  >
                    <Trash2 size={14} /> Remover
                  </button>
                ) : null}
              </div>

              <small className="profile-photo-help">
                Tamanho máximo: 220px altura x 220px largura
              </small>
            </div>

            <div className="profile-data-column">
              <div className="profile-form-grid">
                <label className="profile-field">
                  <span>NOME</span>
                  <div className="profile-input-wrap">
                    <UserRound size={17} />
                    <input
                      value={name}
                      onChange={(event) => { setName(event.target.value); setMessage(null); }}
                      minLength={2}
                      maxLength={160}
                      autoComplete="name"
                      placeholder="Seu nome completo"
                      required
                    />
                  </div>
                </label>

                <label className="profile-field">
                  <span>EMAIL</span>
                  <div className="profile-input-wrap profile-input-wrap--readonly">
                    <Mail size={17} />
                    <input value={email} readOnly aria-readonly="true" />
                  </div>
                </label>

                <label className="profile-field">
                  <span>TELEFONE</span>
                  <div className="profile-input-wrap">
                    <Phone size={17} />
                    <input
                      value={phoneNumber}
                      onChange={(event) => { setPhoneNumber(event.target.value); setMessage(null); }}
                      maxLength={30}
                      inputMode="tel"
                      autoComplete="tel"
                      placeholder="(00) 00000-0000"
                    />
                  </div>
                </label>
              </div>
            </div>
          </div>

          <div className="profile-form-footer">
            <div aria-live="polite">
              {message ? (
                <span className={`profile-save-message profile-save-message--${message.type}`}>
                  {message.type === 'success' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
                  {message.text}
                </span>
              ) : null}
            </div>
            <button type="submit" className="page-primary-action profile-save-button" disabled={saving}>
              <Save size={16} /> {saving ? 'Salvando...' : 'Salvar alterações'}
            </button>
          </div>
        </form>
      </> : null}
    </section>
  );
}
type FriendsPageProps = {
  currentUser: FriendUser;
  users: FriendUser[];
  invitations: FriendshipInvitation[];
  acceptedFriends: FriendUser[];
  sharedTransactions: SharedTransactionRequest[];
  incomeCategories: string[];
  expenseCategories: string[];
  activeTab: 'search' | 'received';
  openSearchSignal?: number;
  backSignal?: number;
  onThreadChange?: (name: string | null) => void;
  onTabChange: (tab: 'search' | 'received') => void;
  onSendInvitation: (user: FriendUser) => string | null;
  onAcceptInvitation: (invitationId: string) => void;
  onDeclineInvitation: (invitationId: string) => void;
  onRemoveFriend: (user: FriendUser) => void;
  onCreateShared: (form: SharedTransactionForm) => string | null;
  onApproveShared: (id: string) => void;
  onDeclineShared: (id: string, reason: string) => void;
  onCancelShared: (id: string) => void;
};

type SharedView = { mine: boolean; headerText: string; statusLabel: string; statusClass: string; pendingForMe: boolean };

function describeShared(item: SharedTransactionRequest, currentUserId: string, friendName: string): SharedView {
  const mine = item.creatorId === currentUserId;
  const typeLabel = item.type === 'income' ? 'receita' : 'despesa';
  let headerText = mine ? `Você criou uma ${typeLabel} para ${friendName}` : `${friendName} criou uma ${typeLabel} para você`;
  let statusLabel = 'Pendente';
  let statusClass = 'pending';
  let pendingForMe = false;
  if (item.status === 'pending') {
    if (mine) { statusLabel = 'Aguardando aprovação'; statusClass = 'pending'; }
    else { statusLabel = 'Pendente'; statusClass = 'pending'; pendingForMe = true; }
  } else if (item.status === 'approved') {
    headerText = item.receiverId === currentUserId ? 'Você aprovou a transação' : `${friendName} aprovou a transação`;
    statusLabel = 'Aprovada'; statusClass = 'approved';
  } else if (item.status === 'declined') {
    headerText = item.receiverId === currentUserId ? 'Você recusou a transação' : `${friendName} recusou a transação`;
    statusLabel = 'Recusada'; statusClass = 'declined';
  } else if (item.status === 'cancelled') {
    headerText = item.creatorId === currentUserId ? 'Você cancelou a transação' : `${friendName} cancelou a transação`;
    statusLabel = 'Cancelada'; statusClass = 'cancelled';
  }
  return { mine, headerText, statusLabel, statusClass, pendingForMe };
}

function FriendsPage({ currentUser, users, invitations, acceptedFriends, sharedTransactions, incomeCategories, expenseCategories, activeTab, openSearchSignal, backSignal, onThreadChange, onTabChange, onSendInvitation, onAcceptInvitation, onDeclineInvitation, onRemoveFriend, onCreateShared, onApproveShared, onDeclineShared, onCancelShared }: FriendsPageProps) {
  const [searchOpen, setSearchOpen] = useState(false);
  const [requestsOpen, setRequestsOpen] = useState(false);
  const [conversationQuery, setConversationQuery] = useState('');
  const [openFriendId, setOpenFriendId] = useState<string | null>(null);
  const [threadFilter, setThreadFilter] = useState<'pending' | 'history'>('pending');
  const [composerOpen, setComposerOpen] = useState(false);
  const [composerFriendId, setComposerFriendId] = useState<string | null>(null);
  const [declineTarget, setDeclineTarget] = useState<SharedTransactionRequest | null>(null);
  const [mobileTab, setMobileTab] = useState<'amigos' | 'solicitacoes'>('amigos');

  useEffect(() => {
    if (openSearchSignal) setSearchOpen(true);
  }, [openSearchSignal]);

  const receivedInvitations = useMemo(() => (
    invitations.filter((invitation) => invitation.receiverId === currentUser.id && invitation.status === 'pending')
  ), [invitations, currentUser.id]);

  const userById = useMemo(() => new Map(users.map((user) => [user.id, user] as const)), [users]);
  function resolveUser(id: string): FriendUser {
    return userById.get(id) ?? { id, publicFriendId: publicFriendIdForUser(id), name: 'Usuário do sistema', email: 'email não disponível' };
  }

  useEffect(() => {
    if (activeTab === 'received') setRequestsOpen(true);
  }, [activeTab]);

  function sharedWith(friendId: string) {
    return sharedTransactions
      .filter((item) => (
        (item.creatorId === currentUser.id && item.receiverId === friendId) ||
        (item.creatorId === friendId && item.receiverId === currentUser.id)
      ))
      .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  function pendingForMeCount(friendId: string) {
    return sharedTransactions.filter((item) => (
      item.creatorId === friendId && item.receiverId === currentUser.id && item.status === 'pending'
    )).length;
  }

  const conversations = useMemo(() => {
    const normalized = conversationQuery.trim().toLowerCase();
    return acceptedFriends
      .filter((friend) => !normalized || friend.name.toLowerCase().includes(normalized) || friend.publicFriendId.toLowerCase().includes(normalized))
      .map((friend) => {
        const items = sharedWith(friend.id);
        return {
          friend,
          pending: pendingForMeCount(friend.id),
          lastAt: items[0]?.createdAt ?? '',
        };
      })
      .sort((a, b) => {
        if ((b.pending > 0 ? 1 : 0) !== (a.pending > 0 ? 1 : 0)) return (b.pending > 0 ? 1 : 0) - (a.pending > 0 ? 1 : 0);
        if (a.lastAt !== b.lastAt) return b.lastAt.localeCompare(a.lastAt);
        return a.friend.name.localeCompare(b.friend.name);
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [acceptedFriends, sharedTransactions, conversationQuery, currentUser.id]);

  const openFriend = openFriendId ? acceptedFriends.find((friend) => friend.id === openFriendId) ?? null : null;

  useEffect(() => {
    onThreadChange?.(openFriend ? shortUserName(openFriend.name) : null);
    return () => onThreadChange?.(null);
  }, [openFriend, onThreadChange]);

  useEffect(() => {
    if (!backSignal) return;
    setOpenFriendId(null);
    setThreadFilter('pending');
    onThreadChange?.(null);
  }, [backSignal, onThreadChange]);

  if (openFriend) {
    const items = sharedWith(openFriend.id);
    const filtered = threadFilter === 'pending'
      ? items.filter((item) => item.status === 'pending')
      : items.filter((item) => item.status !== 'pending');
    return (
      <>
        <section className="dm-thread-header">
          <button type="button" className="dm-back" onClick={() => { setOpenFriendId(null); setThreadFilter('pending'); onThreadChange?.(null); }} aria-label="Voltar">
            <ChevronLeft size={20} />
          </button>
          <FriendAvatar user={openFriend} />
          <div className="dm-thread-heading">
            <strong>{shortUserName(openFriend.name)}</strong>
            <span>Transações compartilhadas</span>
          </div>
          <button type="button" className="friend-action friend-action--danger dm-thread-remove-friend" title="Excluir amizade" aria-label="Excluir amizade" onClick={() => onRemoveFriend(openFriend)}>
            <UserX size={15} /> <span className="dm-thread-remove-text">Excluir amizade</span>
          </button>
        </section>

        <div className="dm-thread-mobile-actions">
          <button type="button" className="friend-action friend-action--danger" onClick={() => onRemoveFriend(openFriend)}>
            <UserX size={15} /> Excluir amizade
          </button>
        </div>

        <div className="dm-thread-filter" role="tablist" aria-label="Filtrar transações">
          <button type="button" role="tab" aria-selected={threadFilter === 'pending'} className={threadFilter === 'pending' ? 'active' : ''} onClick={() => setThreadFilter('pending')}>Pendentes</button>
          <button type="button" role="tab" aria-selected={threadFilter === 'history'} className={threadFilter === 'history' ? 'active' : ''} onClick={() => setThreadFilter('history')}>Histórico</button>
        </div>

        <section className="dm-thread-body">
          {filtered.length ? filtered.map((item) => {
            const view = describeShared(item, currentUser.id, shortUserName(openFriend.name));
            return (
              <article key={item.id} className={`shared-msg shared-msg--${view.mine ? 'mine' : 'theirs'}`}>
                <div className="shared-card">
                  <span className="shared-card-head">{view.headerText}</span>
                  <div className="shared-card-body">
                    <strong className="shared-card-desc">{item.description}</strong>
                    <span className="shared-card-meta">Vencimento: {formatDate(item.dueDate)}</span>
                    {item.category ? <span className="shared-card-meta">{item.category}</span> : null}
                    {item.note ? <span className="shared-card-meta shared-card-note"><strong>Observação:</strong> {item.note}</span> : null}
                    {item.status === 'declined' && item.declineReason ? <span className="shared-card-meta shared-card-decline-reason"><strong>Motivo da recusa:</strong> {item.declineReason}</span> : null}
                  </div>
                  <div className="shared-card-foot">
                    <strong className={`shared-card-value shared-card-value--${item.type}`}>{item.type === 'income' ? '+ ' : '- '}{formatCurrency(item.amount)}</strong>
                    <span className={`friend-status friend-status--${view.statusClass}`}>{view.statusLabel}</span>
                  </div>
                  {item.status === 'pending' && item.receiverId === currentUser.id ? (
                    <div className="shared-card-actions">
                      <button type="button" className="friend-action friend-action--danger" onClick={() => setDeclineTarget(item)}><X size={15} /> Recusar</button>
                      <button type="button" className="friend-action" onClick={() => onApproveShared(item.id)}><CheckCircle2 size={15} /> Aprovar</button>
                    </div>
                  ) : item.status === 'pending' && item.creatorId === currentUser.id ? (
                    <div className="shared-card-actions">
                      <button type="button" className="friend-action friend-action--danger" onClick={() => onCancelShared(item.id)}><X size={15} /> Cancelar</button>
                    </div>
                  ) : null}
                </div>
              </article>
            );
          }) : (
            <div className="friends-empty">
              <ReceiptText size={26} />
              <strong>{threadFilter === 'pending' ? 'Nenhuma transação pendente.' : 'Nenhum histórico ainda.'}</strong>
              <span>{threadFilter === 'pending' ? 'As aprovações pendentes aparecerão aqui.' : 'Transações aprovadas, recusadas ou canceladas aparecerão aqui.'}</span>
            </div>
          )}
        </section>

        <button type="button" className="dm-fab" onClick={() => setComposerOpen(true)}>
          <Plus size={18} /> Nova transação
        </button>

        {composerOpen ? (
          <SharedTransactionModal
            friends={acceptedFriends}
            lockedFriend={openFriend}
            incomeCategories={incomeCategories}
            expenseCategories={expenseCategories}
            onClose={() => setComposerOpen(false)}
            onCreate={(form) => onCreateShared(form)}
          />
        ) : null}
        {declineTarget ? (
          <DeclineSharedTransactionModal
            item={declineTarget}
            onClose={() => setDeclineTarget(null)}
            onConfirm={(reason) => { onDeclineShared(declineTarget.id, reason); setDeclineTarget(null); }}
          />
        ) : null}
      </>
    );
  }

  return (
    <>
      <section className="page-header page-header-split friends-page-header friends-overview-header">
        <div className="page-header-left">
          <h1 className="page-title">Amigos</h1>
          <span className="page-subtitle">Gerencie suas conexões e acompanhe as transações compartilhadas.</span>
        </div>
        <div className="dm-header-actions">
          <button type="button" className="page-secondary-action dm-header-button friends-requests-trigger" onClick={() => { setRequestsOpen(true); onTabChange('received'); }}>
            <Mail size={16} /> Solicitações
            {receivedInvitations.length ? <span className="dm-badge">{receivedInvitations.length}</span> : null}
          </button>
          <button type="button" className="page-primary-action dm-header-button" onClick={() => setSearchOpen(true)}>
            <UserPlus size={16} /> Adicionar amigo
          </button>
          <button type="button" className="friends-mobile-add-btn" onClick={() => setSearchOpen(true)} aria-label="Adicionar amigo">
            <UserPlus size={20} />
          </button>
        </div>
      </section>

      <div className="friends-mobile-tabs" role="tablist">
        <button type="button" role="tab" aria-selected={mobileTab === 'amigos'} className={mobileTab === 'amigos' ? 'active' : ''} onClick={() => setMobileTab('amigos')}>Amigos</button>
        <button type="button" role="tab" aria-selected={mobileTab === 'solicitacoes'} className={mobileTab === 'solicitacoes' ? 'active' : ''} onClick={() => setMobileTab('solicitacoes')}>
          Solicitações
          {receivedInvitations.length ? <span className="friends-mobile-tab-badge">{receivedInvitations.length}</span> : null}
        </button>
      </div>

      {mobileTab === 'solicitacoes' ? (
        receivedInvitations.length ? (
          <div className="friend-requests-inline">
            {receivedInvitations.map((invitation) => {
              const requester = resolveUser(invitation.requesterId);
              return (
                <article className="friend-request-inline-card" key={invitation.id}>
                  <FriendAvatar user={requester} />
                  <div className="friend-request-inline-main">
                    <strong>{requester.name}</strong>
                    <span>Pedido de amizade</span>
                  </div>
                  <div className="friend-request-inline-actions">
                    <button type="button" className="friend-round-action friend-round-action--decline" onClick={() => onDeclineInvitation(invitation.id)} aria-label={`Recusar pedido de ${requester.name}`}><X size={18} /></button>
                    <button type="button" className="friend-round-action friend-round-action--accept" onClick={() => onAcceptInvitation(invitation.id)} aria-label={`Aceitar pedido de ${requester.name}`}><Check size={18} /></button>
                  </div>
                </article>
              );
            })}
          </div>
        ) : (
          <div className="friends-empty"><span className="friends-empty-icon"><Mail size={24} /></span><strong>Nenhuma solicitação pendente.</strong></div>
        )
      ) : null}

      {mobileTab === 'amigos' ? (
        <section className="friends-overview-panel">
          <div className="dm-search friends-overview-search">
            <Search size={18} />
            <input value={conversationQuery} onChange={(event) => setConversationQuery(event.target.value)} placeholder="Pesquisar por nome ou ID" />
          </div>
          <section className="dm-list">
            {conversations.length ? conversations.map(({ friend, pending }) => (
              <div className="dm-conversation" key={friend.id}>
                <button type="button" className="dm-conversation-identity" onClick={() => { setOpenFriendId(friend.id); setThreadFilter('pending'); }}>
                  <FriendAvatar user={friend} />
                  <div className="dm-conversation-main">
                    <strong>{shortUserName(friend.name)}</strong>
                  </div>
                  {pending ? <span className="dm-badge">{pending}</span> : <ChevronRight size={18} className="dm-conversation-chevron" />}
                </button>
                <button type="button" className="friend-round-action friend-round-action--decline" onClick={(e) => { e.stopPropagation(); onRemoveFriend(friend); }} title="Excluir amizade" aria-label={`Excluir ${friend.name}`} style={{ width: 34, height: 34 }}>
                  <UserX size={15} />
                </button>
                <button type="button" className="dm-conversation-launch" onClick={() => setComposerFriendId(friend.id)}>
                  <Plus size={15} /> <span>Lançar</span>
                </button>
              </div>
            )) : (
              <div className="friends-empty">
                <span className="friends-empty-icon"><Users size={24} /></span>
                <strong>{acceptedFriends.length ? 'Nenhum amigo encontrado.' : 'Você ainda não tem amigos.'}</strong>
                <span>{acceptedFriends.length ? 'Tente buscar por outro nome ou pelo ID numérico.' : 'Adicione amigos pelo ID numérico exclusivo de cada usuário.'}</span>
                {acceptedFriends.length ? null : <button type="button" className="button-primary" onClick={() => setSearchOpen(true)}><UserPlus size={15} /> Adicionar amigo</button>}
              </div>
            )}
          </section>
        </section>
      ) : null}

      {composerFriendId ? (() => {
        const composerFriend = acceptedFriends.find((friend) => friend.id === composerFriendId);
        if (!composerFriend) return null;
        return (
          <SharedTransactionModal
            friends={acceptedFriends}
            lockedFriend={composerFriend}
            incomeCategories={incomeCategories}
            expenseCategories={expenseCategories}
            onClose={() => setComposerFriendId(null)}
            onCreate={(form) => onCreateShared(form)}
          />
        );
      })() : null}

      {searchOpen ? (
        <FriendSearchModal
          currentUser={currentUser}
          users={users}
          invitations={invitations}
          onSendInvitation={onSendInvitation}
          onAcceptInvitation={onAcceptInvitation}
          onDeclineInvitation={onDeclineInvitation}
          onRemoveFriend={onRemoveFriend}
          onClose={() => setSearchOpen(false)}
        />
      ) : null}

      {requestsOpen ? (
        <FriendRequestsModal
          currentUser={currentUser}
          users={users}
          invitations={invitations}
          onAcceptInvitation={onAcceptInvitation}
          onDeclineInvitation={onDeclineInvitation}
          onClose={() => { setRequestsOpen(false); onTabChange('search'); }}
        />
      ) : null}
    </>
  );
}

type FriendsManageViewProps = {
  currentUser: FriendUser;
  users: FriendUser[];
  invitations: FriendshipInvitation[];
  activeTab: 'search' | 'received';
  onTabChange: (tab: 'search' | 'received') => void;
  onSendInvitation: (user: FriendUser) => string | null;
  onAcceptInvitation: (invitationId: string) => void;
  onDeclineInvitation: (invitationId: string) => void;
  onRemoveFriend: (user: FriendUser) => void;
  onBack: () => void;
  showHeader?: boolean;
  showTabs?: boolean;
  showOwnId?: boolean;
};

function FriendsManageView({ currentUser, users, invitations, activeTab, onTabChange, onSendInvitation, onAcceptInvitation, onDeclineInvitation, onRemoveFriend, onBack, showHeader = true, showTabs = true, showOwnId = true }: FriendsManageViewProps) {
  const [query, setQuery] = useState('');
  const [message, setMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [idCopied, setIdCopied] = useState(false);
  const [foundUser, setFoundUser] = useState<FriendUser | null>(null);
  const [searchLoading, setSearchLoading] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);
  const userById = useMemo(() => new Map(users.map((user) => [user.id, user] as const)), [users]);

  function resolveUser(id: string): FriendUser {
    return userById.get(id) ?? { id, publicFriendId: publicFriendIdForUser(id), name: 'Usuário do sistema', email: 'email não disponível' };
  }

  const receivedInvitations = invitations.filter((invitation) => invitation.receiverId === currentUser.id && invitation.status === 'pending');
  const searchedSelf = Boolean(foundUser && foundUser.id === currentUser.id);
  const pendingSentToFound = foundUser ? invitations.find((invitation) => invitation.requesterId === currentUser.id && invitation.receiverId === foundUser.id && invitation.status === 'pending') : null;
  const pendingReceivedFromFound = foundUser ? invitations.find((invitation) => invitation.requesterId === foundUser.id && invitation.receiverId === currentUser.id && invitation.status === 'pending') : null;
  const foundIsFriend = foundUser ? areUsersFriends(currentUser.id, foundUser.id, invitations) : false;

  async function submitSearch(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setMessage(null);
    setFoundUser(null);
    setHasSearched(false);
    const normalized = normalizePublicFriendId(query);
    if (!normalized) { setMessage({ type: 'error', text: 'Digite o ID do amigo.' }); return; }
    setSearchLoading(true);
    try {
      const result = await searchUserByFriendId(normalized);
      setHasSearched(true);
      if (!result) return;
      if (result.id === currentUser.id) { setMessage({ type: 'error', text: 'Você não pode pesquisar ou convidar a si mesmo.' }); return; }
      setFoundUser(result);
    } catch {
      setMessage({ type: 'error', text: 'Erro ao buscar usuário. Tente novamente.' });
    } finally {
      setSearchLoading(false);
    }
  }

  function inviteFoundUser(user: FriendUser) {
    const error = onSendInvitation(user);
    setMessage(error ? { type: 'error', text: error } : { type: 'success', text: 'Convite de amizade enviado com sucesso.' });
  }

  function acceptInvitation(invitationId: string) {
    onAcceptInvitation(invitationId);
    setMessage({ type: 'success', text: 'Agora vocês são amigos.' });
  }

  function declineInvitation(invitationId: string) {
    onDeclineInvitation(invitationId);
    setMessage({ type: 'success', text: 'Solicitação recusada.' });
  }

  return (
    <>
      {showHeader ? <section className="page-header page-header-split friends-page-header">
        <div className="page-header-left page-header-left--with-back">
          <button type="button" className="dm-back" onClick={onBack} aria-label="Voltar">
            <ChevronLeft size={20} />
          </button>
          <div>
            <h1 className="page-title">Amigos</h1>
            <span className="page-subtitle">Busque pessoas pelo ID de amizade e aprove somente solicitações recebidas.</span>
          </div>
        </div>
      </section> : null}

      {showOwnId ? <section className="friend-own-id-card">
        <div className="friend-own-id-icon"><UserRound size={20} /></div>
        <div><span>Seu ID de amizade</span><strong>{currentUser.publicFriendId}</strong><small>Envie este número para quem deseja adicionar você.</small></div>
        <button type="button" className="page-secondary-action" onClick={async () => { await navigator.clipboard?.writeText(currentUser.publicFriendId); setIdCopied(true); window.setTimeout(() => setIdCopied(false), 1800); }}><Copy size={15} /> {idCopied ? 'Copiado' : 'Copiar ID'}</button>
      </section> : null}

      {showTabs ? <div className="friends-tabs" role="tablist" aria-label="Amigos">
        <button type="button" className={activeTab === 'search' ? 'active' : ''} role="tab" aria-selected={activeTab === 'search'} onClick={() => { onTabChange('search'); setMessage(null); }}><Search size={16} /> Buscar amigo</button>
        <button type="button" className={activeTab === 'received' ? 'active' : ''} role="tab" aria-selected={activeTab === 'received'} onClick={() => { onTabChange('received'); setMessage(null); }}><Mail size={16} /> Solicitações recebidas{receivedInvitations.length ? <span>{receivedInvitations.length}</span> : null}</button>
      </div> : null}

      {message ? (
        <div className={`friends-feedback friends-feedback--${message.type}`} role="status">
          {message.type === 'success' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
          <span>{message.text}</span>
        </div>
      ) : null}

      {activeTab === 'search' ? (
        <section className="friend-id-search-panel">
          <form className="friend-id-search-form" onSubmit={submitSearch}>
            <label className="friend-id-search-field">
              <span>Buscar por ID</span>
              <div><Search size={18} /><input value={query} onChange={(event) => { setQuery(event.target.value); setMessage(null); }} placeholder="Digite o ID de 6 dígitos ou UUID" autoComplete="off" /></div>
            </label>
            <button type="submit" className="page-primary-action" disabled={searchLoading}><Search size={16} /> {searchLoading ? 'Buscando…' : 'Buscar'}</button>
          </form>

          {hasSearched && !foundUser && !searchLoading ? (
            <div className="friend-search-empty"><AlertCircle size={24} /><strong>Nenhum usuário encontrado com esse ID.</strong><span>Verifique se o ID foi digitado corretamente.</span></div>
          ) : null}

          {foundUser && !searchedSelf ? (
            <article className="friend-result-card">
              <FriendAvatar user={foundUser} />
              <div className="friend-result-main">
                <strong>{foundUser.name}</strong>
                <span>ID: {foundUser.publicFriendId}</span>
                {foundUser.email ? <small>{foundUser.email}</small> : null}
              </div>
              <div className="friend-result-actions">
                {foundIsFriend ? (
                  <button type="button" className="friend-action friend-action--danger" onClick={() => onRemoveFriend(foundUser)}><UserX size={15} /> Remover amizade</button>
                ) : pendingSentToFound ? (
                  <span className="friend-status friend-status--pending">Convite enviado</span>
                ) : pendingReceivedFromFound ? (
                  <><button type="button" className="friend-action friend-action--danger" onClick={() => declineInvitation(pendingReceivedFromFound.id)}><X size={15} /> Recusar</button><button type="button" className="friend-action" onClick={() => acceptInvitation(pendingReceivedFromFound.id)}><CheckCircle2 size={15} /> Aceitar convite</button></>
                ) : (
                  <button type="button" className="friend-action" onClick={() => inviteFoundUser(foundUser)}><UserPlus size={15} /> Convidar para amizade</button>
                )}
              </div>
            </article>
          ) : null}
        </section>
      ) : (
        <section className="friend-requests-panel">
          {receivedInvitations.length ? receivedInvitations.map((invitation) => {
            const requester = resolveUser(invitation.requesterId);
            return (
              <article className="friend-request-card" key={invitation.id}>
                <FriendAvatar user={requester} />
                <div className="friend-result-main"><strong>{requester.name}</strong><span>ID: {requester.publicFriendId}</span><small>Enviado em: {formatDate(invitation.createdAt.slice(0, 10))}</small></div>
                <div className="friend-result-actions"><button type="button" className="friend-action friend-action--danger" onClick={() => declineInvitation(invitation.id)}><X size={15} /> Recusar</button><button type="button" className="friend-action" onClick={() => acceptInvitation(invitation.id)}><CheckCircle2 size={15} /> Aceitar</button></div>
              </article>
            );
          }) : <div className="friend-search-empty"><Mail size={24} /><strong>Nenhuma solicitação recebida.</strong><span>Pedidos pendentes aparecerão aqui.</span></div>}
        </section>
      )}
    </>
  );
}

function FriendSearchModal({ currentUser, users, invitations, onSendInvitation, onAcceptInvitation, onDeclineInvitation, onRemoveFriend, onClose }: {
  currentUser: FriendUser;
  users: FriendUser[];
  invitations: FriendshipInvitation[];
  onSendInvitation: (user: FriendUser) => string | null;
  onAcceptInvitation: (invitationId: string) => void;
  onDeclineInvitation: (invitationId: string) => void;
  onRemoveFriend: (user: FriendUser) => void;
  onClose: () => void;
}) {
  return (
    <div className="modal-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
      <div className="modal-card friend-search-modal" role="dialog" aria-modal="true" aria-label="Adicionar amigo">
        <div className="modal-header">
          <img src={rubyDiamond} alt="" className="modal-logo" />
          <div className="modal-title-container">
            <h2 className="modal-title">Adicionar amigo</h2>
          </div>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="friend-search-modal-body">
          <FriendsManageView
            currentUser={currentUser}
            users={users}
            invitations={invitations}
            activeTab="search"
            onTabChange={() => undefined}
            onSendInvitation={onSendInvitation}
            onAcceptInvitation={onAcceptInvitation}
            onDeclineInvitation={onDeclineInvitation}
            onRemoveFriend={onRemoveFriend}
            onBack={onClose}
            showHeader={false}
            showTabs={false}
            showOwnId
          />
        </div>
      </div>
    </div>
  );
}

function FriendRequestsModal({ currentUser, users, invitations, onAcceptInvitation, onDeclineInvitation, onClose }: {
  currentUser: FriendUser;
  users: FriendUser[];
  invitations: FriendshipInvitation[];
  onAcceptInvitation: (invitationId: string) => void;
  onDeclineInvitation: (invitationId: string) => void;
  onClose: () => void;
}) {
  return (
    <div className="modal-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
      <div className="modal-card friend-search-modal" role="dialog" aria-modal="true" aria-label="Solicitações recebidas">
        <div className="friend-search-modal-body">
          <div className="friend-modal-head">
            <div><span>Solicitações recebidas</span><strong>Revise quem pediu amizade</strong></div>
            <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
          </div>
          <FriendsManageView
            currentUser={currentUser}
            users={users}
            invitations={invitations}
            activeTab="received"
            onTabChange={() => undefined}
            onSendInvitation={() => null}
            onAcceptInvitation={onAcceptInvitation}
            onDeclineInvitation={onDeclineInvitation}
            onRemoveFriend={() => undefined}
            onBack={onClose}
            showHeader={false}
            showTabs={false}
            showOwnId={false}
          />
        </div>
      </div>
    </div>
  );
}
function SharedTransactionModal({ friends, lockedFriend, incomeCategories: _incomeCategories, expenseCategories: _expenseCategories, onClose, onCreate }: { friends: FriendUser[]; lockedFriend?: FriendUser; incomeCategories: string[]; expenseCategories: string[]; onClose: () => void; onCreate: (form: SharedTransactionForm) => string | null }) {
  const [form, setForm] = useState<SharedTransactionForm>(() => ({ type: 'expense', friendId: lockedFriend?.id ?? friends[0]?.id ?? '', description: '', amount: '', dueDate: new Date().toISOString().slice(0, 10), category: 'Solicitação veio pelo amigo', note: '' }));
  const [error, setError] = useState('');
  const [submitted, setSubmitted] = useState(false);
  function update<K extends keyof SharedTransactionForm>(key: K, value: SharedTransactionForm[K]) { setForm((current) => ({ ...current, [key]: value })); setError(''); }
  function selectType(type: TransactionType) { setForm((current) => ({ ...current, type, category: 'Solicitação veio pelo amigo' })); setError(''); }
  function submit(event: FormEvent) {
    event.preventDefault();
    const result = onCreate(form);
    if (result) { setError(result); return; }
    setSubmitted(true);
  }
  const confirmedFriend = lockedFriend ?? friends.find((friend) => friend.id === form.friendId);
  if (submitted) {
    return (
      <div className="modal-overlay" onClick={onClose}>
        <div className="modal-card shared-transaction-modal" role="dialog" aria-modal="true" aria-label="Lançamento enviado" onClick={(event) => event.stopPropagation()}>
          <div className="modal-header">
            <img src={rubyDiamond} alt="" className="modal-logo" />
            <div><h2 className="modal-title modal-title-only">Lançamento enviado</h2></div>
            <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
          </div>
          <div className="modal-body shared-transaction-confirm">
            <span className="confirmation-icon confirmation-icon--success"><CheckCircle2 size={30} /></span>
            <strong>Tudo certo!</strong>
            <p>{confirmedFriend ? `${shortUserName(confirmedFriend.name)} vai receber a transação e precisa aprovar.` : 'O amigo selecionado vai receber a transação e precisa aprovar.'}</p>
          </div>
          <div className="modal-actions"><button type="button" className="button-primary" onClick={onClose}>Concluir</button></div>
        </div>
      </div>
    );
  }
  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card shared-transaction-modal" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <img src={rubyDiamond} alt="" className="modal-logo" />
          <div className="modal-title-container">
            <h2 className="modal-title">
              {lockedFriend ? 'Nova transação' : 'Nova transação compartilhada'}
            </h2>
            {lockedFriend ? (
              <span className="modal-subtitle">
                {shortUserName(lockedFriend.name)}
              </span>
            ) : null}
          </div>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body"><div className="modal-form-grid">
          <div className="form-field form-field-full"><span className="form-label">Tipo</span><div className="launch-type-segmented" role="group" aria-label="Tipo da transação"><button type="button" className={`launch-type-segment launch-type-segment--expense${form.type === 'expense' ? ' active' : ''}`} onClick={() => selectType('expense')}><ArrowDownLeft size={17} /> Despesa</button><button type="button" className={`launch-type-segment launch-type-segment--income${form.type === 'income' ? ' active' : ''}`} onClick={() => selectType('income')}><ArrowUpRight size={17} /> Receita</button></div></div>
          {lockedFriend ? null : <label className="form-field form-field-full"><span className="form-label">Amigo relacionado</span><select className="form-input" value={form.friendId} onChange={(event) => update('friendId', event.target.value)} required><option value="">Selecione um amigo</option>{friends.map((friend) => <option key={friend.id} value={friend.id}>{friend.name}</option>)}</select></label>}
          <label className="form-field form-field-full"><span className="form-label">Descrição</span><input className="form-input" value={form.description} maxLength={TRANSACTION_DESCRIPTION_MAX_LENGTH} onChange={(event) => update('description', event.target.value)} placeholder="Ex.: Mercado, reembolso, aluguel..." required /></label>
          <label className="form-field"><span className="form-label">Valor</span><div className="form-input-wrap"><span className="form-input-prefix">R$</span><input className="form-input" value={form.amount} onChange={(event) => update('amount', formatCurrencyInput(event.target.value))} placeholder="0,00" inputMode="decimal" required /></div></label>
          <label className="form-field"><span className="form-label">Vencimento</span><div className="form-input-wrap"><CalendarDays size={16} /><input className="form-input" type="date" value={form.dueDate} onChange={(event) => update('dueDate', event.target.value)} required /></div></label>
          <label className="form-field form-field-full"><span className="form-label form-label-optional">Observação</span><input className="form-input" value={form.note} onChange={(event) => update('note', event.target.value)} placeholder="Opcional" /></label>
          {error ? <div className="form-error form-field-full"><AlertCircle size={15} /> {error}</div> : null}
          {!friends.length ? <div className="form-error form-field-full"><AlertCircle size={15} /> Aceite um convite de amizade antes de criar transações compartilhadas.</div> : null}
        </div></div>
        <div className="modal-actions"><button type="button" className="button-secondary" onClick={onClose}>Cancelar</button><button type="submit" className="button-primary" disabled={!friends.length}>Enviar para aprovação</button></div>
      </form>
    </div>
  );
}

function DeclineSharedTransactionModal({ item, onClose, onConfirm }: { item: SharedTransactionRequest; onClose: () => void; onConfirm: (reason: string) => void }) {
  const [reason, setReason] = useState('');
  const [error, setError] = useState('');
  const trimmedReason = reason.trim();

  function submit(event: FormEvent) {
    event.preventDefault();
    if (!trimmedReason) {
      setError('Informe o motivo da recusa.');
      return;
    }
    onConfirm(trimmedReason);
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card decline-shared-modal" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <img src={rubyDiamond} alt="" className="modal-logo" />
          <div><h2 className="modal-title modal-title-only">Motivo da recusa</h2></div>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body decline-shared-body">
          <div className="decline-shared-summary">
            <span>{item.type === 'income' ? 'Receita' : 'Despesa'}</span>
            <strong>{item.description}</strong>
            <em>{item.type === 'income' ? '+ ' : '- '}{formatCurrency(item.amount)}</em>
          </div>
          <label className="form-field form-field-full">
            <span className="form-label">Explique o motivo</span>
            <textarea className="form-input decline-shared-textarea" autoFocus maxLength={240} value={reason} onChange={(event) => { setReason(event.target.value); setError(''); }} placeholder="Ex.: Valor incorreto, item não combinado..." />
          </label>
          {error ? <div className="form-error form-field-full"><AlertCircle size={15} /> {error}</div> : null}
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Voltar</button>
          <button type="submit" className="button-danger" disabled={!trimmedReason}>Recusar transação</button>
        </div>
      </form>
    </div>
  );
}

function FriendAvatar({ user }: { user: FriendUser }) {
  return user.avatarUrl ? <img src={user.avatarUrl} alt={user.name} className="friend-avatar" /> : <span className="friend-avatar friend-avatar--initials">{userInitials(user.name)}</span>;
}
function userInitials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return ((parts[0][0] ?? '') + (parts[parts.length - 1][0] ?? '')).toUpperCase();
}

function shortUserName(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length <= 1) return name.trim();
  return `${parts[0]} ${parts[parts.length - 1]}`;
}

function Topbar({ activePage, userName, userAvatarUrl, theme, notificationCount = 0, friendDetailName, onFriendDetailBack, shoppingDetailName, onShoppingDetailBack, onToggleTheme, onNavigate, onLogout, onOpenFriendSearch, onOpenShoppingCreate, onOpenGoalCreate, onGoBack, onOpenTransactionFilters, onImportTransactions, onOpenAccountFilters, accountFilterOpen, accountActiveFilterCount, onOpenCategoryFilters, categoryFilterOpen, categoryActiveFilterCount, onOpenGoalFilters: _onOpenGoalFilters, goalFilterOpen: _goalFilterOpen, goalActiveFilterCount: _goalActiveFilterCount }: {
  activePage: AppPage;
  userName: string;
  userAvatarUrl?: string | null;
  theme: 'light' | 'dark';
  notificationCount?: number;
  friendDetailName?: string | null;
  onFriendDetailBack?: () => void;
  shoppingDetailName?: string | null;
  onShoppingDetailBack?: () => void;
  onOpenFriendRequests?: () => void;
  onOpenFriendSearch?: () => void;
  onOpenShoppingCreate?: () => void;
  onOpenGoalCreate?: () => void;
  onGoBack?: () => void;
  onToggleTheme: () => void;
  onNavigate: (page: AppPage) => void;
  onLogout: () => void;
  onOpenTransactionFilters: () => void;
  onImportTransactions: () => void;
  onOpenAccountFilters: () => void;
  accountFilterOpen: boolean;
  accountActiveFilterCount: number;
  onOpenCategoryFilters: () => void;
  categoryFilterOpen: boolean;
  categoryActiveFilterCount: number;
  onOpenGoalFilters: () => void;
  goalFilterOpen: boolean;
  goalActiveFilterCount: number;
}) {
  const [actionsOpen, setActionsOpen] = useState(false);
  const [profileOpen, setProfileOpen] = useState(false);
  const actionsRef = useRef<HTMLDivElement>(null);
  const profileRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!actionsOpen) return;
    function closeOnOutside(event: PointerEvent) { if (!actionsRef.current?.contains(event.target as Node)) setActionsOpen(false); }
    function closeOnEscape(event: KeyboardEvent) { if (event.key === 'Escape') setActionsOpen(false); }
    document.addEventListener('pointerdown', closeOnOutside);
    document.addEventListener('keydown', closeOnEscape);
    return () => { document.removeEventListener('pointerdown', closeOnOutside); document.removeEventListener('keydown', closeOnEscape); };
  }, [actionsOpen]);

  useEffect(() => {
    if (!profileOpen) return;
    function closeOnOutside(event: PointerEvent) {
      const target = event.target as Element;
      if (!profileRef.current?.contains(target) && !target.closest?.('.settings-overlay'))
        setProfileOpen(false);
    }
    function closeOnEscape(event: KeyboardEvent) { if (event.key === 'Escape') setProfileOpen(false); }
    document.addEventListener('pointerdown', closeOnOutside);
    document.addEventListener('keydown', closeOnEscape);
    return () => { document.removeEventListener('pointerdown', closeOnOutside); document.removeEventListener('keydown', closeOnEscape); };
  }, [profileOpen]);

  const profileMenu = (
    <div className="profile-menu" ref={profileRef}>
      <button type="button" className={`user-profile${profileOpen ? ' active' : ''}`} onClick={() => setProfileOpen((open) => !open)} aria-expanded={profileOpen} aria-haspopup="menu" aria-label="Abrir perfil">
        <span className="user-name">{shortUserName(userName)}</span>
        {userAvatarUrl ? (
          <img src={userAvatarUrl} alt={userName} className="avatar avatar--image" />
        ) : (
          <span className="avatar avatar--initials">{userInitials(userName)}</span>
        )}
        <ChevronDown className="profile-chevron" size={15} />
      </button>
      {profileOpen ? (
        <div className="profile-dropdown" role="menu">
          <div className="profile-dropdown-user"><strong>{userName}</strong><span>Perfil do usuário</span></div>
          <button type="button" className="profile-menu-item profile-menu-item--toggle" role="menuitem" onClick={onToggleTheme}><span>{theme === 'dark' ? <Moon size={16} /> : <Sun size={16} />} Modo escuro</span><span className={`ios-switch ${theme === 'dark' ? 'ios-switch--on' : ''}`} aria-hidden="true"><span className="ios-switch-knob" /></span></button>
          <button type="button" className="profile-menu-item" role="menuitem" onClick={() => { setProfileOpen(false); onNavigate('profile'); }}><UserRound size={16} /> Meu perfil</button>
          <button type="button" className="profile-menu-item" role="menuitem" onClick={() => { setProfileOpen(false); onNavigate('notifications'); }}><Bell size={16} /> Notificações{notificationCount > 0 ? <span className="menu-notif-badge">{notificationCount}</span> : null}</button>
          <button type="button" className="profile-menu-item" role="menuitem" onClick={() => { setProfileOpen(false); onNavigate('help'); }}><HelpCircle size={16} /> Central de Ajuda</button>
          <button type="button" className="profile-logout" role="menuitem" onClick={() => { setProfileOpen(false); onLogout(); }}><LogOut size={16} /> Sair</button>
        </div>
      ) : null}
    </div>
  );

  const topbarTitle =
    activePage === 'friends' && friendDetailName ? friendDetailName :
      activePage === 'shopping' && shoppingDetailName ? shoppingDetailName :
        PAGE_LABELS[activePage];

  const isDetail = (activePage === 'friends' && friendDetailName) || (activePage === 'shopping' && shoppingDetailName);

  return (
    <header className={`topbar topbar--${activePage}${isDetail ? ' topbar--friend-detail' : ''}`}>
      <div className="topbar-desktop-shell">
        <div className="topbar-left">
          <button type="button" className="brand brand-btn" aria-label="Ir para transações" onClick={() => onNavigate('transactions')}>
            <img className="topbar-logo topbar-logo--desktop" src={theme === 'dark' ? rubyLogoWhite : rubyLogoColor} alt="RubyLife" />
            <img className="topbar-logo topbar-logo--mobile" src={rubyDiamond} alt="RubyLife" />
          </button>
          <div className="topbar-separator" />
          <nav className="breadcrumbs" aria-label="Breadcrumb"><Home className="home-icon" /><ChevronRight size={13} className="breadcrumb-chevron" /><span className="breadcrumb-active">{topbarTitle}</span></nav>
        </div>
        <div className="topbar-right">
          {activePage === 'transactions' ? (
            <div className="topbar-actions-menu" ref={actionsRef}>
              <button type="button" className="topbar-more-button" aria-label="Mais opções" aria-expanded={actionsOpen} aria-haspopup="menu" onClick={() => setActionsOpen((open) => !open)}><MoreHorizontal size={20} /></button>
              {actionsOpen ? <div className="topbar-actions-popover" role="menu"><button type="button" role="menuitem" onClick={() => { setActionsOpen(false); onOpenTransactionFilters(); }}><ListFilter size={16} />Filtros</button><button type="button" role="menuitem" onClick={() => { setActionsOpen(false); onImportTransactions(); }}><Upload size={16} />Importar</button></div> : null}
            </div>
          ) : activePage === 'accounts' ? (
            <button type="button" className={`topbar-account-filter-button${accountFilterOpen ? ' active' : ''}`} data-account-filter-trigger aria-label="Filtrar contas" aria-expanded={accountFilterOpen} aria-haspopup="dialog" onClick={onOpenAccountFilters}><ListFilter size={19} />{accountActiveFilterCount > 0 ? <span className="filter-badge">{accountActiveFilterCount}</span> : null}</button>
          ) : activePage === 'categories' ? (
            <button type="button" className={`topbar-category-filter-button${categoryFilterOpen ? ' active' : ''}`} data-category-filter-trigger aria-label="Filtrar categorias" aria-expanded={categoryFilterOpen} aria-haspopup="dialog" onClick={onOpenCategoryFilters}><ListFilter size={19} />{categoryActiveFilterCount > 0 ? <span className="filter-badge">{categoryActiveFilterCount}</span> : null}</button>
          ) : activePage === 'goals' ? (
            <button type="button" className="topbar-goal-add-button" aria-label="Nova meta" onClick={onOpenGoalCreate}><Plus size={20} /></button>
          ) : activePage === 'shopping' && shoppingDetailName ? (
            <button type="button" className="topbar-friend-back-button" aria-label="Voltar para listas" onClick={onShoppingDetailBack}><ChevronLeft size={22} /></button>
          ) : activePage === 'shopping' ? (
            <button type="button" className="topbar-shopping-add-button" aria-label="Nova lista" onClick={onOpenShoppingCreate}><Plus size={20} /></button>
          ) : activePage === 'friends' && friendDetailName ? (
            <button type="button" className="topbar-friend-back-button" aria-label="Voltar para amigos" onClick={onFriendDetailBack}><ChevronLeft size={22} /></button>
          ) : activePage === 'friends' ? (
            <button type="button" className="topbar-friend-add-button" aria-label="Adicionar amigo" onClick={onOpenFriendSearch}><UserPlus size={19} /></button>
          ) : activePage === 'notifications' || activePage === 'help' || activePage === 'profile' ? (
            <button type="button" className="topbar-friend-back-button" aria-label="Voltar" onClick={onGoBack}><ChevronLeft size={22} /></button>
          ) : null}
          {profileMenu}
        </div>
      </div>
      {profileOpen ? (
        <MobileSettingsSheet
          userName={userName}
          userAvatarUrl={userAvatarUrl}
          theme={theme}
          notificationCount={notificationCount}
          onToggleTheme={onToggleTheme}
          onNavigate={onNavigate}
          onLogout={onLogout}
          onClose={() => setProfileOpen(false)}
        />
      ) : null}
    </header>
  );
}

function MobileSettingsSheet({
  userName,
  userAvatarUrl,
  theme,
  notificationCount = 0,
  onToggleTheme,
  onNavigate,
  onLogout,
  onClose,
}: {
  userName: string;
  userAvatarUrl?: string | null;
  theme: 'light' | 'dark';
  notificationCount?: number;
  onToggleTheme: () => void;
  onNavigate: (page: AppPage) => void;
  onLogout: () => void;
  onClose: () => void;
}) {
  return createPortal(
    <div className="settings-overlay" role="dialog" aria-modal="true" aria-label="Ajustes" onClick={onClose}>
      <div className="settings-screen" onClick={(event) => event.stopPropagation()}>
        <div className="settings-screen-head">
          <button type="button" className="settings-close" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <button type="button" className="settings-profile-card" onClick={() => { onClose(); onNavigate('profile'); }}>
          {userAvatarUrl ? (
            <img src={userAvatarUrl} alt={userName} className="settings-avatar" />
          ) : (
            <span className="settings-avatar settings-avatar--initials">{userInitials(userName)}</span>
          )}
          <span className="settings-profile-copy"><strong>{userName}</strong><span>Ver perfil</span></span>
          <ChevronRight className="settings-chevron" size={18} />
        </button>
        <div className="settings-group">
          <button type="button" className="settings-row settings-row--toggle" onClick={onToggleTheme}>
            <span className="settings-row-icon settings-row-icon--purple">{theme === 'dark' ? <Moon size={16} /> : <Sun size={16} />}</span>
            <span className="settings-row-label">Modo escuro</span>
            <span className={`ios-switch ${theme === 'dark' ? 'ios-switch--on' : ''}`} aria-hidden="true"><span className="ios-switch-knob" /></span>
          </button>
        </div>
        <div className="settings-group">
          <button type="button" className="settings-row" onClick={() => { onClose(); onNavigate('categories'); }}>
            <span className="settings-row-icon settings-row-icon--orange"><Tags size={16} /></span>
            <span className="settings-row-label">Categorias</span>
            <ChevronRight className="settings-chevron" size={16} />
          </button>
          <button type="button" className="settings-row" onClick={() => { onClose(); onNavigate('friends'); }}>
            <span className="settings-row-icon settings-row-icon--blue"><Users size={16} /></span>
            <span className="settings-row-label">Amigos</span>
            <ChevronRight className="settings-chevron" size={16} />
          </button>
          <button type="button" className="settings-row" onClick={() => { onClose(); onNavigate('notifications'); }}>
            <span className="settings-row-icon settings-row-icon--red"><Bell size={16} /></span>
            <span className="settings-row-label">Notificações{notificationCount > 0 ? <span className="settings-notif-badge">{notificationCount}</span> : null}</span>
            <ChevronRight className="settings-chevron" size={16} />
          </button>
          <button type="button" className="settings-row" onClick={() => { onClose(); onNavigate('help'); }}>
            <span className="settings-row-icon settings-row-icon--gray"><HelpCircle size={16} /></span>
            <span className="settings-row-label">Central de Ajuda</span>
            <ChevronRight className="settings-chevron" size={16} />
          </button>
        </div>
        <div className="settings-group">
          <button type="button" className="settings-row settings-row--logout" onClick={() => { onClose(); onLogout(); }}>
            <span className="settings-row-icon settings-row-icon--red"><LogOut size={16} /></span>
            <span className="settings-row-label">Sair</span>
          </button>
        </div>
      </div>
    </div>,
    document.body
  );
}
function Sidebar({ activePage, onNavigate }: { activePage: AppPage; onNavigate: (page: AppPage) => void }) {
  const items = [
    { label: 'Visão Geral', icon: LayoutGrid, page: 'dashboard' as const },
    { label: 'Transações', icon: ReceiptText, page: 'transactions' as const },
    { label: 'Categorias', icon: Tags, page: 'categories' as const },
    { label: 'Metas', icon: Target, page: 'goals' as const },
    { label: 'Contas', icon: CreditCard, page: 'accounts' as const },
    { label: 'Listas', icon: ShoppingCart, page: 'shopping' as const },
    { label: 'Amigos', icon: Users, page: 'friends' as const },
  ];

  return (
    <aside className="sidebar sidebar-desktop">
      <div className="sidebar-scroll sidebar-root-nav">
        {items.map((item) => {
          const Icon = item.icon;
          const active = item.page === activePage;
          return (
            <div key={item.label} className={`sidebar-root-group ${active ? 'active' : ''}`}>
              <button type="button" className={`sidebar-dashboard sidebar-root-item ${active ? 'active' : ''}`} aria-label={item.label} onClick={(event) => { event.currentTarget.blur(); onNavigate(item.page); }}><span className="sidebar-dashboard-mark"><Icon size={18} /></span><span className="sidebar-dashboard-label">{item.label}</span></button>
              <span className="sidebar-root-tooltip" role="tooltip">{item.label}</span>
            </div>
          );
        })}
      </div>
    </aside>
  );
}

function MobileTabBar({
  activePage,
  onNavigate,
  onNewIncome,
  onNewExpense,
  onNewGoal,
  onNewCategory,
  onNewSharedTransaction: _onNewSharedTransaction,
  onNewShoppingList: _onNewShoppingList,
}: {
  activePage: AppPage;
  onNavigate: (page: AppPage) => void;
  onNewIncome: () => void;
  onNewExpense: () => void;
  onNewGoal: () => void;
  onNewCategory: () => void;
  onNewSharedTransaction?: () => void;
  onNewShoppingList?: () => void;
}) {
  const [actionOpen, setActionOpen] = useState(false);

  const tabs = [
    { label: 'Início', icon: Home, page: 'dashboard' as const },
    { label: 'Transações', icon: ReceiptText, page: 'transactions' as const },
    { label: 'Metas', icon: Target, page: 'goals' as const },
    { label: 'Compras', icon: ShoppingCart, page: 'shopping' as const },
  ];

  const renderTab = (item: (typeof tabs)[number]) => {
    const Icon = item.icon;
    const active = item.page === activePage;
    return (
      <button
        key={item.label}
        type="button"
        className={`mobile-tab-item ${active ? 'mobile-tab-item--active active' : ''}`}
        onClick={() => { setActionOpen(false); onNavigate(item.page); }}
        aria-label={item.label}
      >
        <Icon className="mobile-tab-icon" size={20} />
        <span className="mobile-tab-label">{item.label}</span>
      </button>
    );
  };

  return (
    <>
      <nav className="mobile-tab-bar" aria-label="Navegação inferior móvel">
        {tabs.slice(0, 2).map(renderTab)}

        <button
          type="button"
          className={`mobile-tab-fab ${actionOpen ? 'mobile-tab-fab--open' : ''}`}
          onClick={() => setActionOpen((open) => !open)}
          aria-expanded={actionOpen}
          aria-label={actionOpen ? 'Fechar menu de novo registro' : 'Novo registro'}
        >
          {actionOpen ? <X size={26} /> : <Plus size={26} />}
        </button>

        {tabs.slice(2).map(renderTab)}
      </nav>

      {actionOpen && (
        <div className="quick-add-overlay" onClick={() => setActionOpen(false)}>
          <div className="radial-menu-container" role="dialog" aria-modal="true" aria-label="Novo registro" onClick={(e) => e.stopPropagation()}>

            <button type="button" className="radial-menu-item radial-menu-item--meta" onClick={() => { setActionOpen(false); onNewGoal(); }}>
              <span className="radial-menu-icon radial-menu-icon--meta"><Target size={24} /></span>
              <span className="radial-menu-label">Meta</span>
            </button>

            <button type="button" className="radial-menu-item radial-menu-item--expense" onClick={() => { setActionOpen(false); onNewExpense(); }}>
              <span className="radial-menu-icon radial-menu-icon--expense"><ArrowUpRight size={24} /></span>
              <span className="radial-menu-label">Despesa</span>
            </button>

            <button type="button" className="radial-menu-item radial-menu-item--income" onClick={() => { setActionOpen(false); onNewIncome(); }}>
              <span className="radial-menu-icon radial-menu-icon--income"><ArrowDownLeft size={24} /></span>
              <span className="radial-menu-label">Receita</span>
            </button>

            <button type="button" className="radial-menu-item radial-menu-item--category" onClick={() => { setActionOpen(false); onNewCategory(); }}>
              <span className="radial-menu-icon radial-menu-icon--category"><Tags size={24} /></span>
              <span className="radial-menu-label">Categoria</span>
            </button>

            <button type="button" className="radial-menu-close" onClick={() => setActionOpen(false)} aria-label="Fechar menu">
              <X size={26} />
            </button>
          </div>
        </div>
      )}
    </>
  );
}

function CategoriesPage({ items, filters, draftFilters, filterOpen, filterControlRef, onOpenFilters, onDraftFiltersChange, onApplyFilters, onNew, onEdit, onDelete, onBulkDelete }: {
  items: CategoryItem[];
  filters: CategoryFilters;
  draftFilters: CategoryFilters;
  filterOpen: boolean;
  filterControlRef: RefObject<HTMLDivElement | null>;
  onOpenFilters: () => void;
  onDraftFiltersChange: (filters: CategoryFilters) => void;
  onApplyFilters: () => void;
  onNew: () => void;
  onEdit: (category: CategoryItem) => void;
  onDelete: (category: CategoryItem) => void;
  onBulkDelete: (ids: string[]) => void;
}) {
  const [selected, setSelected] = useState<Set<string>>(() => new Set());

  const activeFilterCount = Number(filters.type !== 'all') + Number(Boolean(filters.search.trim()));
  const filteredItems = items.filter((item) => (
    (filters.type === 'all' || item.kind === filters.type)
    && item.name.toLowerCase().includes(filters.search.toLowerCase().trim())
  ));
  const selectedInView = filteredItems.filter((item) => selected.has(item.id));
  const allSelected = filteredItems.length > 0 && selectedInView.length === filteredItems.length;
  const incomeItems = filteredItems.filter((item) => item.kind === 'income');
  const expenseItems = filteredItems.filter((item) => item.kind === 'expense');
  const transferItems = filteredItems.filter((item) => item.kind === 'transfer');

  function toggle(id: string) {
    setSelected((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleAll() {
    setSelected((current) => {
      if (filteredItems.length > 0 && filteredItems.every((item) => current.has(item.id))) {
        const next = new Set(current);
        filteredItems.forEach((item) => next.delete(item.id));
        return next;
      }
      const next = new Set(current);
      filteredItems.forEach((item) => next.add(item.id));
      return next;
    });
  }

  function toggleGroup(groupItems: CategoryItem[]) {
    setSelected((current) => {
      if (groupItems.length > 0 && groupItems.every((item) => current.has(item.id))) {
        const next = new Set(current);
        groupItems.forEach((item) => next.delete(item.id));
        return next;
      }
      const next = new Set(current);
      groupItems.forEach((item) => next.add(item.id));
      return next;
    });
  }

  const renderTableColumn = (label: string, kind: CategoryKind, groupItems: CategoryItem[]) => {
    const groupSelected = groupItems.length > 0 && groupItems.every((item) => selected.has(item.id));
    return (
      <div className={`category-table-column category-table-column--${kind}`}>
        <div className={`category-table-header category-table-header--${kind}`}>
          <span className="category-table-title">{label}</span>
          <span className="category-table-count">{groupItems.length}</span>
        </div>
        <div className="category-table-grid category-table-head">
          <span className="row-check-col"><input type="checkbox" className="row-check" checked={groupSelected} onChange={() => toggleGroup(groupItems)} aria-label={`Selecionar todas as ${label}`} /></span>
          <span>Categoria</span>
          <span style={{ textAlign: 'center' }}>Cor</span>
          <span style={{ textAlign: 'center' }}>Ações</span>
        </div>
        <div className="category-table-body">
          {groupItems.length > 0 ? groupItems.map((category) => (
            <article key={category.id} className={`category-table-grid category-table-row${selected.has(category.id) ? ' is-selected' : ''}`}>
              <span className="row-check-col"><input type="checkbox" className="row-check" checked={selected.has(category.id)} onChange={() => toggle(category.id)} aria-label={`Selecionar ${category.name}`} /></span>
              <span className="category-report-name category-desktop-name">
                <span className="category-desktop-icon" style={{ backgroundColor: `${category.color}18`, color: category.color }} aria-hidden="true">
                  <CategoryIconGraphic icon={category.icon} size={16} />
                </span>
                <strong className="category-desktop-label">{category.name}</strong>
              </span>
              <span className="category-color-cell" title={`Cor: ${category.color}`}><i className="category-color-only-dot" style={{ backgroundColor: category.color }} /></span>
              <span className="launch-actions-cell category-icons-actions">
                <button type="button" className="action-icon-btn" onClick={() => onEdit(category)} title="Editar"><Pencil size={15} /></button>
                <button type="button" className="action-icon-btn action-icon-btn--danger" onClick={() => onDelete(category)} title="Excluir"><Trash2 size={15} /></button>
              </span>
            </article>
          )) : (
            <div className="category-table-empty"><span>Nenhuma categoria</span></div>
          )}
        </div>
      </div>
    );
  };

  return (
    <>
      <section className="page-header page-header-split">
        <div className="page-header-left">
          <h1 className="page-title">Categorias</h1>
        </div>
        <div className="page-header-center"></div>
        <div className="page-header-actions">
          <div className="filter-control category-page-filter-control" ref={filterControlRef}>
            <button type="button" className={`filter-trigger btn-icon-only${filterOpen ? ' active' : ''}`} data-category-filter-trigger onClick={onOpenFilters} aria-expanded={filterOpen} aria-haspopup="dialog" title="Filtros">
              <ListFilter size={16} />
              {activeFilterCount ? <span className="filter-badge">{activeFilterCount}</span> : null}
            </button>
            {filterOpen ? (
              <div className="filter-popover category-filter-popover" role="dialog" aria-label="Filtros das categorias">
                <div className="filter-popover-header"><strong>Filtrar categorias</strong><button type="button" className="filter-popover-close" onClick={onOpenFilters} aria-label="Fechar filtros"><X size={18} /></button></div>
                <div className="filter-grid filter-grid--stacked category-filter-grid">
                  <label className="filter-field">
                    <span>Categoria</span>
                    <input value={draftFilters.search} onChange={(event) => onDraftFiltersChange({ ...draftFilters, search: event.target.value })} placeholder="Buscar pelo nome" />
                  </label>
                  <label className="filter-field">
                    <span>Tipo</span>
                    <select value={draftFilters.type} onChange={(event) => onDraftFiltersChange({ ...draftFilters, type: event.target.value as 'all' | CategoryKind })}>
                      <option value="all">Todos</option>
                      <option value="income">Receita</option>
                      <option value="expense">Despesa</option>
                      <option value="transfer">Transferência</option>
                    </select>
                  </label>
                </div>
                <div className="filter-popover-actions">
                  <button type="button" className="filter-clear" onClick={() => onDraftFiltersChange({ search: '', type: 'all' })}><RotateCcw size={14} /> Limpar</button>
                  <button type="button" className="filter-apply" onClick={onApplyFilters}>Aplicar filtros</button>
                </div>
              </div>
            ) : null}
          </div>
          <button type="button" className="page-primary-action" onClick={onNew}>
            <Plus size={16} /> Nova categoria
          </button>
        </div>
      </section>

      {selectedInView.length > 0 ? (
        <div className="bulk-bar">
          <span className="bulk-bar-count"><strong>{selectedInView.length}</strong> {selectedInView.length === 1 ? 'selecionada' : 'selecionadas'}</span>
          <div className="bulk-bar-actions">
            <button type="button" className="bulk-bar-button" onClick={() => setSelected(new Set())}>Limpar seleção</button>
            <button type="button" className="bulk-bar-button bulk-bar-button--danger" onClick={() => onBulkDelete(selectedInView.map((category) => category.id))}><Trash2 size={15} /> Excluir selecionadas</button>
          </div>
        </div>
      ) : null}

      <section className="resource-panel">
        <div className="categories-desktop-view">
          {renderTableColumn('Receita', 'income', incomeItems)}
          {renderTableColumn('Despesa', 'expense', expenseItems)}
          {renderTableColumn('Transferência', 'transfer', transferItems)}
        </div>

        <div className="categories-mobile-view">
          <div className="documents-report categories-report">
            <div className="documents-report-head categories-report-grid">
              <span className="row-check-col"><input type="checkbox" className="row-check" checked={allSelected} onChange={toggleAll} aria-label="Selecionar todas" /></span>
              <span>Categoria</span>
              <span>Tipo</span>
              <span>Cor</span>
              <span>Ações</span>
            </div>
            <div className="documents-report-body">
              {filteredItems.length ? filteredItems.map((category) => (
                <article key={category.id} className={`documents-report-row categories-report-grid${selected.has(category.id) ? ' is-selected' : ''}`}>
                  <span className="row-check-col"><input type="checkbox" className="row-check" checked={selected.has(category.id)} onChange={() => toggle(category.id)} aria-label={`Selecionar ${category.name}`} /></span>
                  <span className="category-report-name"><i className="category-name-dot" style={{ backgroundColor: category.color }} aria-hidden="true" />{category.name}</span>
                  <span className={`category-type category-type--${category.kind}`}>{CATEGORY_KIND_LABELS[category.kind]}</span>
                  <span className="category-color-value" aria-label={`Cor ${category.color}`}><i style={{ backgroundColor: category.color }} /> <em className="category-color-hex">{category.color.toUpperCase()}</em></span>
                  <span className="launch-actions-cell">
                    <RowActions actions={[
                      { key: 'edit', label: 'Editar', icon: <Pencil size={15} />, onClick: () => onEdit(category) },
                      { key: 'delete', label: 'Excluir', icon: <Trash2 size={15} />, onClick: () => onDelete(category), danger: true },
                    ]} />
                    <div className="category-card-actions">
                      <button type="button" className="category-card-action" onClick={() => onEdit(category)}><Pencil size={15} /> Editar</button>
                      <button type="button" className="category-card-action category-card-action--danger" onClick={() => onDelete(category)}><Trash2 size={15} /> Excluir</button>
                    </div>
                  </span>
                </article>
              )) : (
                <div className="resource-empty-state compact">
                  <Tags size={28} />
                  <strong>Nenhuma categoria encontrada</strong>
                </div>
              )}
            </div>
          </div>
        </div>
      </section>
    </>
  );
}

const ALL_CATEGORY_ICONS: CategoryIconKey[] = [
  'salary', 'money', 'work', 'shopping', 'refund', 'investment',
  'home', 'card', 'gift', 'food', 'car', 'health', 'education',
  'leisure', 'repeat', 'family', 'pets', 'donation', 'wallet', 'transfer', 'tag',
  'travel', 'fitness', 'music', 'maintenance'
];

function CategoryModal({ items, category, onClose, onSave }: { items: CategoryItem[]; category: CategoryItem | null; onClose: () => void; onSave: (category: CustomCategory) => void }) {
  const [name, setName] = useState(category?.name ?? '');
  const [kind, setKind] = useState<CategoryKind>(category?.kind ?? 'expense');
  const [color, setColor] = useState(category?.color ?? ACCOUNT_COLORS[0] ?? '#1B99D8');
  const [selectedIcon, setSelectedIcon] = useState<CategoryIconKey>(category?.icon ?? 'tag');
  const [hasManuallySelectedIcon, setHasManuallySelectedIcon] = useState(false);
  const [iconPickerOpen, setIconPickerOpen] = useState(false);
  const [colorPickerOpen, setColorPickerOpen] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    if (!category && !hasManuallySelectedIcon) {
      setSelectedIcon(defaultCategoryIcon(name, kind));
    }
  }, [name, kind, category, hasManuallySelectedIcon]);

  useEffect(() => {
    if (!iconPickerOpen) return;
    function handleEscape(event: KeyboardEvent) {
      if (event.key === 'Escape') setIconPickerOpen(false);
    }
    document.addEventListener('keydown', handleEscape);
    return () => document.removeEventListener('keydown', handleEscape);
  }, [iconPickerOpen]);

  function submit(event: FormEvent) {
    event.preventDefault();
    const normalizedName = name.trim();
    if (!normalizedName) return;
    if (items.some((item) => item.id !== category?.id && item.kind === kind && item.name.toLowerCase() === normalizedName.toLowerCase())) {
      setError('Essa categoria já existe nesse grupo.');
      return;
    }
    onSave({ id: category?.id ?? crypto.randomUUID(), name: normalizedName, kind, color, icon: selectedIcon });
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card settle-modal" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <img src={rubyDiamond} alt="" className="modal-logo" />
          <div className="modal-title-container"><h2 className="modal-title">{category ? 'Editar categoria' : 'Nova categoria'}</h2></div>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body">
          <div className="modal-form-grid">
            <label className="form-field form-field-full">
              <span className="form-label">Nome da categoria</span>
              <input className="form-input" value={name} onChange={(event) => { setName(event.target.value); setError(''); }} placeholder="Ex.: Academia" />
            </label>
            <label className="form-field form-field-full">
              <span className="form-label">Tipo</span>
              <select className="form-input" value={kind} onChange={(event) => { setKind(event.target.value as CategoryKind); setError(''); }}>
                <option value="income">Receita</option>
                <option value="expense">Despesa</option>
                <option value="transfer">Transferência</option>
              </select>
            </label>

            <div className="category-selector-grid">
              <div className="form-field">
                <span className="form-label">Ícone</span>
                <button
                  type="button"
                  className="form-input category-selector-btn"
                  aria-haspopup="dialog"
                  aria-expanded={iconPickerOpen}
                  onClick={() => setIconPickerOpen(true)}
                >
                  <span className="category-selector-icon" style={{ backgroundColor: `${color}18`, color }}>
                    <CategoryIconGraphic icon={selectedIcon} size={16} />
                  </span>
                  <span className="category-selector-label">{selectedIcon}</span>
                </button>
              </div>

              <div className="form-field">
                <span className="form-label">Cor</span>
                <button
                  type="button"
                  className="form-input category-selector-btn"
                  aria-haspopup="dialog"
                  aria-expanded={colorPickerOpen}
                  onClick={() => setColorPickerOpen(true)}
                >
                  <span className="category-selector-color" style={{ backgroundColor: color }} />
                  <span className="category-selector-label">Escolher cor</span>
                </button>
              </div>
            </div>
            {error ? <div className="form-error form-field-full"><AlertCircle size={15} /> {error}</div> : null}
          </div>
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary">{category ? 'Salvar alterações' : 'Criar categoria'}</button>
        </div>
      </form>
      {colorPickerOpen ? <ColorSpectrumSheet value={color} onChange={setColor} onClose={() => setColorPickerOpen(false)} /> : null}
      {iconPickerOpen
        ? createPortal(
          <div className="category-icon-picker-layer" onClick={() => setIconPickerOpen(false)}>
            <div className="category-icon-picker-dialog" role="dialog" aria-modal="true" aria-label="Escolher ícone" onClick={(event) => event.stopPropagation()}>
              <div className="category-icon-picker-head">
                <strong>Escolher ícone</strong>
                <button type="button" onClick={() => setIconPickerOpen(false)} aria-label="Fechar seletor de ícones"><X size={18} /></button>
              </div>
              <div className="category-icon-picker-grid">
                {ALL_CATEGORY_ICONS.map((key) => (
                  <button
                    key={key}
                    type="button"
                    className={selectedIcon === key ? 'active' : ''}
                    style={{ color }}
                    aria-label={`Selecionar ícone ${key}`}
                    aria-pressed={selectedIcon === key}
                    onClick={() => { setSelectedIcon(key); setHasManuallySelectedIcon(true); setIconPickerOpen(false); }}
                  >
                    <CategoryIconGraphic icon={key} size={20} />
                  </button>
                ))}
              </div>
            </div>
          </div>,
          document.body,
        )
        : null}
    </div>
  );
}

function DeleteCategoryModal({ category, onClose, onConfirm }: { category: CategoryItem; onClose: () => void; onConfirm: () => void }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card confirm-modal" role="dialog" aria-modal="true" aria-labelledby="delete-category-title" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only" id="delete-category-title">Excluir categoria</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body confirmation-message">
          <span className="confirmation-icon"><Trash2 size={20} /></span>
          <p>Deseja excluir a categoria <strong>{category.name}</strong>?</p>
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="button" className="button-danger" onClick={onConfirm}><Trash2 size={15} /> Excluir</button>
        </div>
      </div>
    </div>
  );
}

function DeleteTransactionModal({ item, relatedCount, onClose, onConfirmOne, onConfirmAll }: { item: Transaction; relatedCount: number; onClose: () => void; onConfirmOne: () => void; onConfirmAll: () => void }) {
  const canDeleteSeries = item.recurrence !== 'single' && relatedCount > 1;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card confirm-modal" role="dialog" aria-modal="true" aria-labelledby="delete-transaction-title" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only" id="delete-transaction-title">Excluir transação</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body confirmation-message" style={{ flexDirection: 'column', alignItems: 'flex-start', gap: '16px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <span className="confirmation-icon"><Trash2 size={20} /></span>
            <p style={{ margin: 0 }}>Deseja excluir o lançamento <strong>{displayTransactionDescription(item.description)}</strong>?</p>
          </div>
          {canDeleteSeries ? (
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: '10px', padding: '12px 14px', backgroundColor: '#f8fafc', borderRadius: '10px', border: '1px solid #e2e8f0', width: '100%', fontSize: '13px', color: '#475569', lineHeight: 1.4 }}>
              <span style={{ color: '#0ea5e9', fontSize: '16px' }}>ℹ️</span>
              <div>Esta transação faz parte de uma recorrência (despesa fixa/parcelada) com <strong>{relatedCount} registros</strong>. Você vai escolher só essa ou todas as sessões/transações? Atenção: ao escolher <strong>Todas as transações</strong>, será excluído tudo (tanto as anteriores quanto as futuras).</div>
            </div>
          ) : null}
        </div>
        <div className="modal-actions" style={{ flexWrap: 'wrap', gap: '8px' }}>
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="button" className="button-danger" onClick={onConfirmOne}><Trash2 size={15} /> {canDeleteSeries ? 'Excluir só essa' : 'Excluir'}</button>
          {canDeleteSeries ? <button type="button" className="button-danger" style={{ backgroundColor: '#7f1d1d' }} onClick={onConfirmAll}><Trash2 size={15} /> Excluir todas as transações ({relatedCount})</button> : null}
        </div>
      </div>
    </div>
  );
}
function EditTransactionModal({ item, accounts, incomeCategories, expenseCategories, onClose, onSave }: { item: Transaction; accounts: Account[]; incomeCategories: string[]; expenseCategories: string[]; onClose: () => void; onSave: (updated: Transaction) => void }) {
  const [type, setType] = useState<TransactionType>(item.type);
  const [description, setDescription] = useState(item.description);
  const [category, setCategory] = useState(item.category);
  const [amount, setAmount] = useState(() => formatCurrencyInput(item.amount));
  const [dueDate, setDueDate] = useState(item.dueDate);
  const [recurrence, setRecurrence] = useState<RecurrenceType>(item.recurrence);
  const [accountId, setAccountId] = useState(item.accountId ?? '');
  const [error, setError] = useState('');

  const categories = type === 'income' ? incomeCategories : expenseCategories;

  function selectType(next: TransactionType) {
    setType(next);
    const list = next === 'income' ? incomeCategories : expenseCategories;
    if (!list.includes(category)) setCategory(list[0] ?? '');
  }

  function submit(event: FormEvent) {
    event.preventDefault();
    const normalizedDescription = description.trim().slice(0, TRANSACTION_DESCRIPTION_MAX_LENGTH);
    const value = parseAmount(amount);
    if (!normalizedDescription) { setError('Informe uma descrição.'); return; }
    if (value <= 0) { setError('Informe um valor válido.'); return; }
    const account = accounts.find((candidate) => candidate.id === accountId);
    onSave({ ...item, type, description: normalizedDescription, category, amount: value, dueDate, recurrence, accountId: account?.id, account: account?.name });
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">Editar transação</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body">
          <div className="modal-form-grid">
            <div className="form-field form-field-full">
              <span className="form-label">Tipo</span>
              <div className="transaction-type-options" role="group" aria-label="Tipo da transação">
                <button type="button" className={`transaction-type-option transaction-type-option--expense${type === 'expense' ? ' active' : ''}`} aria-pressed={type === 'expense'} onClick={() => selectType('expense')}>
                  <span><ArrowDownLeft size={20} /></span>
                  <strong>Despesa</strong>
                </button>
                <button type="button" className={`transaction-type-option transaction-type-option--income${type === 'income' ? ' active' : ''}`} aria-pressed={type === 'income'} onClick={() => selectType('income')}>
                  <span><ArrowUpRight size={20} /></span>
                  <strong>Receita</strong>
                </button>
              </div>
            </div>
            <label className="form-field form-field-full">
              <span className="form-label">Descrição</span>
              <input className="form-input" value={description} maxLength={TRANSACTION_DESCRIPTION_MAX_LENGTH} onChange={(event) => { setDescription(event.target.value); setError(''); }} placeholder="Ex.: Aluguel, internet, venda..." />
            </label>
            <label className="form-field">
              <span className="form-label">Valor</span>
              <div className="form-input-wrap">
                <span className="form-input-prefix">R$</span>
                <input className="form-input" value={amount} onChange={(event) => { setAmount(formatCurrencyInput(event.target.value)); setError(''); }} placeholder="0,00" inputMode="decimal" />
              </div>
            </label>
            <label className="form-field">
              <span className="form-label">Vencimento</span>
              <div className="form-input-wrap">
                <CalendarDays size={16} />
                <input className="form-input" type="date" value={dueDate} onChange={(event) => setDueDate(event.target.value)} />
              </div>
            </label>
            <label className="form-field">
              <span className="form-label">Categoria</span>
              <select className="form-input" value={category} onChange={(event) => setCategory(event.target.value)}>
                {categories.map((option) => <option key={option}>{option}</option>)}
              </select>
            </label>
            <label className="form-field">
              <span className="form-label">Recorrência</span>
              <select className="form-input" value={recurrence} onChange={(event) => setRecurrence(event.target.value as RecurrenceType)}>
                <option value="single">Única</option>
                <option value="fixed">Fixa</option>
                <option value="installment">Parcelada</option>
              </select>
            </label>
            <label className="form-field form-field-full">
              <span className="form-label form-label-optional">Banco / conta</span>
              <select className="form-input" value={accountId} onChange={(event) => setAccountId(event.target.value)}>
                <option value="">Nenhum banco ou conta selecionado</option>
                {accounts.map((account) => <option key={account.id} value={account.id}>{account.name}</option>)}
              </select>
            </label>
            {error ? <div className="form-error form-field-full"><AlertCircle size={15} /> {error}</div> : null}
          </div>
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary">Salvar alterações</button>
        </div>
      </form>
    </div>
  );
}

function AccountsPage({ accounts, balances, filters, draftFilters, filterOpen, filterControlRef, onOpenFilters, onDraftFiltersChange, onApplyFilters, onNew, onEdit, onDelete, onBulkDelete }: {
  accounts: Account[];
  balances: Record<string, number>;
  filters: AccountFilters;
  draftFilters: AccountFilters;
  filterOpen: boolean;
  filterControlRef: RefObject<HTMLDivElement | null>;
  onOpenFilters: () => void;
  onDraftFiltersChange: (filters: AccountFilters) => void;
  onApplyFilters: () => void;
  onNew: () => void;
  onEdit: (account: Account) => void;
  onDelete: (account: Account) => void;
  onBulkDelete: (ids: string[]) => void;
}) {
  const [selected, setSelected] = useState<Set<string>>(() => new Set());
  const normalizedSearch = filters.search.trim().toLowerCase();
  const filteredAccounts = accounts.filter((account) => (
    (filters.type === 'all' || account.type === filters.type)
    && account.name.toLowerCase().includes(normalizedSearch)
  ));
  const activeFilterCount = Number(Boolean(filters.search.trim())) + Number(filters.type !== 'all');
  const deletable = filteredAccounts;
  const selectedInView = deletable.filter((account) => selected.has(account.id));
  const allSelected = deletable.length > 0 && selectedInView.length === deletable.length;

  function toggle(id: string) {
    setSelected((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  function toggleAll() {
    setSelected((current) => {
      if (deletable.length > 0 && deletable.every((account) => current.has(account.id))) return new Set();
      return new Set(deletable.map((account) => account.id));
    });
  }

  return (
    <>
      <section className="page-header page-header-split">
        <div className="page-header-left">
          <h1 className="page-title">Contas</h1>
        </div>
        <div className="page-header-center"></div>
        <div className="page-header-actions">
          <div className="filter-control account-filter-control" ref={filterControlRef}>
            <button
              type="button"
              className={`filter-trigger btn-icon-only${filterOpen ? ' active' : ''}`}
              data-account-filter-trigger
              onClick={onOpenFilters}
              aria-expanded={filterOpen}
              aria-haspopup="dialog"
              title="Filtros"
            >
              <ListFilter size={16} />
              {activeFilterCount > 0 ? <span className="filter-badge">{activeFilterCount}</span> : null}
            </button>
            {filterOpen ? (
              <div className="filter-popover account-filter-popover" role="dialog" aria-label="Filtros das contas">
                <div className="filter-popover-header"><strong>Filtrar contas</strong><button type="button" className="filter-popover-close" onClick={onOpenFilters} aria-label="Fechar filtros"><X size={18} /></button></div>
                <div className="filter-grid filter-grid--stacked account-filter-grid">
                  <label className="filter-field">
                    <span>Conta</span>
                    <input
                      value={draftFilters.search}
                      onChange={(event) => onDraftFiltersChange({ ...draftFilters, search: event.target.value })}
                      placeholder="Buscar pelo nome"
                    />
                  </label>
                  <label className="filter-field">
                    <span>Tipo</span>
                    <select
                      value={draftFilters.type}
                      onChange={(event) => onDraftFiltersChange({ ...draftFilters, type: event.target.value as AccountTypeFilter })}
                    >
                      <option value="all">Todos</option>
                      {Object.entries(ACCOUNT_TYPE_LABELS).map(([value, label]) => <option key={value} value={value}>{label}</option>)}
                    </select>
                  </label>
                </div>
                <div className="filter-popover-actions">
                  <button type="button" className="filter-clear" onClick={() => onDraftFiltersChange({ search: '', type: 'all' })}><RotateCcw size={14} /> Limpar</button>
                  <button type="button" className="filter-apply" onClick={onApplyFilters}>Aplicar filtros</button>
                </div>
              </div>
            ) : null}
          </div>
          <button type="button" className="page-primary-action" onClick={onNew}>
            <Plus size={16} /> Nova conta
          </button>
        </div>
      </section>

      {selectedInView.length > 0 ? (
        <div className="bulk-bar">
          <span className="bulk-bar-count"><strong>{selectedInView.length}</strong> {selectedInView.length === 1 ? 'selecionada' : 'selecionadas'}</span>
          <div className="bulk-bar-actions">
            <button type="button" className="bulk-bar-button" onClick={() => setSelected(new Set())}>Limpar seleção</button>
            <button type="button" className="bulk-bar-button bulk-bar-button--danger" onClick={() => onBulkDelete(selectedInView.map((account) => account.id))}><Trash2 size={15} /> Excluir selecionadas</button>
          </div>
        </div>
      ) : null}

      <section className="resource-panel">
        <div className="documents-report accounts-report">
          <div className="documents-report-head accounts-report-grid">
            <span className="row-check-col"><input type="checkbox" className="row-check" checked={allSelected} onChange={toggleAll} aria-label="Selecionar todas" /></span>
            <span>Conta</span>
            <span>Tipo</span>
            <span className="report-col-value">Saldo inicial</span>
            <span className="report-col-value">Saldo atual</span>
            <span>Ações</span>
          </div>
          <div className="documents-report-body">
            {filteredAccounts.length === 0 ? (
              <div className="resource-empty-state">
                <ListFilter size={30} />
                <strong>Nenhuma conta encontrada</strong>
                <p>Ajuste os filtros para visualizar outras contas.</p>
              </div>
            ) : filteredAccounts.map((account) => {
              const balance = balances[account.id] ?? account.initialBalance;
              const color = account.color ?? '#1B99D8';
              return (
                <article key={account.id} className={`documents-report-row accounts-report-grid${selected.has(account.id) ? ' is-selected' : ''}`}>
                  <span className="row-check-col">
                    <input type="checkbox" className="row-check" checked={selected.has(account.id)} onChange={() => toggle(account.id)} aria-label={`Selecionar ${account.name}`} />
                  </span>
                  <div className="account-main">
                    <span className="account-color-dot" style={{ backgroundColor: color }} aria-hidden="true" />
                    <span className="account-icon" style={{ backgroundColor: `${color}18`, color }}>
                      <AccountIconGraphic icon={accountIconFor(account)} />
                    </span>
                    <strong>{account.name}</strong>
                  </div>
                  <span className="launch-muted account-type-cell">{ACCOUNT_TYPE_LABELS[account.type]}</span>
                  <span className="account-balance account-balance--muted account-balance--initial">
                    <span className="account-balance-head">
                      <span className={`account-balance-arrow ${account.initialBalance < 0 ? 'account-balance-arrow--down' : 'account-balance-arrow--up'}`}>
                        {account.initialBalance < 0 ? <ArrowDownLeft size={13} /> : <ArrowUpRight size={13} />}
                      </span>
                      <span className="mobile-balance-label">Saldo inicial</span>
                    </span>
                    <span>{formatCurrency(account.initialBalance)}</span>
                  </span>
                  <span className="account-balance account-balance--current">
                    <span className="account-balance-head">
                      <span className={`account-balance-arrow ${balance < 0 ? 'account-balance-arrow--down' : 'account-balance-arrow--up'}`}>
                        {balance < 0 ? <ArrowDownLeft size={13} /> : <ArrowUpRight size={13} />}
                      </span>
                      <span className="mobile-balance-label">Saldo atual</span>
                    </span>
                    <strong className={balance < 0 ? 'negative' : ''}>{formatCurrency(balance)}</strong>
                  </span>
                  <span className="launch-actions-cell">
                    <RowActions actions={[
                      { key: 'edit', label: 'Editar', icon: <Pencil size={15} />, onClick: () => onEdit(account) },
                      { key: 'delete', label: 'Excluir', icon: <Trash2 size={15} />, onClick: () => onDelete(account), danger: true },
                    ]} />
                    <div className="account-card-actions">
                      <button type="button" className="account-card-action" onClick={() => onEdit(account)}><Pencil size={15} /> Editar</button>
                      <button type="button" className="account-card-action account-card-action--danger" onClick={() => onDelete(account)}><Trash2 size={15} /> Excluir</button>
                    </div>
                  </span>
                </article>
              );
            })}
          </div>
        </div>
      </section>
    </>
  );
}

function AccountModal({ accounts, account, onClose, onSave }: { accounts: Account[]; account: Account | null; onClose: () => void; onSave: (account: Account) => void }) {
  const [name, setName] = useState(account?.name ?? '');
  const [type, setType] = useState<AccountType>(account?.type ?? 'wallet');
  const [initialBalance, setInitialBalance] = useState(() => (account ? formatCurrencyInput(account.initialBalance) : ''));
  const [color, setColor] = useState(account?.color ?? ACCOUNT_COLORS[0] ?? '#1B99D8');
  const [icon, setIcon] = useState<AccountIconKey>(account?.icon ?? 'wallet');
  const [accountIconPickerOpen, setAccountIconPickerOpen] = useState(false);
  const [accountColorPickerOpen, setAccountColorPickerOpen] = useState(false);
  const [error, setError] = useState('');

  function submit(event: FormEvent) {
    event.preventDefault();
    const normalizedName = name.trim();
    if (!normalizedName) return;
    if (accounts.some((item) => item.id !== account?.id && item.name.toLowerCase() === normalizedName.toLowerCase())) {
      setError('Já existe uma conta com esse nome.');
      return;
    }
    onSave({
      id: account?.id ?? crypto.randomUUID(),
      name: normalizedName,
      type,
      initialBalance: parseAmount(initialBalance),
      color,
      icon,
    });
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card settle-modal" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">{account ? 'Editar conta' : 'Nova conta'}</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body">
          <div className="modal-form-grid">
            <label className="form-field form-field-full">
              <span className="form-label">Nome da conta</span>
              <input className="form-input" value={name} onChange={(event) => { setName(event.target.value); setError(''); }} placeholder="Ex.: Nubank, Caixa, Banco do Brasil" />
            </label>
            <label className="form-field form-field-full">
              <span className="form-label">Tipo</span>
              <select className="form-input" value={type} onChange={(event) => setType(event.target.value as AccountType)}>
                {Object.entries(ACCOUNT_TYPE_LABELS).map(([value, label]) => <option key={value} value={value}>{label}</option>)}
              </select>
            </label>
            <div className="category-selector-grid">
              <div className="form-field">
                <span className="form-label form-label-optional">Ícone</span>
                <button
                  type="button"
                  className="form-input category-selector-btn"
                  aria-haspopup="dialog"
                  aria-expanded={accountIconPickerOpen}
                  onClick={() => setAccountIconPickerOpen(true)}
                >
                  <span className="category-selector-icon" style={{ backgroundColor: `${color}18`, color }}>
                    <AccountIconGraphic icon={icon} size={16} />
                  </span>
                  <span className="category-selector-label">{ACCOUNT_ICON_OPTIONS.find((option) => option.key === icon)?.label ?? 'Ícone'}</span>
                </button>
              </div>

              <div className="form-field">
                <span className="form-label form-label-optional">Cor</span>
                <button
                  type="button"
                  className="form-input category-selector-btn"
                  aria-haspopup="dialog"
                  aria-expanded={accountColorPickerOpen}
                  onClick={() => setAccountColorPickerOpen(true)}
                >
                  <span className="category-selector-color" style={{ backgroundColor: color }} />
                  <span className="category-selector-label">Escolher cor</span>
                </button>
              </div>
            </div>
            <label className="form-field form-field-full">
              <span className="form-label form-label-optional">Saldo inicial</span>
              <div className="form-input-wrap">
                <span className="form-input-prefix">R$</span>
                <input className="form-input" value={initialBalance} onChange={(event) => setInitialBalance(formatCurrencyInput(event.target.value))} placeholder="0,00" inputMode="decimal" />
              </div>
            </label>
            {error ? <div className="form-error form-field-full"><AlertCircle size={15} /> {error}</div> : null}
          </div>
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary">{account ? 'Salvar conta' : 'Criar conta'}</button>
        </div>
      </form>
      {accountColorPickerOpen ? <ColorSpectrumSheet value={color} onChange={setColor} onClose={() => setAccountColorPickerOpen(false)} /> : null}
      {accountIconPickerOpen
        ? createPortal(
          <div className="category-icon-picker-layer" onClick={() => setAccountIconPickerOpen(false)}>
            <div className="category-icon-picker-dialog" role="dialog" aria-modal="true" aria-label="Escolher ícone" onClick={(event) => event.stopPropagation()}>
              <div className="category-icon-picker-head">
                <strong>Escolher ícone</strong>
                <button type="button" onClick={() => setAccountIconPickerOpen(false)} aria-label="Fechar seletor"><X size={18} /></button>
              </div>
              <div className="category-icon-picker-grid">
                {ACCOUNT_ICON_OPTIONS.map((option) => (
                  <button key={option.key} type="button" className={icon === option.key ? 'active' : ''} style={{ color }} aria-label={`Selecionar ${option.label}`} aria-pressed={icon === option.key} onClick={() => { setIcon(option.key); setAccountIconPickerOpen(false); }}>
                    <AccountIconGraphic icon={option.key} size={20} />
                  </button>
                ))}
              </div>
            </div>
          </div>,
          document.body,
        )
        : null}
    </div>
  );
}

function TransactionSummaryCard({ summary, activeMode, onModeChange }: { summary: { income: number; expense: number; balance: number; pendingIncome?: number; pendingExpense?: number }; activeMode: TransactionSummaryMode; onModeChange: (mode: TransactionSummaryMode) => void }) {
  const options: Array<{ key: TransactionSummaryMode; label: string; value: number; icon: ReactNode }> = [
    { key: 'income', label: 'Receita', value: summary.income, icon: <ArrowUpRight size={19} /> },
    { key: 'expense', label: 'Despesas', value: summary.expense, icon: <ArrowDownLeft size={19} /> },
    { key: 'balance', label: 'Saldo', value: summary.balance, icon: <CircleDollarSign size={19} /> },
  ];
  const active = options.find((option) => option.key === activeMode) ?? options[2];

  return (
    <article className={`transaction-summary-card transaction-summary-card--${active.key}`}>
      <div className="transaction-summary-main">
        <span className="transaction-summary-label">{active.label}</span>
        <strong className="transaction-summary-value">{formatCurrency(active.value)}</strong>
      </div>
      <div className="transaction-summary-actions" role="tablist" aria-label="Resumo financeiro">
        {options.map((option) => (
          <button
            key={option.key}
            type="button"
            className={`transaction-summary-action${activeMode === option.key ? ' active' : ''}`}
            onClick={() => onModeChange(option.key)}
            role="tab"
            aria-selected={activeMode === option.key}
            title={option.label}
          >
            <span className="transaction-summary-action-icon">{option.icon}</span>
            <span>{option.label}</span>
          </button>
        ))}
      </div>
      <div className="transaction-summary-pending-bar">
        <span>Falta receber: <strong>{formatCurrency(summary.pendingIncome ?? 0)}</strong></span>
        <span className="pending-divider">•</span>
        <span>Falta pagar: <strong>{formatCurrency(summary.pendingExpense ?? 0)}</strong></span>
      </div>
    </article>
  );
}

function TransactionSummaryStats({ summary }: { summary: { income: number; expense: number; balance: number; pendingIncome?: number; pendingExpense?: number } }) {
  const settledIncome = summary.income - (summary.pendingIncome ?? 0);
  const settledExpense = summary.expense - (summary.pendingExpense ?? 0);
  return (
    <div className="transaction-summary-stats">
      <MetricCard label="Receita" value={formatCurrency(summary.income)} tone="info" icon={<TrendingUp size={18} />}
        sub1={`Recebido: ${formatCurrency(settledIncome)}`} sub2={`A receber: ${formatCurrency(summary.pendingIncome ?? 0)}`} />
      <MetricCard label="Despesas" value={formatCurrency(summary.expense)} tone="danger" icon={<TrendingDown size={18} />}
        sub1={`Pago: ${formatCurrency(settledExpense)}`} sub2={`A pagar: ${formatCurrency(summary.pendingExpense ?? 0)}`} />
      <MetricCard
        label="Saldo"
        value={formatCurrency(summary.balance)}
        tone={summary.balance >= 0 ? 'info' : 'danger'}
        icon={<CircleDollarSign size={18} />}
        sub1={`Realizado: ${formatCurrency(settledIncome - settledExpense)}`}
      />
    </div>
  );
}
function MetricCard({ label, value, icon, tone, sub1, sub2 }: { label: string; value: string; icon: ReactNode; tone: string; sub1?: string; sub2?: string }) {
  return (
    <article className={`signature-stat-card signature-stat-card--${tone}`}>
      <span className="signature-stat-icon">{icon}</span>
      <span className="signature-stat-copy">
        <span className="signature-stat-label">{label}</span>
        <strong className="signature-stat-value">{value}</strong>
        {(sub1 || sub2) && (
          <span className="signature-stat-subs">
            {sub1 && <span>{sub1}</span>}
            {sub2 && <span>{sub2}</span>}
          </span>
        )}
      </span>
    </article>
  );
}

function LaunchModal({ accounts, incomeCategories, expenseCategories, initialType = 'expense', onClose, onCreate }: { accounts: Account[]; incomeCategories: string[]; expenseCategories: string[]; initialType?: TransactionType; onClose: () => void; onCreate: (items: Transaction[]) => void }) {
  const [type, setType] = useState<TransactionType>(initialType);
  const [form, setForm] = useState<LaunchForm>(() => ({ ...initialForm, category: (initialType === 'income' ? incomeCategories[0] : expenseCategories[0]) ?? '' }));

  function update<K extends keyof LaunchForm>(key: K, value: LaunchForm[K]) {
    setForm((current) => ({ ...current, [key]: value }));
  }

  function selectTransactionType(nextType: TransactionType) {
    const categories = nextType === 'income' ? incomeCategories : expenseCategories;
    setType(nextType);
    setForm((current) => ({ ...current, category: categories[0] ?? '' }));
  }

  function submit(event: FormEvent) {
    event.preventDefault();
    if (!form.description.trim() || parseAmount(form.amount) <= 0) return;
    const account = accounts.find((candidate) => candidate.id === form.accountId);
    onCreate(generateTransactions(form, type).map((item) => ({ ...item, accountId: account?.id, account: account?.name })));
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">Nova transação</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>

        <div className="modal-body">
          <div className="modal-form-grid">
            <div className="form-field form-field-full">
              <span className="form-label">Tipo</span>
              <div className="launch-type-segmented" role="group" aria-label="Tipo da transação">
                <button type="button" className={`launch-type-segment launch-type-segment--expense${type === 'expense' ? ' active' : ''}`} aria-pressed={type === 'expense'} onClick={() => selectTransactionType('expense')}>
                  <ArrowDownLeft size={17} /> Despesa
                </button>
                <button type="button" className={`launch-type-segment launch-type-segment--income${type === 'income' ? ' active' : ''}`} aria-pressed={type === 'income'} onClick={() => selectTransactionType('income')}>
                  <ArrowUpRight size={17} /> Receita
                </button>
              </div>
            </div>
            <label className="form-field form-field-full">
              <span className="form-label">Descrição</span>
              <input className="form-input" value={form.description} maxLength={TRANSACTION_DESCRIPTION_MAX_LENGTH} onChange={(event) => update('description', event.target.value)} placeholder="Ex.: Aluguel, internet, venda..." />
            </label>
            <label className="form-field">
              <span className="form-label">Valor</span>
              <div className="form-input-wrap">
                <span className="form-input-prefix">R$</span>
                <input className="form-input" value={form.amount} onChange={(event) => update('amount', formatCurrencyInput(event.target.value))} placeholder="0,00" inputMode="decimal" />
              </div>
            </label>
            <label className="form-field">
              <span className="form-label">Vencimento</span>
              <div className="form-input-wrap">
                <CalendarDays size={16} />
                <input className="form-input" type="date" value={form.dueDate} onChange={(event) => {
                  const dueDate = event.target.value;
                  setForm((current) => ({ ...current, dueDate, fixedUntil: current.fixedUntil && current.fixedUntil >= dueDate.slice(0, 7) ? current.fixedUntil : `${dueDate.slice(0, 4)}-12` }));
                }} />
              </div>
            </label>
            <label className="form-field">
              <span className="form-label">Categoria</span>
              <select className="form-input" value={form.category} onChange={(event) => update('category', event.target.value)}>
                {(type === 'income' ? incomeCategories : expenseCategories).map((category) => <option key={category}>{category}</option>)}
              </select>
            </label>
            <label className="form-field">
              <span className="form-label">Recorrência</span>
              <select className="form-input" value={form.recurrence} onChange={(event) => {
                const next = event.target.value as RecurrenceType;
                setForm((current) => ({ ...current, recurrence: next, fixedUntil: next === 'fixed' && !current.fixedUntil ? `${current.dueDate.slice(0, 4)}-12` : current.fixedUntil }));
              }}>
                <option value="single">Única</option>
                <option value="fixed">Fixa</option>
                <option value="installment">Parcelada</option>
              </select>
            </label>
            <label className="form-field form-field-full">
              <span className="form-label form-label-optional">Banco / conta</span>
              <select className="form-input" value={form.accountId} onChange={(event) => update('accountId', event.target.value)}>
                <option value="">Selecione um banco ou conta</option>
                {accounts.map((account) => <option key={account.id} value={account.id}>{account.name}</option>)}
              </select>
            </label>
            {form.recurrence === 'fixed' ? (
              <label className="form-field">
                <span className="form-label">Repetir até</span>
                <input className="form-input" type="month" min={form.dueDate.slice(0, 7)} value={form.fixedUntil || form.dueDate.slice(0, 7)} onChange={(event) => update('fixedUntil', event.target.value)} />
              </label>
            ) : null}
            {form.recurrence === 'installment' ? (
              <label className="form-field">
                <span className="form-label">Parcelas</span>
                <input className="form-input" type="number" min="1" max="120" value={form.installments} onChange={(event) => update('installments', event.target.value)} />
              </label>
            ) : null}
          </div>
        </div>

        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary">Criar transação</button>
        </div>
      </form>
    </div>
  );
}
function ImportPreviewModal({ rows, onClose }: { rows: ImportPreviewRow[]; onClose: () => void }) {
  return createPortal(
    <div className="modal-overlay import-preview-overlay" onClick={onClose}>
      <div className="modal-card import-preview-modal" role="dialog" aria-modal="true" aria-label="Pré-visualização" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">Pré-visualização</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="import-preview-modal-body">
          <div className="import-preview">
            <div className="import-preview-head">
              <span>Linha</span><span>Tipo</span><span>Descrição</span><span>Categoria</span><span>Vencimento</span><span>Valor</span><span>Status</span>
            </div>
            {rows.map((row) => (
              <div key={row.rowNumber} className={`import-preview-row${row.error ? ' invalid' : ''}`}>
                <span>{row.rowNumber}</span>
                <span>{row.type === 'income' ? 'Receita' : row.type === 'expense' ? 'Despesa' : '-'}</span>
                <strong>{row.description || '-'}</strong>
                <span>{row.category || '-'}</span>
                <span>{row.dueDate ? formatDate(row.dueDate) : '-'}</span>
                <span>{row.amount > 0 ? formatCurrency(row.amount) : '-'}</span>
                <span className={row.error ? 'import-row-error' : 'import-row-valid'}>{row.error || 'Válida'}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>,
    document.body
  );
}

function ImportTransactionsModal({ onClose, onImport, existingCategories }: { onClose: () => void; onImport: (items: Transaction[], newCategories: CustomCategory[]) => void; existingCategories: CustomCategory[] }) {
  const [fileName, setFileName] = useState('');
  const [rows, setRows] = useState<ImportPreviewRow[]>([]);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [step, setStep] = useState<'review' | 'confirm-categories' | 'cancelled'>('review');

  const validRows = rows.filter((row) => !row.error);
  const importedTransactions = validRows.flatMap((row) => row.transactions);
  const invalidCount = rows.length - validRows.length;

  const MONTH_NAMES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  const coveredMonths = Array.from(new Set(importedTransactions.map((t) => t.dueDate.slice(0, 7)))).sort();
  const formattedMonths = coveredMonths.map((m) => {
    const [y, mo] = m.split('-');
    return `${MONTH_NAMES[Number(mo) - 1]}/${y.slice(2)}`;
  }).join(', ');

  const existingNames = new Set(existingCategories.map((c) => c.name.toLowerCase().trim()));
  const newCategoryMap = new Map<string, CustomCategory>();
  for (const row of validRows) {
    const name = row.category.trim();
    if (name && !existingNames.has(name.toLowerCase()) && !newCategoryMap.has(name.toLowerCase())) {
      const kind: CustomCategory['kind'] = row.type === 'income' ? 'income' : row.type === 'expense' ? 'expense' : 'transfer';
      newCategoryMap.set(name.toLowerCase(), { id: crypto.randomUUID(), name, kind, color: '#94a3b8', icon: 'tag' });
    }
  }
  const newCategories = Array.from(newCategoryMap.values());

  async function handleFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    setFileName(file.name.slice(0, 180));
    setRows([]);
    setError('');
    setPreviewOpen(false);
    setStep('review');

    if (!file.name.toLowerCase().endsWith('.xlsx')) { setError('Selecione um arquivo Excel no formato .xlsx.'); return; }
    if (file.size > MAX_IMPORT_FILE_SIZE) { setError('O arquivo deve ter no máximo 5 MB.'); return; }
    if (!(await isSafeXlsxContainer(file))) { setError('O conteúdo do arquivo não corresponde a uma planilha XLSX válida.'); return; }

    setLoading(true);
    try {
      const parsedRows = await parseExcelTransactions(file);
      if (!parsedRows.length) throw new Error('Nenhuma transação foi encontrada na planilha.');
      setRows(parsedRows);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Não foi possível ler a planilha.');
    } finally {
      setLoading(false);
    }
  }

  function handleImportClick() {
    if (newCategories.length > 0) { setStep('confirm-categories'); return; }
    onImport(importedTransactions, []);
  }

  function handleConfirmCategories() {
    onImport(importedTransactions, newCategories);
  }

  function handleCancelCategories() {
    setStep('cancelled');
  }

  const KIND_LABELS: Record<CustomCategory['kind'], string> = { income: 'Receita', expense: 'Despesa', transfer: 'Transferência' };

  return (
    <>
      {previewOpen ? <ImportPreviewModal rows={rows} onClose={() => setPreviewOpen(false)} /> : null}
      <div className="modal-overlay" onClick={onClose}>
        <div className="modal-card import-modal" role="dialog" aria-modal="true" aria-labelledby="import-modal-title" onClick={(event) => event.stopPropagation()}>
          <div className="modal-header">
            <h2 className="modal-title modal-title-only" id="import-modal-title">Importar transações</h2>
            <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
          </div>

          {step === 'confirm-categories' ? (
            <>
              <div className="modal-body import-modal-body">
                <div className="import-alert import-alert--warning">
                  <AlertCircle size={17} />
                  <span>A planilha contém <strong>{newCategories.length}</strong> {newCategories.length === 1 ? 'categoria que não existe' : 'categorias que não existem'} no sistema. Deseja criá-{newCategories.length === 1 ? 'la' : 'las'} automaticamente e continuar a importação?</span>
                </div>
                <div className="import-new-categories">
                  {newCategories.map((cat) => (
                    <div key={cat.id} className="import-new-category-row">
                      <span className="import-new-category-dot" style={{ background: cat.color }} />
                      <span className="import-new-category-name">{cat.name}</span>
                      <span className={`category-type category-type--${cat.kind}`}>{KIND_LABELS[cat.kind]}</span>
                    </div>
                  ))}
                </div>
              </div>
              <div className="modal-actions">
                <button type="button" className="button-secondary" onClick={handleCancelCategories}>Não criar, cancelar importação</button>
                <button type="button" className="button-primary" onClick={handleConfirmCategories}>Criar categorias e importar</button>
              </div>
            </>
          ) : step === 'cancelled' ? (
            <>
              <div className="modal-body import-modal-body">
                <div className="import-alert import-alert--error">
                  <AlertCircle size={17} />
                  <span>A importação foi cancelada porque há categorias não cadastradas. Ajuste as categorias na planilha ou cadastre-as manualmente antes de importar.</span>
                </div>
              </div>
              <div className="modal-actions">
                <button type="button" className="button-secondary" onClick={() => setStep('review')}>Voltar</button>
                <button type="button" className="button-primary" onClick={onClose}>Fechar</button>
              </div>
            </>
          ) : (
            <>
              <div className="modal-body import-modal-body">
                <label className={`import-upload${loading ? ' loading' : ''}`}>
                  <input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={handleFile} disabled={loading} />
                  <span className="import-upload-icon"><Upload size={22} /></span>
                  <strong>{loading ? 'Lendo planilha...' : fileName || 'Selecionar planilha'}</strong>
                  <small>Arquivo Excel no formato .xlsx</small>
                </label>

                <div className="import-columns">
                  <div className="import-columns-info">
                    <strong>Colunas obrigatórias</strong>
                    <span>Tipo, Descrição, Valor, Vencimento e Categoria</span>
                    <small>Opcionais: Recorrência e Parcelas</small>
                    <small className="import-recurrence-hint">Recorrência: <em>Fixa</em> expande até dezembro do ano da data · <em>Parcelada</em> expande pelo nº de parcelas</small>
                  </div>
                  <button type="button" className="import-template-button" onClick={downloadImportTemplate}>
                    <FileSpreadsheet size={16} /> Baixar modelo
                  </button>
                </div>

                {error ? <div className="import-alert import-alert--error"><AlertCircle size={17} /><span>{error}</span></div> : null}

                {rows.length ? (
                  <div className="import-info-card">
                    <div className="import-info-row">
                      <span className="import-info-label">Linhas válidas</span>
                      <strong className="import-info-value">{validRows.length}</strong>
                    </div>
                    <div className="import-info-row">
                      <span className="import-info-label">Transações a criar</span>
                      <strong className="import-info-value">{importedTransactions.length}</strong>
                    </div>
                    {invalidCount > 0 ? (
                      <div className="import-info-row import-info-row--error">
                        <span className="import-info-label">Com erro</span>
                        <strong className="import-info-value">{invalidCount}</strong>
                      </div>
                    ) : null}
                    {coveredMonths.length > 0 ? (
                      <div className="import-info-row">
                        <span className="import-info-label">Meses cobertos</span>
                        <span className="import-info-months">{formattedMonths}</span>
                      </div>
                    ) : null}
                    {newCategories.length > 0 ? (
                      <div className="import-info-row import-info-row--warning">
                        <span className="import-info-label">Categorias novas</span>
                        <strong className="import-info-value import-info-value--warning">{newCategories.length} serão criadas</strong>
                      </div>
                    ) : null}
                    <div className="import-info-row import-info-row--action">
                      <button type="button" className="import-preview-btn" onClick={() => setPreviewOpen(true)}>
                        <Eye size={14} /> Ver pré-visualização
                      </button>
                    </div>
                  </div>
                ) : null}
              </div>

              <div className="modal-actions">
                <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
                <button type="button" className="button-primary" disabled={!importedTransactions.length || loading} onClick={handleImportClick}>
                  Importar {importedTransactions.length || ''}
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </>
  );
}

function SettleModal({ item, accounts, onClose, onSave }: { item: Transaction; accounts: Account[]; onClose: () => void; onSave: (accountId: string, settledAt: string, settledAmount: number) => void }) {
  const [accountId, setAccountId] = useState(() => accounts.find((account) => account.id === item.accountId)?.id ?? '');
  const [settledAmount, setSettledAmount] = useState(() => formatCurrencyInput(item.settledAmount ?? item.amount));
  const [settledAt, setSettledAt] = useState(new Date().toISOString().slice(0, 10));
  const [error, setError] = useState('');

  function submit(event: FormEvent) {
    event.preventDefault();
    const amount = parseAmount(settledAmount);
    if (!accountId) { setError('Selecione o banco ou conta utilizado.'); return; }
    if (amount <= 0) { setError('Informe um valor válido.'); return; }
    onSave(accountId, settledAt, amount);
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card settle-modal" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">{item.type === 'income' ? 'Registrar recebimento' : 'Registrar pagamento'}</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="modal-body">
          <div className="modal-form-grid">
            <label className="form-field form-field-full">
              <span className="form-label">Banco / conta</span>
              <select className="form-input" value={accountId} required onChange={(event) => { setAccountId(event.target.value); setError(''); }}>
                <option value="">Selecione o banco ou conta</option>
                {accounts.map((account) => <option key={account.id} value={account.id}>{account.name}</option>)}
              </select>
            </label>
            <label className="form-field form-field-full">
              <span className="form-label">Valor {item.type === 'income' ? 'recebido' : 'pago'}</span>
              <div className="form-input-wrap">
                <span className="form-input-prefix">R$</span>
                <input className="form-input" value={settledAmount} onChange={(event) => { setSettledAmount(formatCurrencyInput(event.target.value)); setError(''); }} inputMode="decimal" placeholder="0,00" />
              </div>
            </label>
            <label className="form-field form-field-full">
              <span className="form-label">Data da baixa</span>
              <input className="form-input" type="date" value={settledAt} onChange={(event) => setSettledAt(event.target.value)} />
            </label>
            {error ? <div className="form-error form-field-full"><AlertCircle size={15} /> {error}</div> : null}
            {!accounts.length ? <div className="form-error form-field-full"><AlertCircle size={15} /> Cadastre um banco ou conta antes de efetivar a transação.</div> : null}
          </div>
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary" disabled={!accounts.length}>Confirmar baixa</button>
        </div>
      </form>
    </div>
  );
}










