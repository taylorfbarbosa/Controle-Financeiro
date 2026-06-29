import { Fragment, useEffect, useLayoutEffect, useMemo, useRef, useState, type ChangeEvent, type FormEvent, type ReactNode, type RefObject } from 'react';
import { readSheet, type Row } from 'read-excel-file/browser';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { Bar, BarChart, CartesianGrid, LabelList, Legend, Line, LineChart, ReferenceLine, ResponsiveContainer, Tooltip as RechartsTooltip, XAxis, YAxis } from 'recharts';
import type { Session } from '@supabase/supabase-js';
import caesarFinanceLogo from './assets/caesar-finance-logo.png';
import { supabase } from './lib/supabase';
import { loadAll, syncAccounts, syncCategories, syncGoals, syncTransactions } from './lib/db';
import {
  AlertCircle,
  BadgeDollarSign,

  ArrowLeftRight,
  ArrowDownLeft,
  ArrowUpRight,
  Banknote,
  BriefcaseBusiness,
  CalendarDays,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  CircleDollarSign,
  Car,
  Clock3,
  Coins,
  CreditCard,
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
  Minus,
  MoreVertical,
  PiggyBank,
  PawPrint,
  Pencil,
  Plus,
  ReceiptText,
  RotateCcw,
  Repeat2,
  Share,
  ShoppingBag,
  Target,
  Tags,
  TrendingDown,
  TrendingUp,
  Trash2,
  Utensils,
  Users,
  UserCircle,
  Upload,
  WalletCards,
  X,
} from 'lucide-react';

type TransactionType = 'income' | 'expense';
type AppPage = 'dashboard' | 'transactions' | 'categories' | 'goals' | 'reports' | 'accounts' | 'help';
const PAGE_LABELS: Record<AppPage, string> = {
  dashboard: 'Visão Geral',
  transactions: 'Transações',
  categories: 'Categorias',
  goals: 'Metas',
  reports: 'Relatórios',
  accounts: 'Contas',
  help: 'Ajuda',
};
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
type CategoryIconKey = 'salary' | 'money' | 'work' | 'shopping' | 'refund' | 'investment' | 'home' | 'card' | 'gift' | 'food' | 'car' | 'health' | 'education' | 'leisure' | 'repeat' | 'family' | 'pets' | 'donation' | 'wallet' | 'transfer' | 'tag';

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

type BalanceLabelProps = {
  x?: number | string;
  y?: number | string;
  value?: unknown;
};

type BalanceDotProps = {
  cx?: number | string;
  cy?: number | string;
  payload?: { Saldo?: number };
};

const TRANSACTION_DESCRIPTION_MAX_LENGTH = 24;

type LaunchForm = {
  description: string;
  category: string;
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

const DEFAULT_ACCOUNT: Account = { id: 'wallet', name: 'Carteira', type: 'wallet', initialBalance: 0, color: '#1B99D8', icon: 'wallet' };
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
const CATEGORY_ICON_OPTIONS: Array<{ key: CategoryIconKey; label: string }> = [
  { key: 'tag', label: 'Etiqueta' },
  { key: 'money', label: 'Dinheiro' },
  { key: 'work', label: 'Trabalho' },
  { key: 'shopping', label: 'Compras' },
  { key: 'home', label: 'Moradia' },
  { key: 'food', label: 'Alimentação' },
  { key: 'car', label: 'Transporte' },
  { key: 'health', label: 'Saúde' },
  { key: 'education', label: 'Educação' },
  { key: 'gift', label: 'Presente' },
  { key: 'family', label: 'Família' },
  { key: 'transfer', label: 'Transferência' },
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
    .replace(/[̀-ͯ]/g, '')
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

function exportReportExcel(items: Transaction[], title: string) {
  const totals = reportTotals(items);
  const workbook = XLSX.utils.book_new();

  // Summary sheet
  const summaryData = [
    ['', 'Receitas', 'Despesas', 'Saldo'],
    ['Previsto', totals.income, totals.expense, totals.balance],
    ['Realizado', totals.settledIncome, totals.settledExpense, totals.settledIncome - totals.settledExpense],
    ['Pendente', totals.pendingIncome, totals.pendingExpense, totals.pendingIncome - totals.pendingExpense],
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  summarySheet['!cols'] = [{ wch: 14 }, { wch: 18 }, { wch: 18 }, { wch: 18 }];
  for (let row = 2; row <= 4; row++) {
    ['B', 'C', 'D'].forEach((col) => {
      const cell = summarySheet[`${col}${row}`];
      if (cell && typeof cell.v === 'number') cell.z = '"R$" #,##0.00';
    });
  }
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Resumo');

  // Transactions sheet
  const rows = items.map((item) => ({
    Tipo: item.type === 'income' ? 'Receita' : 'Despesa',
    'Descrição': item.description,
    Categoria: item.category,
    Vencimento: formatDate(item.dueDate),
    'Valor previsto': item.amount,
    'Valor real': item.status === 'settled' ? (item.settledAmount ?? item.amount) : '',
    Status: item.status === 'settled' ? (item.type === 'income' ? 'Recebido' : 'Pago') : 'Aberto',
    Conta: item.status === 'settled' ? (item.account ?? '') : '',
  }));
  const worksheet = XLSX.utils.json_to_sheet(rows, { header: REPORT_COLUMNS });
  worksheet['!cols'] = [{ wch: 10 }, { wch: 28 }, { wch: 18 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 16 }];
  const range = XLSX.utils.decode_range(worksheet['!ref'] ?? 'A1');
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    ['E', 'F'].forEach((column) => {
      const cell = worksheet[`${column}${row + 1}`];
      if (cell && typeof cell.v === 'number') cell.z = '"R$" #,##0.00';
    });
  }
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Transações');

  XLSX.writeFile(workbook, `${slugifyFileName(title)}.xlsx`);
}

function downloadImportTemplate() {
  const headers = ['Tipo', 'Descrição', 'Valor', 'Vencimento', 'Categoria', 'Recorrência', 'Parcelas'];
  const rows = [
    { Tipo: 'Receita', 'Descrição': 'Salário', Valor: 3500, Vencimento: '05/01/2026', Categoria: 'Salário', 'Recorrência': 'Fixa', Parcelas: '' },
    { Tipo: 'Despesa', 'Descrição': 'Aluguel', Valor: 1200, Vencimento: '10/01/2026', Categoria: 'Moradia', 'Recorrência': 'Fixa', Parcelas: '' },
    { Tipo: 'Despesa', 'Descrição': 'Notebook', Valor: 4800, Vencimento: '15/01/2026', Categoria: 'Compras', 'Recorrência': 'Parcelada', Parcelas: 12 },
    { Tipo: 'Receita', 'Descrição': 'Freelance', Valor: 800, Vencimento: '20/01/2026', Categoria: 'Freelance', 'Recorrência': 'Única', Parcelas: '' },
  ];
  const worksheet = XLSX.utils.json_to_sheet(rows, { header: headers });
  worksheet['!cols'] = [{ wch: 10 }, { wch: 26 }, { wch: 12 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 10 }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Modelo');
  XLSX.writeFile(workbook, 'modelo-importacao-transacoes.xlsx');
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
    'Salário': 'salary', Extra: 'money', Freelance: 'work', Vendas: 'shopping', Reembolso: 'refund', Investimentos: 'investment',
    Aluguel: 'home', Cashback: 'card', Presente: 'gift', Empréstimo: 'money', Moradia: 'home', Alimentação: 'food',
    Transporte: 'car', Saúde: 'health', Educação: 'education', Lazer: 'leisure', Compras: 'shopping', Assinaturas: 'repeat',
    Dívidas: 'money', Cartão: 'card', Família: 'family', Pets: 'pets', Impostos: 'money', Trabalho: 'work', Doações: 'donation',
    Saques: 'wallet', 'Entre contas': 'transfer', 'Para carteira': 'wallet', 'Para poupança': 'money', 'Para investimentos': 'investment',
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
  return <Tags size={size} />;
}

function parseAmount(value: string) {
  const normalized = value.replace(/\./g, '').replace(',', '.').replace(/[^\d.]/g, '');
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
  const clean = raw.replace(/[^\d,]/g, '');
  if (!clean) return '';
  const parts = clean.split(',');
  let intPart = parts[0].replace(/^0+(?=\d)/, '');
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
  const text = String(value ?? '').replace(/[^\d,.-]/g, '').trim();
  if (!text) return 0;
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

async function parseExcelTransactions(file: File): Promise<ImportPreviewRow[]> {
  const rows: Row[] = await readSheet(file);
  if (rows.length < 2) throw new Error('A planilha precisa ter um cabeçalho e pelo menos uma transação.');

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
      amount: String(amount),
      dueDate,
      category,
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
    <div className="splash-screen" role="status" aria-label="Carregando Caesar Finance">
      <img className="splash-logo" src={caesarFinanceLogo} alt="Caesar Finance" />
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

function LoginScreen() {
  const [mode, setMode] = useState<'login' | 'recovery' | 'register'>('login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);

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
      if (password !== confirmPassword) {
        setMessage({ text: 'As senhas não coincidem. Verifique e tente novamente.', type: 'error' });
        return;
      }
      setLoading(true);
      const { data, error } = await supabase.auth.signUp({ email, password });
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
    if (error) setMessage({ text: translateAuthError(error.message), type: 'error' });
    // Sucesso: a sessão muda e a App troca de tela automaticamente.
  }

  function switchMode(newMode: 'login' | 'recovery' | 'register') {
    setMode(newMode);
    setMessage(null);
  }

  return (
    <main className="toodledo-login-shell" aria-label="Login Caesar Finance">
      <div className="toodledo-login-container">
        <div className="toodledo-brand-side">
          <div className="toodledo-logo-wrapper">
            <img className="toodledo-logo-image" src={caesarFinanceLogo} alt="Caesar Finance" />
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
                <input
                  value={password}
                  onChange={(event) => setPassword(event.target.value)}
                  type="password"
                  placeholder={mode === 'register' ? 'Criar senha' : 'Senha'}
                  autoComplete="current-password"
                  required
                  className="toodledo-input"
                />
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
    if (password.length < 6) {
      setMessage({ text: 'A senha deve ter pelo menos 6 caracteres.', type: 'error' });
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
            <img className="toodledo-logo-image" src={caesarFinanceLogo} alt="Caesar Finance" />
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
  const userId = session?.user.id;
  const [showSplash, setShowSplash] = useState(
    () => typeof window !== 'undefined' && window.matchMedia('(max-width: 760px)').matches,
  );
  const [activePage, setActivePage] = useState<AppPage>('dashboard');
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [customCategories, setCustomCategories] = useState<CustomCategory[]>([]);
  const [goals, setGoals] = useState<Goal[]>([]);
  const [referenceDate, setReferenceDate] = useState(() => new Date(2026, 5, 1));
  const [launchOpen, setLaunchOpen] = useState(false);
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
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setAuthReady(true);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((event, nextSession) => {
      if (event === 'PASSWORD_RECOVERY') setRecoveryMode(true);
      setSession(nextSession);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  // Carrega os dados do usuário quando há sessão; limpa ao sair.
  useEffect(() => {
    if (!userId) {
      setAccounts([]);
      setCustomCategories([]);
      setTransactions([]);
      setGoals([]);
      return;
    }
    let cancelled = false;
    setDataLoading(true);
    setSyncError(null);
    loadAll()
      .then((data) => {
        if (cancelled) return;
        setAccounts(data.accounts);
        setCustomCategories(data.categories);
        setTransactions(data.transactions);
        setGoals(data.goals);
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
    setActivePage((current) => {
      if (current === page) scrollContentToTop();
      return page;
    });
  }

  useLayoutEffect(() => {
    scrollContentToTop();
  }, [activePage]);

  useEffect(() => {
    if (!filterOpen) return;

    function closeOnOutsideClick(event: PointerEvent) {
      if (!filterControlRef.current?.contains(event.target as Node)) setFilterOpen(false);
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
  const monthItems = transactions
    .filter((item) => (dateFilter ? item.dueDate === dateFilter : item.dueDate.slice(0, 7) === currentMonth))
    .filter((item) => categoryFilter === 'all' || item.category === categoryFilter)
    .filter((item) => typeFilter === 'all' || item.type === typeFilter)
    .sort((a, b) => a.dueDate.localeCompare(b.dueDate));

  const selectedInView = monthItems.filter((item) => selectedTxIds.has(item.id));
  const selectedOpenInView = selectedInView.filter((item) => item.status === 'open');
  const selectedSettledInView = selectedInView.filter((item) => item.status === 'settled');
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
    const fallbackAccount = accounts[0];
    const settledAt = new Date().toISOString().slice(0, 10);

    commitTransactions(transactions.map((item) => {
      if (!selectedIds.has(item.id) || item.status === 'settled') return item;
      const account = accounts.find((candidate) => candidate.id === item.accountId || candidate.name === item.account) ?? fallbackAccount;
      return {
        ...item,
        status: 'settled',
        settledAt,
        settledAmount: item.amount,
        accountId: account?.id ?? item.accountId,
        account: account?.name ?? item.account,
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
      <Topbar activePage={activePage} onNavigate={handleNavigate} onLogout={() => supabase.auth.signOut()} onOpenTransactionFilters={openFilters} onImportTransactions={() => setImportOpen(true)} onOpenAccountFilters={openAccountFilters} accountFilterOpen={accountFilterOpen} accountActiveFilterCount={Number(Boolean(accountFilters.search.trim())) + Number(accountFilters.type !== 'all')} onOpenCategoryFilters={openCategoryPageFilters} categoryFilterOpen={categoryPageFilterOpen} categoryActiveFilterCount={Number(Boolean(categoryPageFilters.search.trim())) + Number(categoryPageFilters.type !== 'all')} onOpenGoalFilters={openGoalFilters} goalFilterOpen={goalFilterOpen} goalActiveFilterCount={Number(Boolean(goalFilters.search.trim())) + Number(goalFilters.status !== 'all')} />
      <div className="content-layout">
        <Sidebar activePage={activePage} onNavigate={handleNavigate} />
        <main ref={mainContentRef} className={`main-content main-content--${activePage}`}>
          {activePage === 'dashboard' ? (
            <DashboardPage
              transactions={transactions}
              goals={goals}
              referenceDate={referenceDate}
              onChangeDate={setReferenceDate}
              onNavigate={handleNavigate}
            />
          ) : activePage === 'transactions' ? (
            <>
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
                    <button type="button" className="bulk-bar-button bulk-bar-button--success" disabled={selectedOpenInView.length === 0} onClick={() => setConfirmDialog({
                      title: 'Marcar como efetivado',
                      message: `Deseja marcar ${selectedOpenInView.length} ${selectedOpenInView.length === 1 ? 'transação em aberto' : 'transações em aberto'} como efetivado?`,
                      confirmLabel: 'Efetivar',
                      onConfirm: bulkSettleSelectedTransactions,
                    })}><CheckCircle2 size={15} /> Marcar como efetivo</button>
                    <button type="button" className="bulk-bar-button" disabled={selectedSettledInView.length === 0} onClick={() => setConfirmDialog({
                      title: 'Marcar como pendente',
                      message: `Deseja marcar ${selectedSettledInView.length} ${selectedSettledInView.length === 1 ? 'transação efetivada' : 'transações efetivadas'} como pendente?`,
                      confirmLabel: 'Marcar pendente',
                      onConfirm: bulkMarkSelectedTransactionsAsPending,
                    })}><RotateCcw size={15} /> Marcar como pendente</button>
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
          ) : activePage === 'help' ? (
            <HelpPage />
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
          incomeCategories={incomeCategories}
          expenseCategories={expenseCategories}
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
          onImport={(items) => {
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
            const account = accounts.find((candidate) => candidate.id === accountId) ?? DEFAULT_ACCOUNT;
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

  useEffect(() => {
    if (!open) return;
    function onDown(event: PointerEvent) { if (!wrapRef.current?.contains(event.target as Node)) setOpen(false); }
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
      setPos({ top: rect.bottom + 6, left: Math.max(12, rect.right - 196) });
    }
    setOpen((value) => !value);
  }

  return (
    <div className="row-actions" ref={wrapRef}>
      <button type="button" ref={buttonRef} className={`row-actions-trigger${open ? ' active' : ''}`} aria-label="Ações" aria-haspopup="menu" aria-expanded={open} onClick={toggle}>
        <MoreVertical size={16} />
      </button>
      {open ? (
        <div className="row-actions-menu" role="menu" style={{ top: pos.top, left: pos.left }}>
          {actions.map((action) => (
            <button key={action.key} type="button" role="menuitem" className={`row-actions-item${action.danger ? ' row-actions-item--danger' : ''}`} onClick={() => { setOpen(false); action.onClick(); }}>
              {action.icon}
              {action.label}
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function LongPressOptionsModal({ item, onClose, onEdit, onSettle, onDelete, onMarkPending, onSelect }: { item: Transaction; onClose: () => void; onEdit: () => void; onSettle: () => void; onDelete: () => void; onMarkPending: () => void; onSelect: () => void }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card longpress-modal" role="dialog" aria-modal="true" onClick={(event) => event.stopPropagation()} style={{ width: 'min(100%, 360px)', padding: '20px' }}>
        <div className="modal-header" style={{ padding: '0 0 14px', borderBottom: '1px solid #e2e8f0', marginBottom: '14px' }}>
          <h2 className="modal-title modal-title-only" style={{ fontSize: '18px' }}>Opções da transação</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div style={{ marginBottom: '16px', fontSize: '14px', fontWeight: 700, color: '#334155' }}>
          {displayTransactionDescription(item.description)}
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
          <button type="button" className="button-secondary" style={{ justifyContent: 'flex-start', gap: '10px', height: '44px', color: '#1B99D8', borderColor: '#bae6fd', backgroundColor: '#eff8ff', fontWeight: 600 }} onClick={() => { onClose(); onSelect(); }}>
            <CheckCircle2 size={18} /> Selecionar
          </button>
          {item.status === 'open' ? (
            <button type="button" className="button-secondary" style={{ justifyContent: 'flex-start', gap: '10px', height: '44px', color: '#059669', borderColor: '#a7f3d0', backgroundColor: '#ecfdf5', fontWeight: 600 }} onClick={() => { onClose(); onSettle(); }}>
              <CheckCircle2 size={18} /> Marcar como efetivado
            </button>
          ) : (
            <button type="button" className="button-secondary" style={{ justifyContent: 'flex-start', gap: '10px', height: '44px', color: '#1B99D8', borderColor: '#bae6fd', backgroundColor: '#eff8ff', fontWeight: 600 }} onClick={() => { onClose(); onMarkPending(); }}>
              <RotateCcw size={18} /> Marcar como pendente
            </button>
          )}
          <button type="button" className="button-secondary" style={{ justifyContent: 'flex-start', gap: '10px', height: '44px', fontWeight: 600 }} onClick={() => { onClose(); onEdit(); }}>
            <Pencil size={18} /> Editar
          </button>
          <button type="button" className="button-danger" style={{ justifyContent: 'flex-start', gap: '10px', height: '44px', fontWeight: 600 }} onClick={() => { onClose(); onDelete(); }}>
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
          <span>Efetivar</span>
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
        <div className="launch-main">
          <strong>{displayTransactionDescription(item.description)}</strong>
          {item.notes ? <span className="launch-notes">{item.notes}</span> : null}
        </div>
        <span className="launch-category">
          {categoryMeta ? (
            <span className="launch-category-icon" style={{ backgroundColor: `${categoryMeta.color}18`, color: categoryMeta.color }}>
              <CategoryIconGraphic icon={categoryMeta.icon} size={14} />
            </span>
          ) : null}
          <span className="launch-category-name">{item.category}</span>
        </span>
        <span className={`status-icon transaction-kind transaction-kind--${item.type}`}>
          {item.type === 'income' ? <ArrowUpRight size={16} /> : <ArrowDownLeft size={16} />}
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
          <StatusIcon status={item.status} />
        </span>
        <span className="launch-actions-cell" onClick={(e) => e.stopPropagation()}>
          <RowActions actions={[
            { key: 'edit', label: 'Editar', icon: <Pencil size={15} />, onClick: onEdit },
            ...(item.status === 'open'
              ? [{ key: 'settle', label: item.type === 'income' ? 'Marcar como recebido' : 'Marcar como pago', icon: <CheckCircle2 size={15} />, onClick: onSettle }]
              : [{ key: 'pending', label: 'Marcar como pendente', icon: <RotateCcw size={15} />, onClick: onMarkPending }]),
            { key: 'delete', label: 'Excluir', icon: <Trash2 size={15} />, onClick: onDelete, danger: true },
          ]} />
        </span>
      </article>
    </div>
  );
}

function StatusIcon({ status }: { status: TransactionStatus }) {
  const settled = status === 'settled';
  const label = settled ? 'Efetivado' : 'Pendente';
  return (
    <span className={`status-icon status-icon--${status}`} title={label}>
      {settled ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
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

function DashboardPage({ transactions, goals, referenceDate, onChangeDate, onNavigate }: {
  transactions: Transaction[];
  goals: Goal[];
  referenceDate: Date;
  onChangeDate: (date: Date) => void;
  onNavigate: (page: AppPage) => void;
}) {
  const monthKeyValue = monthKey(referenceDate);
  const monthTransactions = transactions.filter((item) => item.dueDate.slice(0, 7) === monthKeyValue);
  const income = monthTransactions.filter((item) => item.type === 'income').reduce((sum, item) => sum + item.amount, 0);
  const expense = monthTransactions.filter((item) => item.type === 'expense').reduce((sum, item) => sum + item.amount, 0);
  const pendingIncome = monthTransactions.filter((item) => item.type === 'income' && item.status === 'open').reduce((sum, item) => sum + item.amount, 0);
  const pendingExpense = monthTransactions.filter((item) => item.type === 'expense' && item.status === 'open').reduce((sum, item) => sum + item.amount, 0);
  const projectedBalance = income - expense;

  const dashboardWindow = Array.from({ length: 6 }, (_, index) => new Date(referenceDate.getFullYear(), referenceDate.getMonth() - 1 + index, 1));
  const monthlySeries = dashboardWindow.map((date) => {
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

  const balanceValues = monthlySeries.map((item) => item.Saldo);
  const balanceMin = Math.min(...balanceValues, 0);
  const balanceMax = Math.max(...balanceValues, 0);
  const balanceRange = balanceMax - balanceMin;
  const balanceZeroOffset = balanceRange === 0 ? 0 : (balanceMax / balanceRange) * 100;
  const balanceColor = (value: number) => (value < 0 ? '#dc2626' : '#0284c7');
  const renderBalanceLabel = ({ x, y, value }: BalanceLabelProps) => {
    const balance = Number(value ?? 0);
    const labelX = Number(x);
    const labelY = Number(y);

    if (Number.isNaN(labelX) || Number.isNaN(labelY)) return <g />;

    return (
      <text x={labelX} y={labelY + (balance < 0 ? 20 : -8)} textAnchor="middle" fill={balanceColor(balance)} fontSize={11} fontWeight={700}>
        {formatCurrency(balance)}
      </text>
    );
  };
  const renderBalanceDot = ({ cx, cy, payload }: BalanceDotProps) => {
    const dotX = Number(cx);
    const dotY = Number(cy);
    const balance = Number(payload?.Saldo ?? 0);

    if (Number.isNaN(dotX) || Number.isNaN(dotY)) return null;

    return <circle cx={dotX} cy={dotY} r={4} fill="#ffffff" stroke={balanceColor(balance)} strokeWidth={2} />;
  };

  const goalSeries = goals
    .map((goal) => {
      const saved = goalSaved(goal);
      const percent = goal.targetAmount > 0 ? Math.min(100, Math.round((saved / goal.targetAmount) * 100)) : 0;

      return {
        id: goal.id,
        name: goal.name,
        saved,
        target: goal.targetAmount,
        percent,
        color: goal.color,
      };
    })
    .sort((a, b) => b.target - a.target)
    .slice(0, 6);

const hasData = transactions.length > 0;
  const compact = (value: number) => (Math.abs(value) >= 1000 ? `${Math.round(value / 1000)}k` : String(value));
  const settledTotal = monthTransactions.filter((item) => item.status === 'settled').reduce((sum, item) => sum + item.amount, 0);
  const openTotal = pendingIncome + pendingExpense;
  const monthTotal = income + expense;
  const settledPercent = monthTotal > 0 ? Math.round((settledTotal / monthTotal) * 100) : 0;
  const expenseCategories = Object.values(monthTransactions.filter((item) => item.type === 'expense').reduce<Record<string, { category: string; total: number; count: number }>>((acc, item) => {
    const entry = acc[item.category] ?? { category: item.category, total: 0, count: 0 };
    entry.total += item.amount;
    entry.count += 1;
    acc[item.category] = entry;
    return acc;
  }, {})).sort((a, b) => b.total - a.total).slice(0, 5);
  const incomeCategories = Object.values(monthTransactions.filter((item) => item.type === 'income').reduce<Record<string, { category: string; total: number; count: number }>>((acc, item) => {
    const entry = acc[item.category] ?? { category: item.category, total: 0, count: 0 };
    entry.total += item.amount;
    entry.count += 1;
    acc[item.category] = entry;
    return acc;
  }, {})).sort((a, b) => b.total - a.total).slice(0, 4);
  const highestExpense = expenseCategories[0]?.total ?? 0;
  const highestIncome = incomeCategories[0]?.total ?? 0;

  return (
    <>
      <section className="page-header page-header-split">
        <div className="page-header-left">
          <h1 className="page-title">Visão Geral</h1>
        </div>
        <div className="page-header-center">
          <MonthNavigator date={referenceDate} onChange={onChangeDate} />
        </div>
        <div className="page-header-actions"></div>
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
              <div className="chart-head"><strong>Receitas x Despesas</strong><span>Últimos 6 meses</span></div>
              <div className="chart-box chart-box--wide">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={monthlySeries} margin={{ top: 12, right: 18, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f6" />
                    <XAxis dataKey="month" tickLine={false} axisLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                    <YAxis tickLine={false} axisLine={false} width={44} tick={{ fontSize: 11, fill: '#94a3b8' }} tickFormatter={compact} />
                    <RechartsTooltip formatter={(value) => formatCurrency(Number(value))} cursor={{ fill: 'rgba(148,163,184,0.08)' }} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Bar dataKey="Receitas" fill="#1B99D8" radius={[6, 6, 0, 0]} maxBarSize={30} />
                    <Bar dataKey="Despesas" fill="#dc2626" radius={[6, 6, 0, 0]} maxBarSize={30} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </section>
          </div>

          <div className="dashboard-charts dashboard-charts--single">
            <section className="resource-panel chart-panel chart-panel--wide">
              <div className="chart-head"><strong>Saldo mensal</strong><span>Últimos 6 meses</span></div>
              <div className="chart-box chart-box--wide">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={monthlySeries} margin={{ top: 28, right: 18, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f6" />
                    <XAxis dataKey="month" tickLine={false} axisLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                    <YAxis tickLine={false} axisLine={false} width={44} tick={{ fontSize: 11, fill: '#94a3b8' }} tickFormatter={compact} />
                    <defs>
                      <linearGradient id="balanceLineGradient" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#0284c7" />
                        <stop offset={`${balanceZeroOffset}%`} stopColor="#0284c7" />
                        <stop offset={`${balanceZeroOffset}%`} stopColor="#dc2626" />
                        <stop offset="100%" stopColor="#dc2626" />
                      </linearGradient>
                    </defs>
                    <ReferenceLine y={0} stroke="#cbd5e1" strokeDasharray="4 4" />
                    <RechartsTooltip formatter={(value) => formatCurrency(Number(value))} cursor={{ stroke: '#bae6fd', strokeWidth: 1 }} />
                    <Line type="monotone" dataKey="Saldo" stroke="url(#balanceLineGradient)" strokeWidth={3} dot={renderBalanceDot} activeDot={{ r: 6 }}>
                      <LabelList dataKey="Saldo" content={renderBalanceLabel} />
                    </Line>
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </section>
          </div>

          <section className="dashboard-category-grid">
            <article className="resource-panel dashboard-category-panel dashboard-category-panel--expense">
              <div className="chart-head"><strong>Maiores despesas</strong><span>Categorias do mês</span></div>
              {expenseCategories.length ? (
                <div className="dashboard-category-list">
                  {expenseCategories.map((item) => (
                    <div className="dashboard-category-row" key={`expense-${item.category}`}>
                      <div>
                        <strong>{item.category}</strong>
                        <span>{item.count} {item.count === 1 ? 'lançamento' : 'lançamentos'}</span>
                      </div>
                      <div className="dashboard-category-meter"><span style={{ width: `${highestExpense > 0 ? Math.max(8, Math.round((item.total / highestExpense) * 100)) : 0}%` }} /></div>
                      <em>{formatCurrency(item.total)}</em>
                    </div>
                  ))}
                </div>
              ) : <div className="chart-empty">Sem despesas no mês</div>}
            </article>

            <article className="resource-panel dashboard-category-panel dashboard-category-panel--income">
              <div className="chart-head"><strong>Receitas por categoria</strong><span>Entradas do mês</span></div>
              {incomeCategories.length ? (
                <div className="dashboard-category-list">
                  {incomeCategories.map((item) => (
                    <div className="dashboard-category-row" key={`income-${item.category}`}>
                      <div>
                        <strong>{item.category}</strong>
                        <span>{item.count} {item.count === 1 ? 'lançamento' : 'lançamentos'}</span>
                      </div>
                      <div className="dashboard-category-meter"><span style={{ width: `${highestIncome > 0 ? Math.max(8, Math.round((item.total / highestIncome) * 100)) : 0}%` }} /></div>
                      <em>{formatCurrency(item.total)}</em>
                    </div>
                  ))}
                </div>
              ) : <div className="chart-empty">Sem receitas no mês</div>}
            </article>
          </section>

          <section className="resource-panel chart-panel goals-overview-panel">
            <div className="chart-head"><strong>Metas</strong><span>Progresso das metas</span></div>
            {goalSeries.length ? (
              <div className="goals-overview-grid">
                {goalSeries.map((goal) => (
                  <article key={goal.id} className="goal-overview-card">
                    <div className="goal-overview-head">
                      <strong>{goal.name}</strong>
                      <span style={{ color: goal.color }}>{goal.percent}%</span>
                    </div>
                    <div className="goal-overview-values">
                      <span>{formatCurrency(goal.saved)}</span>
                      <small>de {formatCurrency(goal.target)}</small>
                    </div>
                    <div className="goal-overview-track">
                      <span style={{ width: `${goal.percent}%`, background: `linear-gradient(90deg, ${goal.color}, ${goal.color}cc)` }} />
                    </div>
                  </article>
                ))}
              </div>
            ) : (
              <div className="chart-empty goals-overview-empty">Nenhuma meta criada</div>
            )}
          </section>


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
    </>
  );
}

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
                    <div className="goal-card-actions">
                      <button type="button" className="category-action-button" title="Editar meta" aria-label={`Editar ${goal.name}`} onClick={() => onEdit(goal)}><Pencil size={15} /></button>
                      <button type="button" className="category-action-button category-action-button--danger" title="Excluir meta" aria-label={`Excluir ${goal.name}`} onClick={() => onDelete(goal)}><Trash2 size={15} /></button>
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
                    <button type="button" className="goal-button goal-button--deposit" onClick={() => onDeposit(goal)}><Plus size={15} /> Depositar</button>
                    <button type="button" className="goal-button goal-button--withdraw" onClick={() => onWithdraw(goal)} disabled={saved <= 0}><Minus size={15} /> Resgatar</button>
                  </div>

                  {goal.movements.length > 0 ? (
                    <button type="button" className="goal-view-movements" onClick={() => setMovementsGoal(goal)}>
                      <ReceiptText size={14} />
                      Ver lançamentos ({goal.movements.length})
                    </button>
                  ) : null}
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
                  <p className="goal-modal-subtitle">{movementsGoal.movements.length} {movementsGoal.movements.length === 1 ? 'lançamento' : 'lançamentos'} · Guardado: {formatCurrency(goalSaved(movementsGoal))}</p>
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
  const [error, setError] = useState('');

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
            <div className="form-field form-field-full">
              <span className="form-label form-label-optional">Cor</span>
              <div className="account-color-options" role="group" aria-label="Cor da meta">
                {ACCOUNT_COLORS.map((option) => <button type="button" key={option} className={`account-color-swatch${color === option ? ' active' : ''}`} style={{ backgroundColor: option }} onClick={() => setColor(option)} aria-label={`Selecionar cor ${option}`} aria-pressed={color === option} />)}
              </div>
            </div>
            <div className="form-field form-field-full">
              <span className="form-label form-label-optional">Ícone</span>
              <div className="account-icon-options" role="group" aria-label="Ícone da meta">
                {CATEGORY_ICON_OPTIONS.map((option) => <button type="button" key={option.key} className={`account-icon-option${icon === option.key ? ' active' : ''}`} style={icon === option.key ? { borderColor: color, color, backgroundColor: `${color}12` } : undefined} onClick={() => setIcon(option.key)} title={option.label} aria-label={option.label} aria-pressed={icon === option.key}><CategoryIconGraphic icon={option.key} size={18} /></button>)}
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

function HelpPage() {
  const sections = [
    {
      title: '1. Como o sistema funciona',
      intro: 'O Caesar Finance organiza sua vida financeira por mês. A ideia é cadastrar a base correta, lançar receitas e despesas, acompanhar o que está em aberto e conferir tudo pelos relatórios.',
      steps: [
        'Use a Visão Geral para acompanhar o resumo do mês: receita prevista, despesa prevista, saldo, valores a receber, valores a pagar e evolução dos lançamentos.',
        'Use Transações para registrar tudo que entra e sai. É nessa aba que você controla vencimento, categoria, recorrência, status e conta vinculada.',
        'Use Contas para separar onde o dinheiro está: carteira, conta corrente, poupança, investimento ou outra conta cadastrada.',
        'Use Categorias para organizar o motivo de cada lançamento, como salário, moradia, mercado, transporte, investimentos e outros grupos.',
        'Use Metas para acompanhar objetivos financeiros separados dos lançamentos comuns, como reserva, carro, viagem ou compra planejada.',
        'Use Relatórios para conferir o fechamento do mês e exportar os dados quando precisar salvar ou enviar as informações.',
      ],
    },
    {
      title: '2. Como cadastrar as contas',
      intro: 'Cadastre primeiro as contas que você usa no dia a dia. Isso ajuda o sistema a calcular o saldo real quando uma transação for efetivada.',
      steps: [
        'Entre na aba Contas e clique no botão de adicionar.',
        'Informe o nome da conta, por exemplo Carteira, Nubank, Banco principal, Poupança ou Investimentos.',
        'Escolha o tipo da conta para deixar a lista organizada e fácil de filtrar depois.',
        'Informe o saldo inicial. Esse valor deve ser o saldo que existia antes de começar a controlar as movimentações no sistema.',
        'Depois de salvar, use essa conta nos lançamentos. Quando uma receita ou despesa for marcada como efetivada, o saldo atual será atualizado.',
      ],
    },
    {
      title: '3. Como cadastrar categorias',
      intro: 'As categorias servem para identificar de onde vem o dinheiro e para onde ele vai. Quanto melhor elas forem cadastradas, mais claros ficam os relatórios.',
      steps: [
        'Entre na aba Categorias e clique no botão de adicionar.',
        'Escolha se a categoria é de receita ou despesa.',
        'Dê um nome simples e direto, como Salário, Extra, Moradia, Alimentação, Transporte, Cartão ou Investimentos.',
        'Use cores e ícones apenas para facilitar a identificação visual, sem criar categorias repetidas com nomes parecidos.',
        'Ao lançar uma transação, selecione sempre a categoria correta para que os totais e filtros funcionem bem.',
      ],
    },
    {
      title: '4. Como lançar receitas e despesas',
      intro: 'Todo valor que você quer controlar deve virar uma transação. Pode ser uma entrada, como salário, ou uma saída, como aluguel, mercado ou cartão.',
      steps: [
        'Entre na aba Transações e clique no botão de adicionar.',
        'Escolha o tipo do lançamento: receita para dinheiro que entra, despesa para dinheiro que sai.',
        'Preencha a descrição com um nome fácil de entender, como Salário mensal, Aluguel do apartamento ou Supermercado do mês.',
        'Selecione a categoria correta, informe o valor e escolha a data de vencimento.',
        'Escolha a recorrência: única para lançamentos pontuais, fixa para valores que se repetem todo mês e parcelada quando houver número definido de parcelas.',
        'Se já souber em qual conta o valor será pago ou recebido, selecione a conta no lançamento.',
        'Salve o lançamento. Ele ficará em aberto até você marcar como efetivado.',
      ],
    },
    {
      title: '5. Como efetivar pagamentos e recebimentos',
      intro: 'Efetivar uma transação significa confirmar que aquele dinheiro realmente entrou ou saiu da conta.',
      steps: [
        'Quando uma receita cair na conta, abra a transação e marque como efetivada.',
        'Quando uma despesa for paga, marque também como efetivada.',
        'Confira o valor real antes de confirmar. Se o valor pago ou recebido for diferente do previsto, ajuste o valor real.',
        'Selecione a conta correta para que o saldo atual seja atualizado no lugar certo.',
        'Depois de efetivar, o sistema passa a considerar essa movimentação no saldo real da conta e no status dos relatórios.',
      ],
    },
    {
      title: '6. Como acompanhar o mês e usar relatórios',
      intro: 'Depois dos cadastros e lançamentos, o acompanhamento fica nas consultas, filtros e relatórios.',
      steps: [
        'Use os filtros para encontrar lançamentos por data, categoria, tipo, conta ou status.',
        'Na Visão Geral, confira rapidamente se o mês está positivo ou negativo e veja o que ainda falta receber ou pagar.',
        'Na aba Transações, acompanhe os vencimentos e resolva os itens pendentes conforme eles forem acontecendo.',
        'Na aba Relatórios, compare receita, despesa e saldo final do mês.',
        'Antes de exportar PDF ou Excel, confira se os lançamentos importantes estão cadastrados e se os pagamentos já foram efetivados.',
      ],
    },
  ];

  return (
    <section className="help-page">
      <div className="page-header help-header">
        <div className="page-header-left">
          <h1 className="page-title help-title">Central de ajuda</h1>
        </div>
      </div>

      <div className="help-grid">
        {sections.map((section) => (
          <article className="help-card" key={section.title}>
            <h2>{section.title}</h2>
            <p className="help-card-intro">{section.intro}</p>
            <ol className="help-step-list">
              {section.steps.map((item, index) => (
                <li key={item}>
                  <span className="help-step-number">{index + 1}</span>
                  <span>{item}</span>
                </li>
              ))}
            </ol>
          </article>
        ))}
      </div>
    </section>
  );
}
function Topbar({ activePage, onNavigate, onLogout, onOpenTransactionFilters, onImportTransactions, onOpenAccountFilters, accountFilterOpen, accountActiveFilterCount, onOpenCategoryFilters, categoryFilterOpen, categoryActiveFilterCount, onOpenGoalFilters, goalFilterOpen, goalActiveFilterCount }: {
  activePage: AppPage;
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
    function closeOnOutside(event: PointerEvent) { if (!profileRef.current?.contains(event.target as Node)) setProfileOpen(false); }
    function closeOnEscape(event: KeyboardEvent) { if (event.key === 'Escape') setProfileOpen(false); }
    document.addEventListener('pointerdown', closeOnOutside);
    document.addEventListener('keydown', closeOnEscape);
    return () => { document.removeEventListener('pointerdown', closeOnOutside); document.removeEventListener('keydown', closeOnEscape); };
  }, [profileOpen]);

  const profileMenu = (
    <div className="profile-menu" ref={profileRef}>
      <button type="button" className={`user-profile${profileOpen ? ' active' : ''}`} onClick={() => setProfileOpen((open) => !open)} aria-expanded={profileOpen} aria-haspopup="menu" aria-label="Abrir perfil">
        <span className="avatar"><UserCircle size={18} /></span>
        <span className="user-name">Taylor Felipe Barbosa</span>
        <ChevronRight className="profile-chevron" size={15} />
      </button>
      {profileOpen ? (
        <div className="profile-dropdown" role="menu">
          <div className="profile-dropdown-user"><strong>Taylor Felipe Barbosa</strong><span>Perfil do usuário</span></div>
          <button type="button" className="profile-menu-item" role="menuitem" onClick={() => { setProfileOpen(false); onNavigate('help'); }}><HelpCircle size={16} /> Ajuda</button>
          <button type="button" className="profile-logout" role="menuitem" onClick={() => { setProfileOpen(false); onLogout(); }}><LogOut size={16} /> Sair</button>
        </div>
      ) : null}
    </div>
  );

  return (
    <header className={`topbar topbar--${activePage}`}>
      <div className="topbar-desktop-shell">
        <div className="topbar-left">
          <button type="button" className="brand brand-btn" aria-label="Ir para transações" onClick={() => onNavigate('transactions')}>
            <svg width="28" height="28" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="17" y="50" width="15" height="35" rx="4" fill="var(--primary)" /><rect x="43" y="30" width="15" height="55" rx="4" fill="var(--primary)" /><rect x="69" y="15" width="15" height="70" rx="4" fill="var(--primary)" /></svg>
          </button>
          <div className="topbar-separator" />
          <nav className="breadcrumbs" aria-label="Breadcrumb"><Home className="home-icon" /><ChevronRight size={13} className="breadcrumb-chevron" /><span>Caesar Finance</span><ChevronRight size={13} className="breadcrumb-chevron" /><span className="breadcrumb-active">{PAGE_LABELS[activePage]}</span></nav>
        </div>
        <div className="topbar-right">
          {activePage === 'transactions' ? (
            <div className="topbar-actions-menu" ref={actionsRef}>
              <button type="button" className="topbar-more-button" aria-label="Mais opções" aria-expanded={actionsOpen} aria-haspopup="menu" onClick={() => setActionsOpen((open) => !open)}><MoreVertical size={20} /></button>
              {actionsOpen ? <div className="topbar-actions-popover" role="menu"><button type="button" role="menuitem" onClick={() => { setActionsOpen(false); onOpenTransactionFilters(); }}><ListFilter size={16} />Filtros</button><button type="button" role="menuitem" onClick={() => { setActionsOpen(false); onImportTransactions(); }}><Upload size={16} />Importar</button></div> : null}
            </div>
          ) : activePage === 'accounts' ? (
            <button type="button" className={`topbar-account-filter-button${accountFilterOpen ? ' active' : ''}`} data-account-filter-trigger aria-label="Filtrar contas" aria-expanded={accountFilterOpen} aria-haspopup="dialog" onClick={onOpenAccountFilters}><ListFilter size={19} />{accountActiveFilterCount > 0 ? <span className="filter-badge">{accountActiveFilterCount}</span> : null}</button>
          ) : activePage === 'categories' ? (
            <button type="button" className={`topbar-category-filter-button${categoryFilterOpen ? ' active' : ''}`} data-category-filter-trigger aria-label="Filtrar categorias" aria-expanded={categoryFilterOpen} aria-haspopup="dialog" onClick={onOpenCategoryFilters}><ListFilter size={19} />{categoryActiveFilterCount > 0 ? <span className="filter-badge">{categoryActiveFilterCount}</span> : null}</button>
          ) : activePage === 'goals' ? (
            <button type="button" className={`topbar-goal-filter-button${goalFilterOpen ? ' active' : ''}`} data-goal-filter-trigger aria-label="Filtrar metas" aria-expanded={goalFilterOpen} aria-haspopup="dialog" onClick={onOpenGoalFilters}><ListFilter size={19} />{goalActiveFilterCount > 0 ? <span className="filter-badge">{goalActiveFilterCount}</span> : null}</button>
          ) : null}
          {profileMenu}
        </div>
      </div>
    </header>
  );
}
function Sidebar({ activePage, onNavigate }: { activePage: AppPage; onNavigate: (page: AppPage) => void }) {
  const items = [
    { label: 'Visão Geral', icon: LayoutGrid, page: 'dashboard' as const },
    { label: 'Transações', icon: ReceiptText, page: 'transactions' as const },
    { label: 'Categorias', icon: Tags, page: 'categories' as const },
    { label: 'Metas', icon: Target, page: 'goals' as const },
    { label: 'Relatórios', icon: FileText, page: 'reports' as const },
    { label: 'Contas', icon: CreditCard, page: 'accounts' as const },
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
              <span className="category-report-name">{category.name}</span>
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

function CategoryModal({ items, category, onClose, onSave }: { items: CategoryItem[]; category: CategoryItem | null; onClose: () => void; onSave: (category: CustomCategory) => void }) {
  const [name, setName] = useState(category?.name ?? '');
  const [kind, setKind] = useState<CategoryKind>(category?.kind ?? 'expense');
  const [color, setColor] = useState(category?.color ?? ACCOUNT_COLORS[0] ?? '#1B99D8');
  const [error, setError] = useState('');

  function submit(event: FormEvent) {
    event.preventDefault();
    const normalizedName = name.trim();
    if (!normalizedName) return;
    if (items.some((item) => item.id !== category?.id && item.kind === kind && item.name.toLowerCase() === normalizedName.toLowerCase())) {
      setError('Essa categoria já existe nesse grupo.');
      return;
    }
    const icon = category?.icon ?? defaultCategoryIcon(normalizedName, kind);
    onSave({ id: category?.id ?? crypto.randomUUID(), name: normalizedName, kind, color, icon });
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card settle-modal" onSubmit={submit} onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only">{category ? 'Editar categoria' : 'Nova categoria'}</h2>
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
            <div className="form-field form-field-full">
              <span className="form-label form-label-optional">Cor</span>
              <div className="account-color-options" role="group" aria-label="Cor da categoria">
                {ACCOUNT_COLORS.map((option) => (
                  <button key={option} type="button" className={`account-color-swatch${color === option ? ' active' : ''}`} style={{ backgroundColor: option }} onClick={() => setColor(option)} aria-label={`Selecionar cor ${option}`} aria-pressed={color === option} />
                ))}
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
function EditTransactionModal({ item, incomeCategories, expenseCategories, onClose, onSave }: { item: Transaction; incomeCategories: string[]; expenseCategories: string[]; onClose: () => void; onSave: (updated: Transaction) => void }) {
  const [type, setType] = useState<TransactionType>(item.type);
  const [description, setDescription] = useState(item.description);
  const [category, setCategory] = useState(item.category);
  const [amount, setAmount] = useState(() => formatCurrencyInput(item.amount));
  const [dueDate, setDueDate] = useState(item.dueDate);
  const [recurrence, setRecurrence] = useState<RecurrenceType>(item.recurrence);
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
    onSave({ ...item, type, description: normalizedDescription, category, amount: value, dueDate, recurrence });
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
            <div className="form-field form-field-full">
              <span className="form-label form-label-optional">Cor</span>
              <div className="account-color-options" role="group" aria-label="Cor da conta">
                {ACCOUNT_COLORS.map((option) => (
                  <button
                    key={option}
                    type="button"
                    className={`account-color-swatch${color === option ? ' active' : ''}`}
                    style={{ backgroundColor: option }}
                    onClick={() => setColor(option)}
                    aria-label={`Selecionar cor ${option}`}
                    aria-pressed={color === option}
                  />
                ))}
              </div>
            </div>
            <div className="form-field form-field-full">
              <span className="form-label form-label-optional">Ícone</span>
              <div className="account-icon-options" role="group" aria-label="Ícone da conta">
                {ACCOUNT_ICON_OPTIONS.map((option) => (
                  <button
                    key={option.key}
                    type="button"
                    className={`account-icon-option${icon === option.key ? ' active' : ''}`}
                    style={icon === option.key ? { borderColor: color, color, backgroundColor: `${color}12` } : undefined}
                    onClick={() => setIcon(option.key)}
                    title={option.label}
                    aria-label={option.label}
                    aria-pressed={icon === option.key}
                  >
                    <AccountIconGraphic icon={option.key} size={18} />
                  </button>
                ))}
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
  return (
    <div className="transaction-summary-stats">
      <MetricCard label="Receita" value={formatCurrency(summary.income)} tone="info" icon={<TrendingUp size={18} />} />
      <MetricCard label="Despesas" value={formatCurrency(summary.expense)} tone="danger" icon={<TrendingDown size={18} />} />
      <MetricCard
        label="Saldo"
        value={formatCurrency(summary.balance)}
        tone={summary.balance >= 0 ? 'info' : 'danger'}
        icon={<CircleDollarSign size={18} />}
      />
    </div>
  );
}
function MetricCard({ label, value, icon, tone }: { label: string; value: string; icon: ReactNode; tone: string }) {
  return (
    <article className={`signature-stat-card signature-stat-card--${tone}`}>
      <span className="signature-stat-icon">{icon}</span>
      <span className="signature-stat-copy">
        <span className="signature-stat-label">{label}</span>
        <strong className="signature-stat-value">{value}</strong>
      </span>
    </article>
  );
}

function LaunchModal({ incomeCategories, expenseCategories, onClose, onCreate }: { incomeCategories: string[]; expenseCategories: string[]; onClose: () => void; onCreate: (items: Transaction[]) => void }) {
  const [type, setType] = useState<TransactionType>('expense');
  const [form, setForm] = useState<LaunchForm>(() => ({ ...initialForm, category: expenseCategories[0] ?? '' }));

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
    onCreate(generateTransactions(form, type));
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
function ImportTransactionsModal({ onClose, onImport }: { onClose: () => void; onImport: (items: Transaction[]) => void }) {
  const [fileName, setFileName] = useState('');
  const [rows, setRows] = useState<ImportPreviewRow[]>([]);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const validRows = rows.filter((row) => !row.error);
  const importedTransactions = validRows.flatMap((row) => row.transactions);
  const invalidCount = rows.length - validRows.length;

  async function handleFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    setFileName(file.name);
    setRows([]);
    setError('');

    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      setError('Selecione um arquivo Excel no formato .xlsx.');
      return;
    }

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

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card import-modal" role="dialog" aria-modal="true" aria-labelledby="import-modal-title" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2 className="modal-title modal-title-only" id="import-modal-title">Importar transações</h2>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>

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
            </div>
            <button type="button" className="import-template-button" onClick={downloadImportTemplate}>
              <FileSpreadsheet size={16} /> Baixar modelo
            </button>
          </div>

          {error ? <div className="import-alert import-alert--error"><AlertCircle size={17} /><span>{error}</span></div> : null}

          {rows.length ? (
            <>
              <div className="import-summary">
                <span><strong>{validRows.length}</strong> linhas válidas</span>
                <span><strong>{importedTransactions.length}</strong> transações serão criadas</span>
                {invalidCount ? <span className="has-error"><strong>{invalidCount}</strong> com erro</span> : null}
              </div>

              <div className="import-preview">
                <div className="import-preview-head">
                  <span>Linha</span><span>Tipo</span><span>Descrição</span><span>Categoria</span><span>Vencimento</span><span>Valor</span><span>Status</span>
                </div>
                {rows.slice(0, 8).map((row) => (
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
                {rows.length > 8 ? <div className="import-preview-more">Mais {rows.length - 8} linhas na planilha</div> : null}
              </div>
            </>
          ) : null}
        </div>

        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="button" className="button-primary" disabled={!importedTransactions.length || loading} onClick={() => onImport(importedTransactions)}>
            Importar {importedTransactions.length || ''}
          </button>
        </div>
      </div>
    </div>
  );
}

function SettleModal({ item, accounts, onClose, onSave }: { item: Transaction; accounts: Account[]; onClose: () => void; onSave: (accountId: string, settledAt: string, settledAmount: number) => void }) {
  const [accountId, setAccountId] = useState(() => accounts.find((account) => account.id === DEFAULT_ACCOUNT.id)?.id ?? accounts[0]?.id ?? '');
  const [settledAmount, setSettledAmount] = useState(() => formatCurrencyInput(item.settledAmount ?? item.amount));
  const [settledAt, setSettledAt] = useState(new Date().toISOString().slice(0, 10));

  function submit(event: FormEvent) {
    event.preventDefault();
    const amount = parseAmount(settledAmount);
    if (!accountId || amount <= 0) return;
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
              <span className="form-label">Conta utilizada</span>
              <select className="form-input" value={accountId} onChange={(event) => setAccountId(event.target.value)}>
                {accounts.map((account) => <option key={account.id} value={account.id}>{account.name}</option>)}
              </select>
            </label>
            <label className="form-field form-field-full">
              <span className="form-label">Valor {item.type === 'income' ? 'recebido' : 'pago'}</span>
              <div className="form-input-wrap">
                <span className="form-input-prefix">R$</span>
                <input className="form-input" value={settledAmount} onChange={(event) => setSettledAmount(formatCurrencyInput(event.target.value))} inputMode="decimal" placeholder="0,00" />
              </div>
            </label>
            <label className="form-field form-field-full">
              <span className="form-label">Data da baixa</span>
              <input className="form-input" type="date" value={settledAt} onChange={(event) => setSettledAt(event.target.value)} />
            </label>
          </div>
        </div>
        <div className="modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary">Confirmar baixa</button>
        </div>
      </form>
    </div>
  );
}
