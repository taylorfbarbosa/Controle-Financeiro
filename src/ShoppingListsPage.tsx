import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent } from 'react';
import {
  CalendarDays,
  Check,
  ChevronLeft,
  CheckCircle2,
  Clock3,
  Copy,
  PackageOpen,
  Pencil,
  Plus,
  Search,
  ShoppingCart,
  Trash2,
  UserRound,
  Users,
  X,
  XCircle,
} from 'lucide-react';
import type { FriendUser } from './App';
import rubyDiamond from './assets/Diamante.png';
import { deleteShoppingList, loadShoppingLists, saveShoppingList, updateShoppingList } from './lib/db';
import './styles/shopping-lists.css';

type ShoppingListStatus = 'open' | 'finalized' | 'cancelled';

type ShoppingListItem = {
  id: string;
  name: string;
  quantity: number;
  unitPrice: number;
  addedById: string;
  createdAt: string;
  purchased?: boolean;
};

type ShoppingList = {
  id: string;
  name: string;
  date: string;
  note?: string;
  status: ShoppingListStatus;
  creatorId: string;
  participantIds: string[];
  items: ShoppingListItem[];
  createdAt: string;
  finalizedAt?: string;
  finalizedById?: string;
  cancelledAt?: string;
};

type ItemDraft = { name: string; quantity: string; unitPrice: string };

const emptyItem: ItemDraft = { name: '', quantity: '1', unitPrice: '' };

function money(value: number) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
}

function dateLabel(value?: string) {
  if (!value) return '—';
  const normalized = value.length === 10 ? `${value}T12:00:00` : value;
  return new Intl.DateTimeFormat('pt-BR').format(new Date(normalized));
}

const AVATAR_COLORS = [
  { bg: '#1B99D8', text: '#fff' },
  { bg: '#10b981', text: '#fff' },
  { bg: '#f59e0b', text: '#fff' },
  { bg: '#8b5cf6', text: '#fff' },
  { bg: '#ef4444', text: '#fff' },
  { bg: '#06b6d4', text: '#fff' },
  { bg: '#ec4899', text: '#fff' },
  { bg: '#84cc16', text: '#fff' },
];

function initials(name: string) {
  return name.trim().split(/\s+/).slice(0, 2).map((part) => part[0]).join('').toUpperCase() || 'U';
}

function avatarColor(name: string) {
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
  return AVATAR_COLORS[Math.abs(hash) % AVATAR_COLORS.length];
}

function normalizeProduct(value: string) {
  return value.trim().toLocaleLowerCase('pt-BR');
}

function parseDecimal(value: string) {
  const normalized = value.replace(/\s/g, '').replace(/\.(?=\d{3}(?:\D|$))/g, '').replace(',', '.');
  return Number(normalized) || 0;
}

function listTotal(list: ShoppingList) {
  return list.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
}

function participantAvatar(user: FriendUser, small = false) {
  const cls = `shopping-avatar${small ? ' shopping-avatar--small' : ''}`;
  if (user.avatarUrl) return <img className={cls} src={user.avatarUrl} alt={user.name} />;
  const { bg, text } = avatarColor(user.name);
  return <span className={`${cls} shopping-avatar--initials`} style={{ background: bg, color: text }} aria-label={user.name}>{initials(user.name)}</span>;
}

export function ShoppingListsPage({
  currentUser,
  friends,
  onListItemAdded,
  openCreateSignal,
  openItemSignal,
  onDetailChange,
}: {
  currentUser: FriendUser;
  friends: FriendUser[];
  onListItemAdded?: (event: { listId: string; listName: string; itemName: string; participantIds: string[] }) => void;
  openCreateSignal?: number;
  openItemSignal?: number;
  onDetailChange?: (name: string | null, canAdd?: boolean) => void;
}) {
  const [lists, setLists] = useState<ShoppingList[]>([]);
  const [listsLoading, setListsLoading] = useState(true);
  const [search, setSearch] = useState('');
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [createOpen, setCreateOpen] = useState(false);
  const [editingList, setEditingList] = useState<ShoppingList | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<ShoppingList | null>(null);
  const [actionTarget, setActionTarget] = useState<{ list: ShoppingList; action: 'finalize' | 'cancel' } | null>(null);

  const fetchLists = useCallback(() => {
    setListsLoading(true);
    loadShoppingLists<ShoppingList>()
      .then((data) => setLists(data))
      .catch((err) => console.warn('Falha ao carregar listas:', err))
      .finally(() => setListsLoading(false));
  }, []);

  useEffect(() => { fetchLists(); }, [fetchLists]);

  const lastProcessedSignalRef = useRef(openCreateSignal);
  useEffect(() => {
    if (openCreateSignal && openCreateSignal !== lastProcessedSignalRef.current) {
      setCreateOpen(true);
      lastProcessedSignalRef.current = openCreateSignal;
    }
  }, [openCreateSignal]);


  const visibleLists = useMemo(() => lists.filter((list) => list.participantIds.includes(currentUser.id)), [currentUser.id, lists]);
  const selected = visibleLists.find((list) => list.id === selectedId) ?? null;

  useEffect(() => {
    onDetailChange?.(selected ? selected.name : null, selected?.status === 'open');
  }, [selected, onDetailChange]);

  function updateList(nextList: ShoppingList) {
    setLists((prev) => prev.map((list) => list.id === nextList.id ? nextList : list));
    updateShoppingList<ShoppingList>(nextList).then((saved) => {
      setLists((prev) => prev.map((list) => list.id === saved.id ? saved : list));
    }).catch((err) => console.warn('Falha ao atualizar lista:', err));
  }

  function deleteList(id: string) {
    setLists((prev) => prev.filter((list) => list.id !== id));
    setDeleteTarget(null);
    deleteShoppingList(id).catch((err) => console.warn('Falha ao excluir lista:', err));
  }

  function createList(form: { name: string; date: string; note: string; friendIds: string[] }) {
    const list: ShoppingList = {
      id: crypto.randomUUID(),
      name: form.name.trim(),
      date: form.date,
      note: form.note.trim() || undefined,
      status: 'open',
      creatorId: currentUser.id,
      participantIds: [currentUser.id, ...form.friendIds],
      items: [],
      createdAt: new Date().toISOString(),
    };
    setLists((prev) => [list, ...prev]);
    setCreateOpen(false);
    setSelectedId(list.id);
    saveShoppingList<ShoppingList>(list).then((saved) => {
      setLists((prev) => prev.map((l) => l.id === list.id ? saved : l));
    }).catch((err) => console.warn('Falha ao criar lista:', err));
  }

  function saveEditedList(form: { name: string; date: string; note: string }) {
    if (!editingList) return;
    updateList({ ...editingList, name: form.name.trim(), date: form.date, note: form.note.trim() || undefined });
    setEditingList(null);
  }

  function duplicateList(source: ShoppingList) {
    const allowedFriendIds = new Set(friends.map((friend) => friend.id));
    const list: ShoppingList = {
      ...source,
      id: crypto.randomUUID(),
      name: `${source.name} (cópia)`,
      date: new Date().toISOString().slice(0, 10),
      status: 'open',
      creatorId: currentUser.id,
      participantIds: [currentUser.id, ...source.participantIds.filter((id) => id !== currentUser.id && allowedFriendIds.has(id))],
      items: source.items.map((item) => ({ ...item, id: crypto.randomUUID(), addedById: currentUser.id, createdAt: new Date().toISOString() })),
      createdAt: new Date().toISOString(),
      finalizedAt: undefined,
      finalizedById: undefined,
      cancelledAt: undefined,
    };
    setLists((prev) => [list, ...prev]);
    setSelectedId(list.id);
    saveShoppingList<ShoppingList>(list).then((saved) => {
      setLists((prev) => prev.map((l) => l.id === list.id ? saved : l));
    }).catch((err) => console.warn('Falha ao duplicar lista:', err));
  }

  function completeAction() {
    if (!actionTarget || actionTarget.list.status !== 'open') return;
    const { list, action } = actionTarget;
    const now = new Date().toISOString();
    updateList(action === 'finalize'
      ? { ...list, status: 'finalized', finalizedAt: now, finalizedById: currentUser.id }
      : { ...list, status: 'cancelled', cancelledAt: now });
    setActionTarget(null);
  }

  if (selected) {
    return (
      <ShoppingListDetail
        list={selected}
        currentUser={currentUser}
        friends={friends}
        allLists={visibleLists}
        onBack={() => setSelectedId(null)}
        onChange={updateList}
        onListItemAdded={onListItemAdded}
        openItemSignal={openItemSignal}
      />
    );
  }

  const query = normalizeProduct(search);
  const filtered = visibleLists
    .filter((list) => !query || normalizeProduct(list.name).includes(query) || list.items.some((item) => normalizeProduct(item.name).includes(query)))
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  return (
    <div className="shopping-page">
      <section className="page-header page-header-split shopping-page-header">
        <div className="page-header-left">
          <h1 className="page-title">Listas de compras</h1>
          <p className="page-subtitle">Planeje compras em conjunto e acompanhe o total em tempo real.</p>
        </div>
        <div className="page-header-actions">
          <button type="button" className="page-primary-action" onClick={() => setCreateOpen(true)}><Plus size={17} /> Nova lista</button>
        </div>
      </section>

      <section className="shopping-toolbar" aria-label="Controles das listas">
        <label className="shopping-search"><Search size={17} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Buscar lista ou produto" aria-label="Buscar lista ou produto" /></label>
      </section>

      {listsLoading ? (
        <section className="shopping-empty"><span className="shopping-loading-spinner" /></section>
      ) : filtered.length ? (
        <section className="shopping-card-grid">
          {filtered.map((list) => (
            <ShoppingListCard
              key={list.id}
              list={list}
              currentUser={currentUser}
              friends={friends}
              onOpen={() => setSelectedId(list.id)}
              onEdit={() => setEditingList(list)}
              onDelete={() => setDeleteTarget(list)}
              onDuplicate={() => duplicateList(list)}
              onFinalize={() => setActionTarget({ list, action: 'finalize' })}
              onCancel={() => setActionTarget({ list, action: 'cancel' })}
            />
          ))}
        </section>
      ) : (
        <section className="shopping-empty">
          <span><ShoppingCart size={27} /></span>
          <h2>{search ? 'Nenhuma lista encontrada' : 'Sua próxima compra começa aqui'}</h2>
          <p>{search ? 'Tente buscar por outro nome ou produto.' : 'Crie uma lista, convide seus amigos e organizem os itens juntos.'}</p>
          {!search ? <button type="button" className="button-primary" onClick={() => setCreateOpen(true)}><Plus size={16} /> Criar primeira lista</button> : null}
        </section>
      )}

      {createOpen ? <CreateShoppingListModal friends={friends} onClose={() => setCreateOpen(false)} onCreate={createList} /> : null}
      {editingList ? <EditShoppingListModal list={editingList} onClose={() => setEditingList(null)} onSave={saveEditedList} /> : null}
      {actionTarget ? <ConfirmListAction action={actionTarget.action} list={actionTarget.list} onClose={() => setActionTarget(null)} onConfirm={completeAction} /> : null}
      {deleteTarget ? (
        <div className="modal-overlay">
          <div className="modal-card shopping-confirm-modal">
            <div className="shopping-confirm-icon danger"><Trash2 size={24} /></div>
            <h2>Apagar esta lista?</h2>
            <p>A lista <strong>{deleteTarget.name}</strong> e todos os seus itens serão removidos permanentemente.</p>
            <div>
              <button type="button" className="button-secondary" onClick={() => setDeleteTarget(null)}>Cancelar</button>
              <button type="button" className="button-danger" onClick={() => deleteList(deleteTarget.id)}>Apagar lista</button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function ShoppingListCard({ list, currentUser, friends, onOpen, onEdit, onDelete, onDuplicate, onFinalize, onCancel }: { list: ShoppingList; currentUser: FriendUser; friends: FriendUser[]; onOpen: () => void; onEdit: () => void; onDelete: () => void; onDuplicate: () => void; onFinalize: () => void; onCancel: () => void }) {
  const people = [currentUser, ...friends].filter((user) => list.participantIds.includes(user.id));
  const total = listTotal(list);
  const totalItemsCount = list.items.length;
  return (
    <article className="shopping-list-card">
      <div className="shopping-card-actions-row">
        <span className={`shopping-status shopping-status--${list.status}`}>
          {list.status === 'open' ? <Clock3 size={13} /> : list.status === 'finalized' ? <CheckCircle2 size={13} /> : <XCircle size={13} />}
          {list.status === 'open' ? 'Aberta' : list.status === 'finalized' ? 'Finalizada' : 'Cancelada'}
        </span>
        <div className="shopping-card-icon-actions">
          <button type="button" className="shopping-card-icon-btn" onClick={onEdit} aria-label="Editar lista"><Pencil size={15} /></button>
          <button type="button" className="shopping-card-icon-btn shopping-card-icon-btn--danger" onClick={onDelete} aria-label="Apagar lista"><Trash2 size={15} /></button>
        </div>
      </div>
      <button type="button" className="shopping-card-main" onClick={onOpen} aria-label={`Abrir ${list.name}`}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '16px', width: '100%' }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', textAlign: 'left', minWidth: 0 }}>
            <h2 className="shopping-list-title-text" style={{ margin: 0, fontSize: '17px', fontWeight: 700, color: 'var(--text-primary)', textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap' }}>{list.name}</h2>
            <span className="shopping-card-date" style={{ color: '#82909a', fontSize: '12px' }}><CalendarDays size={13} /> {dateLabel(list.date)}</span>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', textAlign: 'right', flexShrink: 0 }}>
            <strong style={{ fontSize: '17px', fontWeight: 800, color: '#087eb7' }}>{money(total)}</strong>
            <span style={{ fontSize: '11px', color: '#7b8993', fontWeight: 500 }}>{totalItemsCount} {totalItemsCount === 1 ? 'item' : 'itens'}</span>
          </div>
        </div>
        <div className="shopping-card-footer" style={{ marginTop: '16px', borderTop: '1px solid #edf1f3', paddingTop: '12px' }}><div className="shopping-avatar-stack">{people.slice(0, 4).map((person) => <span key={person.id}>{participantAvatar(person, true)}</span>)}{people.length > 4 ? <em>+{people.length - 4}</em> : null}</div><span>{people.length} {people.length === 1 ? 'participante' : 'participantes'}</span></div>
      </button>
      {list.status === 'open' ? (
        <div className="shopping-card-open-actions">
          <button type="button" className="shopping-card-action shopping-card-action--cancel" onClick={onCancel}><XCircle size={15} /> Cancelar</button>
          <button type="button" className="shopping-card-action shopping-card-action--finalize" onClick={onFinalize}><CheckCircle2 size={15} /> Finalizar</button>
        </div>
      ) : (
        <button type="button" className="shopping-duplicate-button" onClick={onDuplicate}><Copy size={15} /> Duplicar</button>
      )}
    </article>
  );
}

function EditShoppingListModal({ list, onClose, onSave }: { list: ShoppingList; onClose: () => void; onSave: (form: { name: string; date: string; note: string }) => void }) {
  const [name, setName] = useState(list.name);
  const [date, setDate] = useState(list.date);
  const [note, setNote] = useState(list.note ?? '');
  function submit(event: FormEvent) { event.preventDefault(); if (name.trim() && date) onSave({ name, date, note }); }
  return (
    <div className="modal-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
      <form className="modal-card shopping-create-modal" onSubmit={submit}>
        <div className="modal-header">
          <img src={rubyDiamond} alt="" className="modal-logo" />
          <div className="modal-title-container">
            <h2 className="modal-title">Editar lista</h2>
            <span className="modal-subtitle">Atualize os detalhes da lista.</span>
          </div>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="shopping-modal-body">
          <div className="shopping-form-grid">
            <label className="shopping-field shopping-field--wide"><span>Nome da lista</span><input autoFocus maxLength={80} value={name} onChange={(e) => setName(e.target.value)} required /></label>
            <label className="shopping-field"><span>Data da lista</span><input type="date" value={date} onChange={(e) => setDate(e.target.value)} required /></label>
            <label className="shopping-field shopping-field--wide"><span>Observação <small>opcional</small></span><textarea maxLength={300} value={note} onChange={(e) => setNote(e.target.value)} /></label>
          </div>
        </div>
        <div className="shopping-modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary" disabled={!name.trim() || !date}>Salvar</button>
        </div>
      </form>
    </div>
  );
}

function CreateShoppingListModal({ friends, onClose, onCreate }: { friends: FriendUser[]; onClose: () => void; onCreate: (form: { name: string; date: string; note: string; friendIds: string[] }) => void }) {
  const [name, setName] = useState('');
  const [shareMode, setShareMode] = useState<'private' | 'shared'>('private');
  const [friendIds, setFriendIds] = useState<string[]>([]);
  function selectShareMode(mode: 'private' | 'shared') { setShareMode(mode); if (mode === 'private') setFriendIds([]); }
  function submit(event: FormEvent) {
    event.preventDefault();
    if (!name.trim()) return;
    onCreate({ name, date: new Date().toISOString().slice(0, 10), note: '', friendIds: shareMode === 'shared' ? friendIds : [] });
  }
  return (
    <div className="modal-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
      <form className="modal-card shopping-create-modal" onSubmit={submit}>
        <div className="modal-header">
          <img src={rubyDiamond} alt="" className="modal-logo" />
          <div className="modal-title-container">
            <h2 className="modal-title">Nova lista de compras</h2>
            <span className="modal-subtitle">Dê um nome e escolha quem participa.</span>
          </div>
          <button type="button" className="modal-close-button" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="shopping-modal-body">
          <label className="shopping-field"><span>Nome da lista</span><input autoFocus maxLength={80} value={name} onChange={(e) => setName(e.target.value)} placeholder="Ex.: Churrasco de sábado" required /></label>
          <div className="shopping-share-mode" role="radiogroup" aria-label="Tipo da lista" style={{ marginTop: 20 }}>
            <button type="button" className={shareMode === 'private' ? 'active' : ''} onClick={() => selectShareMode('private')}><UserRound size={16} /> Só para mim</button>
            <button type="button" className={shareMode === 'shared' ? 'active' : ''} onClick={() => selectShareMode('shared')}><Users size={16} /> Compartilhar</button>
          </div>
          {shareMode === 'shared'
            ? <fieldset className="shopping-friends-picker"><legend>Amigos participantes</legend>{friends.length ? <div>{friends.map((friend) => { const checked = friendIds.includes(friend.id); return <label key={friend.id} className={checked ? 'selected' : ''}><input type="checkbox" checked={checked} onChange={() => setFriendIds((current) => checked ? current.filter((id) => id !== friend.id) : [...current, friend.id])} />{participantAvatar(friend)}<span><strong>{friend.name}</strong><small>Poderá adicionar e editar itens</small></span><em>{checked ? <Check size={15} /> : <Plus size={15} />}</em></label>; })}</div> : <p className="shopping-no-friends"><Users size={18} /> Conecte-se com amigos para compartilhar uma lista.</p>}</fieldset>
            : <p className="shopping-private-note"><UserRound size={17} /> Esta lista ficará visível e editável apenas para você.</p>}
        </div>
        <div className="shopping-modal-actions">
          <button type="button" className="button-secondary" onClick={onClose}>Cancelar</button>
          <button type="submit" className="button-primary" disabled={!name.trim()}><Plus size={16} /> Criar lista</button>
        </div>
      </form>
    </div>
  );
}

function ShoppingListDetail({ list, currentUser, friends, allLists, onBack, onChange, onListItemAdded, openItemSignal }: { list: ShoppingList; currentUser: FriendUser; friends: FriendUser[]; allLists: ShoppingList[]; onBack: () => void; onChange: (list: ShoppingList) => void; onListItemAdded?: (event: { listId: string; listName: string; itemName: string; participantIds: string[] }) => void; openItemSignal?: number }) {
  const [draft, setDraft] = useState<ItemDraft>(emptyItem);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [itemModalOpen, setItemModalOpen] = useState(false);
  const [suggestionsOpen, setSuggestionsOpen] = useState(false);
  const suggestionRef = useRef<HTMLDivElement>(null);
  const lastProcessedItemSignalRef = useRef(openItemSignal);
  const users = [currentUser, ...friends];
  const userById = new Map(users.map((user) => [user.id, user]));
  const canEdit = list.status === 'open';
  const productNames = useMemo(() => [...new Map(allLists.flatMap((entry) => entry.items).map((item) => [normalizeProduct(item.name), item.name])).values()].sort((a, b) => a.localeCompare(b, 'pt-BR')), [allLists]);
  const suggestions = draft.name.trim() ? productNames.filter((name) => normalizeProduct(name).includes(normalizeProduct(draft.name)) && normalizeProduct(name) !== normalizeProduct(draft.name)).slice(0, 6) : [];
  const draftQuantity = parseDecimal(draft.quantity);
  const draftUnitPrice = draft.unitPrice.trim() ? parseDecimal(draft.unitPrice) : 0;
  const itemDraftValid = Boolean(draft.name.trim()) && draftQuantity > 0 && draftUnitPrice >= 0;

  useEffect(() => {
    function outside(event: PointerEvent) { if (!suggestionRef.current?.contains(event.target as Node)) setSuggestionsOpen(false); }
    document.addEventListener('pointerdown', outside);
    return () => document.removeEventListener('pointerdown', outside);
  }, []);
  useEffect(() => {
    if (openItemSignal && openItemSignal !== lastProcessedItemSignalRef.current) {
      lastProcessedItemSignalRef.current = openItemSignal;
      if (canEdit) {
        resetDraft();
        setItemModalOpen(true);
      }
    }
  }, [openItemSignal, canEdit]);
  function resetDraft() { setDraft(emptyItem); setEditingId(null); setSuggestionsOpen(false); }
  function closeItemModal() { resetDraft(); setItemModalOpen(false); }
  function saveItem(event: FormEvent) {
    event.preventDefault();
    const quantity = parseDecimal(draft.quantity);
    const unitPrice = draft.unitPrice.trim() ? parseDecimal(draft.unitPrice) : 0;
    if (!draft.name.trim() || quantity <= 0 || unitPrice < 0) return;
    const now = new Date().toISOString();
    const items = editingId
      ? list.items.map((item) => item.id === editingId ? { ...item, name: draft.name.trim(), quantity, unitPrice } : item)
      : [...list.items, { id: crypto.randomUUID(), name: draft.name.trim(), quantity, unitPrice, addedById: currentUser.id, createdAt: now }];
    onChange({ ...list, items });
    if (!editingId) onListItemAdded?.({ listId: list.id, listName: list.name, itemName: draft.name.trim(), participantIds: list.participantIds.filter((id) => id !== currentUser.id) });
    resetDraft();
    setItemModalOpen(false);
  }
  function editItem(item: ShoppingListItem) { setEditingId(item.id); setDraft({ name: item.name, quantity: String(item.quantity).replace('.', ','), unitPrice: item.unitPrice > 0 ? item.unitPrice.toFixed(2).replace('.', ',') : '' }); setItemModalOpen(true); }
  function removeItem(id: string) { onChange({ ...list, items: list.items.filter((item) => item.id !== id) }); if (editingId === id) resetDraft(); }
  function togglePurchased(id: string) { onChange({ ...list, items: list.items.map((item) => item.id === id ? { ...item, purchased: !item.purchased } : item) }); }
  const total = listTotal(list);
  const byName = (a: ShoppingListItem, b: ShoppingListItem) => a.name.localeCompare(b.name, 'pt-BR');
  const toBuyItems = list.items.filter((item) => !item.purchased).sort(byName);
  const boughtItems = list.items.filter((item) => item.purchased).sort(byName);

  const renderItemRow = (item: ShoppingListItem) => {
    const hasPrice = item.unitPrice > 0;
    return (
      <div className={`shopping-item-row${item.purchased ? ' shopping-item-row--done' : ''}`} key={item.id}>
        <button type="button" className="shopping-item-check" data-checked={item.purchased ? 'true' : 'false'} onClick={() => togglePurchased(item.id)} disabled={!canEdit} aria-label={item.purchased ? `Desmarcar ${item.name}` : `Marcar ${item.name} como comprado`}>{item.purchased ? <Check size={14} /> : null}</button>
        <strong className="shopping-item-title">{item.name}</strong>
        <span className="shopping-item-qty" data-label="Qtd">{item.quantity.toLocaleString('pt-BR')}</span>
        <span className={`shopping-item-value${!hasPrice ? ' shopping-price-pending' : ''}`} data-label="Valor">{hasPrice ? money(item.unitPrice) : 'A definir'}</span>
        <div className="shopping-item-actions">{canEdit ? <><button type="button" onClick={() => editItem(item)} aria-label={`Editar ${item.name}`}><Pencil size={15} /></button><button type="button" className="danger" onClick={() => removeItem(item.id)} aria-label={`Remover ${item.name}`}><Trash2 size={15} /></button></> : null}</div>
      </div>
    );
  };

  return (
    <div className="shopping-page shopping-detail-page">
      <button type="button" className="shopping-detail-back-button" onClick={onBack}><ChevronLeft size={17} /> Voltar</button>
      {!canEdit ? <div className="shopping-readonly-note"><CheckCircle2 size={17} /><span><strong>Lista somente para consulta.</strong> Itens de listas {list.status === 'finalized' ? 'finalizadas' : 'canceladas'} não podem ser alterados.</span></div> : null}
      <section className="shopping-items-panel">
        {list.items.length ? (
          <div className="shopping-items-table">
            <div className="shopping-items-columns"><span aria-hidden="true" /><span>Produto</span><span>Qtd</span><span>Valor</span><span aria-hidden="true" /></div>
            {toBuyItems.map(renderItemRow)}
            {boughtItems.map(renderItemRow)}
          </div>
        ) : (
          <div className="shopping-items-empty"><PackageOpen size={25} /><strong>Nenhum item adicionado</strong><span>{canEdit ? 'Toque em Adicionar para começar a lista.' : 'Esta lista foi encerrada sem itens.'}</span></div>
        )}
      </section>
      {list.status === 'finalized' ? (
        <footer className="shopping-completion"><CheckCircle2 size={18} /><span>Finalizada em <strong>{dateLabel(list.finalizedAt)}</strong> por <strong>{list.finalizedById === currentUser.id ? 'você' : userById.get(list.finalizedById ?? '')?.name ?? 'um participante'}</strong>.</span></footer>
      ) : null}

      <div className="shopping-detail-total-bar" aria-label="Total da lista">
        <span>Total da lista</span>
        <strong>{money(total)}</strong>
      </div>
      {itemModalOpen ? (
        <div className="modal-overlay shopping-item-modal-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) closeItemModal(); }}>
          <form className="modal-card shopping-create-modal shopping-item-modal" onSubmit={saveItem}>
            <div className="modal-header">
              <img src={rubyDiamond} alt="" className="modal-logo" />
              <div className="modal-title-container">
                <h2 className="modal-title">{editingId ? 'Editar produto' : 'Adicionar produto'}</h2>
                <span className="modal-subtitle">Informe nome, quantidade e valor.</span>
              </div>
              <button type="button" className="modal-close-button" onClick={closeItemModal} aria-label="Fechar"><X size={18} /></button>
            </div>
            <div className="shopping-modal-body shopping-item-modal-body">
              <div className="shopping-item-modal-grid">
                <div className="shopping-field shopping-product-field" ref={suggestionRef}>
                  <label htmlFor="shopping-product-modal">Nome do produto</label>
                  <div><Search size={16} /><input id="shopping-product-modal" autoFocus autoComplete="off" value={draft.name} onFocus={() => setSuggestionsOpen(true)} onChange={(event) => { setDraft({ ...draft, name: event.target.value }); setSuggestionsOpen(true); }} placeholder="Digite o nome do produto" required /></div>
                  {suggestionsOpen && suggestions.length ? <div className="shopping-suggestions" role="listbox">{suggestions.map((name) => <button type="button" role="option" key={name} onClick={() => { setDraft({ ...draft, name }); setSuggestionsOpen(false); }}><Search size={14} />{name}</button>)}</div> : null}
                </div>
                <label className="shopping-field"><span>Quantidade</span><input inputMode="decimal" value={draft.quantity} onChange={(event) => setDraft({ ...draft, quantity: event.target.value })} required /></label>
                <label className="shopping-field"><span>Valor</span><div className="shopping-money-field"><em>R$</em><input inputMode="decimal" value={draft.unitPrice} onChange={(event) => setDraft({ ...draft, unitPrice: event.target.value })} placeholder="0,00" /></div></label>
              </div>
            </div>
            <div className="shopping-modal-actions">
              <button type="button" className="button-secondary" onClick={closeItemModal}>Cancelar</button>
              <button type="submit" className="button-primary" disabled={!itemDraftValid}>{editingId ? <Check size={16} /> : <Plus size={16} />}{editingId ? 'Salvar' : 'Adicionar'}</button>
            </div>
          </form>
        </div>
      ) : null}
    </div>
  );
}

function ConfirmListAction({ action, list, onClose, onConfirm }: { action: 'finalize' | 'cancel'; list: ShoppingList; onClose: () => void; onConfirm: () => void }) {
  const finalize = action === 'finalize';
  return <div className="modal-overlay"><div className="modal-card shopping-confirm-modal"><div className={`shopping-confirm-icon ${finalize ? '' : 'danger'}`}>{finalize ? <CheckCircle2 size={24} /> : <XCircle size={24} />}</div><h2>{finalize ? 'Finalizar compra?' : 'Cancelar esta lista?'}</h2><p>{finalize ? <>A lista <strong>{list.name}</strong> ficará somente para consulta e será movida para o histórico.</> : <>A lista <strong>{list.name}</strong> será encerrada e os itens não poderão mais ser editados.</>}</p><div><button type="button" className="button-secondary" onClick={onClose}>Voltar</button><button type="button" className={finalize ? 'button-primary' : 'button-danger'} onClick={onConfirm}>{finalize ? 'Finalizar compra' : 'Cancelar lista'}</button></div></div></div>;
}


