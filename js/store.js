import { dbGet, dbSet } from './db.js';
import { debounce } from './utils.js';

export const TABS = [
  { id: 'agape', file: 'agape.json' },
  { id: 'action-plan', file: 'actionplan.json' },
  { id: 'selling-prospect-list', file: 'selling_prospect_list.json' },
  { id: 'scorecard', file: 'scorecard.json' },
  { id: 'product-knowledge', file: 'product_knowledge.json' },
  { id: 'selling', file: 'selling.json' },
  { id: 'recruitment', file: 'recruitment.json' },
  { id: 'sales-pipeline', file: 'sales_pipeline.json' },
  { id: 'todo', file: 'todo.json' },
];

const cache = new Map();
const changeListeners = new Set();

function setSaveStatus(text) {
  const el = document.getElementById('saveStatus');
  if (el) el.textContent = text;
}

const debouncedPersist = debounce(async (id) => {
  setSaveStatus('Saving…');
  await dbSet(id, cache.get(id));
  setSaveStatus('All changes saved');
  changeListeners.forEach((fn) => fn(id));
}, 400);

export async function loadTab(id) {
  if (cache.has(id)) return cache.get(id);
  let data = await dbGet(id);
  if (data === null || data === undefined) {
    const meta = TABS.find((t) => t.id === id);
    const res = await fetch(`data/${meta.file}`);
    data = await res.json();
  }
  cache.set(id, data);
  return data;
}

export function getTab(id) {
  return cache.get(id);
}

// Call after mutating the object returned by loadTab/getTab, to persist.
export function markDirty(id) {
  setSaveStatus('Saving…');
  debouncedPersist(id);
}

export async function replaceTab(id, data) {
  cache.set(id, data);
  await dbSet(id, data);
  changeListeners.forEach((fn) => fn(id));
}

export function onTabSaved(fn) {
  changeListeners.add(fn);
  return () => changeListeners.delete(fn);
}

export async function loadAllTabs() {
  await Promise.all(TABS.map((t) => loadTab(t.id)));
}

export async function exportFullBackupData() {
  const out = {};
  for (const t of TABS) out[t.id] = await loadTab(t.id);
  return out;
}

export async function restoreFullBackupData(obj) {
  for (const t of TABS) {
    if (obj[t.id] !== undefined) {
      cache.set(t.id, obj[t.id]);
      await dbSet(t.id, obj[t.id]);
    }
  }
  changeListeners.forEach((fn) => fn('*'));
}
