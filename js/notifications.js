import { getTab, loadTab } from './store.js';
import { daysSince, escapeHtml } from './utils.js';

const SOURCES = [
  { tabId: 'selling', label: 'Selling' },
  { tabId: 'recruitment', label: 'Recruitment' },
];

let navigateHandler = null;
export function setNotifNavigateHandler(fn) { navigateHandler = fn; }

function bucketFor(days) {
  if (days >= 21) return 21;
  if (days >= 14) return 14;
  if (days >= 7) return 7;
  return null;
}

export function computeOverdue() {
  const items = [];
  for (const src of SOURCES) {
    const data = getTab(src.tabId);
    if (!data || !data.rows) continue;
    for (const row of data.rows) {
      if (!row.lastApproachDate) continue;
      if (row.response === 'CLIENT') continue; // already converted, no follow-up needed
      const days = daysSince(row.lastApproachDate);
      if (days === null) continue;
      const bucket = bucketFor(days);
      if (!bucket) continue;
      items.push({ tabId: src.tabId, sourceLabel: src.label, id: row._id, name: row.name || '(no name)', days, bucket });
    }
  }
  items.sort((a, b) => b.days - a.days);
  return items;
}

// "Read" alerts are tracked per person PER THRESHOLD (7/14/21), not just per
// person — so marking a 7-day alert read doesn't silence the 21-day
// escalation later if it's still unresolved by then.
let readSet = null;
function loadReadSet() {
  if (readSet) return readSet;
  try { readSet = new Set(JSON.parse(localStorage.getItem('crm:notifRead') || '[]')); } catch (e) { readSet = new Set(); }
  return readSet;
}
function saveReadSet() { localStorage.setItem('crm:notifRead', JSON.stringify(Array.from(loadReadSet()))); }
function readKey(item) { return `${item.tabId}:${item.id}:${item.bucket}`; }
function isRead(item) { return loadReadSet().has(readKey(item)); }
function markRead(item) { loadReadSet().add(readKey(item)); saveReadSet(); }
function markAllRead(items) {
  const rs = loadReadSet();
  items.forEach((i) => rs.add(readKey(i)));
  saveReadSet();
}
// Drop read-markers for alerts that no longer exist (contact was made,
// converted to CLIENT, etc.) so storage doesn't grow forever.
function pruneReadSet(allItems) {
  const rs = loadReadSet();
  const live = new Set(allItems.map(readKey));
  let changed = false;
  for (const key of Array.from(rs)) {
    if (!live.has(key)) { rs.delete(key); changed = true; }
  }
  if (changed) saveReadSet();
}

let currentUnread = [];

function renderPanel(items) {
  const list = document.getElementById('notifList');
  const markAllBtn = document.getElementById('markAllReadBtn');
  markAllBtn.classList.toggle('hidden', items.length === 0);

  if (!items.length) {
    list.innerHTML = `<div class="notif-empty">No unread follow-ups overdue by 7+ days. 🎉</div>`;
    return;
  }
  const groups = [21, 14, 7].map((b) => ({ b, rows: items.filter((i) => i.bucket === b) })).filter((g) => g.rows.length);
  list.innerHTML = groups
    .map(
      (g) => `
      <div class="notif-group-label">${g.b}+ days since last contact</div>
      ${g.rows
        .map(
          (i) => `
        <div class="notif-item" data-tab="${i.tabId}" data-rid="${i.id}" data-bucket="${i.bucket}">
          <div class="ni-top">
            <span>${escapeHtml(i.name)}</span>
            <span class="pill pill-${i.bucket}">${i.days}d</span>
          </div>
          <div class="ni-sub">
            <span>${i.sourceLabel} pipeline</span>
            <button class="notif-mark-read" title="Mark as read">✓ Mark as read</button>
          </div>
        </div>`
        )
        .join('')}
    `
    )
    .join('');

  list.querySelectorAll('.notif-item').forEach((el) => {
    el.addEventListener('click', () => {
      if (navigateHandler) navigateHandler(el.dataset.tab, el.dataset.rid);
      document.getElementById('notifPanel').classList.add('hidden');
    });
    el.querySelector('.notif-mark-read').addEventListener('click', (e) => {
      e.stopPropagation();
      markRead({ tabId: el.dataset.tab, id: el.dataset.rid, bucket: +el.dataset.bucket });
      refreshNotifications();
    });
  });
}

let lastNotifiedBucket = null;
try { lastNotifiedBucket = JSON.parse(localStorage.getItem('crm:notified') || '{}'); } catch (e) { lastNotifiedBucket = {}; }

function maybeFireDesktopNotifications(items) {
  if (typeof Notification === 'undefined' || Notification.permission !== 'granted') return;
  let changed = false;
  for (const i of items) {
    const key = `${i.tabId}:${i.id}`;
    if (lastNotifiedBucket[key] !== i.bucket) {
      lastNotifiedBucket[key] = i.bucket;
      changed = true;
      try {
        new Notification(`${i.name} — ${i.bucket}+ days since last contact`, {
          body: `${i.sourceLabel} pipeline follow-up is overdue.`,
          tag: key,
        });
      } catch (e) { /* ignore */ }
    }
  }
  if (changed) localStorage.setItem('crm:notified', JSON.stringify(lastNotifiedBucket));
}

export async function refreshNotifications() {
  await Promise.all(SOURCES.map((s) => loadTab(s.tabId)));
  const allItems = computeOverdue();
  pruneReadSet(allItems);
  const unread = allItems.filter((i) => !isRead(i));
  currentUnread = unread;

  const badge = document.getElementById('notifBadge');
  if (unread.length) {
    badge.textContent = unread.length > 99 ? '99+' : String(unread.length);
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }
  renderPanel(unread);
  maybeFireDesktopNotifications(unread);
  return unread;
}

export function initNotifications() {
  const bell = document.getElementById('notifBell');
  const panel = document.getElementById('notifPanel');
  bell.addEventListener('click', (e) => {
    e.stopPropagation();
    panel.classList.toggle('hidden');
  });
  document.addEventListener('click', (e) => {
    if (!panel.contains(e.target) && e.target !== bell) panel.classList.add('hidden');
  });

  document.getElementById('markAllReadBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    markAllRead(currentUnread);
    refreshNotifications();
  });

  const toggle = document.getElementById('browserNotifToggle');
  if (typeof Notification === 'undefined') {
    toggle.disabled = true;
    toggle.closest('.notif-toggle').title = 'Your browser does not support desktop notifications';
  } else {
    toggle.checked = localStorage.getItem('crm:desktopNotif') === '1' && Notification.permission === 'granted';
    toggle.addEventListener('change', async () => {
      if (toggle.checked) {
        const perm = await Notification.requestPermission();
        if (perm !== 'granted') { toggle.checked = false; return; }
        localStorage.setItem('crm:desktopNotif', '1');
      } else {
        localStorage.setItem('crm:desktopNotif', '0');
      }
    });
  }

  refreshNotifications();
  setInterval(refreshNotifications, 5 * 60 * 1000); // recheck every 5 min while the tab is open
}
