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

function renderPanel(items) {
  const list = document.getElementById('notifList');
  if (!items.length) {
    list.innerHTML = `<div class="notif-empty">No follow-ups overdue by 7+ days. 🎉</div>`;
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
        <div class="notif-item" data-tab="${i.tabId}" data-rid="${i.id}">
          <div class="ni-top"><span>${escapeHtml(i.name)}</span><span class="pill pill-${i.bucket}">${i.days}d</span></div>
          <div class="ni-sub">${i.sourceLabel} pipeline</div>
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
  const items = computeOverdue();
  const badge = document.getElementById('notifBadge');
  if (items.length) {
    badge.textContent = items.length > 99 ? '99+' : String(items.length);
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }
  renderPanel(items);
  maybeFireDesktopNotifications(items);
  return items;
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
