import { loadAllTabs, exportFullBackupData, restoreFullBackupData, TABS } from './store.js';
import { initNotifications, setNotifNavigateHandler } from './notifications.js';
import { downloadWorkbook, readWorkbook, parseSheetRows } from './excel.js';
import { showToast } from './utils.js';

import * as agapeTab from './tabs/agape.js';
import * as actionPlanTab from './tabs/actionplan.js';
import * as sellingProspectTab from './tabs/sellingProspectList.js';
import * as recruitProspectTab from './tabs/recruitProspectList.js';
import * as scorecardTab from './tabs/scorecard.js';
import * as productKnowledgeTab from './tabs/productKnowledge.js';
import * as sellingTab from './tabs/selling.js';
import * as recruitmentTab from './tabs/recruitment.js';
import * as pipelineTab from './tabs/salesPipeline.js';

const TAB_MODULES = {
  'agape': agapeTab,
  'action-plan': actionPlanTab,
  'selling-prospect-list': sellingProspectTab,
  'recruit-prospect-list': recruitProspectTab,
  'scorecard': scorecardTab,
  'product-knowledge': productKnowledgeTab,
  'selling': sellingTab,
  'recruitment': recruitmentTab,
  'sales-pipeline': pipelineTab,
};

const DEFAULT_TAB = 'agape';
const mainContent = document.getElementById('mainContent');

async function renderRoute() {
  const hash = (location.hash || '#' + DEFAULT_TAB).slice(1);
  const focusParam = new URLSearchParams(location.hash.split('?')[1] || '');
  const tabId = hash.split('?')[0];
  const mod = TAB_MODULES[tabId] || TAB_MODULES[DEFAULT_TAB];

  document.querySelectorAll('.tab-nav a').forEach((a) => {
    a.classList.toggle('active', a.dataset.tab === tabId);
  });

  mainContent.innerHTML = '<div class="empty-state">Loading…</div>';
  try {
    await mod.render(mainContent, { focusRowId: focusParam.get('focus') || null });
  } catch (err) {
    console.error(err);
    mainContent.innerHTML = `<div class="empty-state">Something went wrong rendering this tab. Check the console for details.</div>`;
  }

  // collapse mobile sidebar after navigating
  if (window.innerWidth <= 760) document.getElementById('sidebar').classList.remove('open');
}

function initNav() {
  window.addEventListener('hashchange', renderRoute);
  document.getElementById('navToggle').addEventListener('click', () => {
    const sb = document.getElementById('sidebar');
    if (window.innerWidth <= 760) sb.classList.toggle('open');
    else sb.classList.toggle('collapsed');
  });
}

function initBrandName() {
  const el = document.getElementById('brandName');
  const saved = localStorage.getItem('crm:brandName');
  if (saved) el.textContent = saved;
  el.title = 'Click to rename';
  el.style.cursor = 'pointer';
  el.addEventListener('click', () => {
    const name = prompt('Name this CRM (shown only in your browser):', el.textContent);
    if (name && name.trim()) {
      el.textContent = name.trim();
      localStorage.setItem('crm:brandName', name.trim());
    }
  });
}

function initBackupControls() {
  // Excel cells cap out around 32,767 characters, and some tabs (Selling,
  // Recruitment, Recruit Prospect List...) serialize to far more than that.
  // So each tab's JSON is split into chunks, one per row, and reassembled on import.
  const CHUNK_SIZE = 25000;
  function chunkString(s) {
    const out = [];
    for (let i = 0; i < s.length; i += CHUNK_SIZE) out.push(s.slice(i, i + CHUNK_SIZE));
    return out.length ? out : [''];
  }

  document.getElementById('backupExportBtn').addEventListener('click', async () => {
    const all = await exportFullBackupData();
    const sheets = TABS.map((t) => ({
      name: t.id,
      columns: [{ key: 'json', header: 'data' }],
      rows: chunkString(JSON.stringify(all[t.id])).map((chunk) => ({ json: chunk })),
    }));
    downloadWorkbook(sheets, `crm-full-backup-${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast('Full backup downloaded.');
  });

  document.getElementById('backupImportInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const restored = {};
      for (const t of TABS) {
        if (!wb.SheetNames.includes(t.id)) continue;
        const rows = parseSheetRows(wb, [{ key: 'json', header: 'data', aliases: ['data'] }], t.id);
        const joined = rows.map((r) => r.json || '').join('');
        if (joined) {
          try { restored[t.id] = JSON.parse(joined); } catch (err) { /* skip bad sheet */ }
        }
      }
      if (!Object.keys(restored).length) {
        showToast('This file does not look like a My CRM backup.');
        return;
      }
      if (!confirm('This will replace ALL current data in this browser with the backup contents. Continue?')) return;
      await restoreFullBackupData(restored);
      showToast('Backup restored.');
      renderRoute();
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  });
}

async function main() {
  initNav();
  initBrandName();
  initBackupControls();
  await loadAllTabs();
  setNotifNavigateHandler((tabId, rowId) => {
    location.hash = `#${tabId}?focus=${encodeURIComponent(rowId)}`;
  });
  initNotifications();
  renderRoute();
}

main();
