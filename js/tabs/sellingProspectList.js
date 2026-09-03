import { loadTab, markDirty } from '../store.js';
import { uid, escapeHtml, num, showToast, withScrollPreserved, wireMoneyInput, todayISO } from '../utils.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';

const STAGE_NAMES = {
  'Prospecting': 'Level 1', 'Engagement & Approaching': 'Level 2', 'Set Appointment': 'Level 3',
  'Client Meeting': 'Level 4', 'Closing & Referrals': 'Level 5', 'Servicing': 'Level 6',
};

export async function render(container) {
  const d = await loadTab('selling-prospect-list');
  d.levels.forEach((lvl) => (d.board[lvl] || (d.board[lvl] = [])).forEach((c) => { if (!c._id) c._id = uid(); }));

  // Which stage the carousel is currently showing. UI-only state (not persisted) —
  // same pattern as the DataGrid's page number.
  let stageIdx = 0;

  function paint() { withScrollPreserved(container, paintInner); }

  function paintInner() {
    const lvl = d.levels[stageIdx];
    const cards = d.board[lvl] || (d.board[lvl] = []);
    container.innerHTML = `
      <div class="tab-header">
        <div>
          <h1 class="tab-title">🧲 Selling Journey</h1>
          <p class="tab-subtitle">Move each lead forward one stage at a time as they progress — no dragging, just tap ›.</p>
        </div>
        <div class="tab-actions">
          <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
          <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
        </div>
      </div>
      <div class="carousel-stepper">
        ${d.levels.map((l, i) => `
          <button class="carousel-step ${i === stageIdx ? 'active' : ''}" data-idx="${i}">
            ${escapeHtml(STAGE_NAMES[l] || '')} · ${escapeHtml(l)}<span class="step-count">${(d.board[l] || []).length}</span>
          </button>`).join('')}
      </div>
      <div class="carousel-nav">
        <button class="btn btn-light carousel-nav-btn" id="prevStageBtn" ${stageIdx === 0 ? 'disabled' : ''} title="Previous stage">‹ Prev stage</button>
        <h4>${escapeHtml(lvl)} <span class="text-muted">(${cards.length})</span></h4>
        <button class="btn btn-light carousel-nav-btn" id="nextStageBtn" ${stageIdx === d.levels.length - 1 ? 'disabled' : ''} title="Next stage">Next stage ›</button>
      </div>
      <div class="selling-cards">
        ${cards.length === 0 ? `<div class="empty-state"><div class="es-emoji">🗂️</div>No leads in ${escapeHtml(lvl)} yet. Add one below.</div>` : ''}
        ${cards.map((c) => `
          <div class="selling-card" data-id="${c._id}">
            <button class="card-move card-move-back" ${stageIdx === 0 ? 'disabled' : ''} title="Move back to ${escapeHtml(d.levels[stageIdx - 1] || '')}">‹</button>
            <input type="text" class="card-name" value="${escapeHtml(c.name)}" placeholder="Name" title="${escapeHtml(c.name)}" />
            <input type="text" inputmode="decimal" class="card-ape money-input" value="${c.potentialApe ?? ''}" placeholder="APE" />
            <button class="dg-del-btn card-del" title="Remove">✕</button>
            <button class="card-move card-move-fwd" ${stageIdx === d.levels.length - 1 ? 'disabled' : ''} title="Move forward to ${escapeHtml(d.levels[stageIdx + 1] || '')}">›</button>
          </div>`).join('')}
      </div>
      <button class="btn btn-light btn-sm selling-add" id="addCardBtn">+ Add to ${escapeHtml(lvl)}</button>
    `;
    wire();
  }

  function persist() { markDirty('selling-prospect-list'); }

  function goToStage(i) {
    stageIdx = Math.max(0, Math.min(d.levels.length - 1, i));
    paint();
  }

  function moveCard(item, fromLvl, direction) {
    const toIdx = d.levels.indexOf(fromLvl) + direction;
    if (toIdx < 0 || toIdx >= d.levels.length) return;
    const toLvl = d.levels[toIdx];
    d.board[fromLvl].splice(d.board[fromLvl].indexOf(item), 1);
    (d.board[toLvl] || (d.board[toLvl] = [])).push(item);
    persist();
    paint();
    // Crossing from Prospecting into Engagement & Approaching means this lead
    // is now being actively worked — mirror them into the Selling tab so
    // they're tracked for follow-up alerts too.
    if (fromLvl === d.levels[0] && toLvl === d.levels[1] && item.name && item.name.trim()) {
      syncToSellingTab(item.name.trim());
    }
  }

  async function syncToSellingTab(name) {
    const sd = await loadTab('selling');
    const today = todayISO();
    const norm = name.toLowerCase();
    const existing = sd.rows.find((r) => (r.name || '').trim().toLowerCase() === norm);
    if (existing) {
      existing.lastApproachDate = today;
      if (!existing.dateAdded) existing.dateAdded = today;
      markDirty('selling');
      showToast(`${name} is already in Selling — Last Approach Date updated to today.`);
    } else {
      sd.rows.unshift({
        _id: uid(), response: '', source: '', demographic: '', name, status: '',
        occupation: '', dateAdded: today, lastApproachDate: today, approachMethod: '', remarks: '',
      });
      markDirty('selling');
      showToast(`Added ${name} to Selling — Date Added & Last Approach Date set to today.`);
    }
  }

  function wire() {
    container.querySelectorAll('.carousel-step').forEach((btn) => {
      btn.addEventListener('click', () => goToStage(parseInt(btn.dataset.idx, 10)));
    });
    const prevBtn = container.querySelector('#prevStageBtn');
    const nextBtn = container.querySelector('#nextStageBtn');
    if (prevBtn) prevBtn.addEventListener('click', () => goToStage(stageIdx - 1));
    if (nextBtn) nextBtn.addEventListener('click', () => goToStage(stageIdx + 1));

    const lvl = d.levels[stageIdx];
    container.querySelectorAll('.selling-card').forEach((card) => {
      const item = d.board[lvl].find((c) => c._id === card.dataset.id);
      if (!item) return;
      card.querySelector('.card-name').addEventListener('input', (e) => { item.name = e.target.value; persist(); });
      card.querySelector('.card-ape').addEventListener('input', (e) => { item.potentialApe = e.target.value === '' ? null : num(e.target.value); persist(); });
      card.querySelector('.card-del').addEventListener('click', () => {
        d.board[lvl].splice(d.board[lvl].indexOf(item), 1);
        persist(); paint();
      });
      const backBtn = card.querySelector('.card-move-back');
      if (backBtn) backBtn.addEventListener('click', () => moveCard(item, lvl, -1));
      const fwdBtn = card.querySelector('.card-move-fwd');
      if (fwdBtn) fwdBtn.addEventListener('click', () => moveCard(item, lvl, 1));
    });

    const addBtn = container.querySelector('#addCardBtn');
    if (addBtn) addBtn.addEventListener('click', () => {
      d.board[lvl].push({ _id: uid(), name: '', potentialApe: null });
      persist(); paint();
      const inputs = container.querySelectorAll('.selling-card .card-name');
      if (inputs.length) inputs[inputs.length - 1].focus();
    });

    container.querySelector('#exportBtn').addEventListener('click', doExport);
    container.querySelector('#importInput').addEventListener('change', doImport);

    container.querySelectorAll('.money-input').forEach((el) => wireMoneyInput(el));
  }

  function doExport() {
    const maxLen = Math.max(0, ...d.levels.map((lvl) => (d.board[lvl] || []).length));
    const rows = [];
    for (let i = 0; i < maxLen; i++) {
      const row = { no: i + 1 };
      d.levels.forEach((lvl) => {
        const c = (d.board[lvl] || [])[i];
        row[lvl] = c ? c.name : '';
        if (lvl === 'Closing & Referrals' || lvl === 'Servicing') row.potentialApe = c && c.potentialApe != null ? c.potentialApe : (row.potentialApe || '');
      });
      rows.push(row);
    }
    const columns = [
      { key: 'no', header: '#' },
      ...d.levels.map((lvl) => ({ key: lvl, header: `${STAGE_NAMES[lvl] || ''} - ${lvl}`.trim() })),
      { key: 'potentialApe', header: 'Potential APE' },
    ];
    downloadSingleSheet('SELLING JOURNEY', columns, rows, 'Selling Journey.xlsx');
  }

  async function doImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const columns = [
        { key: 'no', header: '#', aliases: ['#', 'No.'] },
        ...d.levels.map((lvl) => ({ key: lvl, header: lvl, aliases: [lvl, STAGE_NAMES[lvl], `${STAGE_NAMES[lvl]} - ${lvl}`] })),
        { key: 'potentialApe', header: 'Potential APE', aliases: ['Potential APE'] },
      ];
      const rows = parseSheetRows(wb, columns, 'SELLING JOURNEY');
      d.levels.forEach((lvl) => { d.board[lvl] = []; });
      rows.forEach((r) => {
        d.levels.forEach((lvl) => {
          const name = r[lvl];
          if (name && String(name).trim()) {
            d.board[lvl].push({ _id: uid(), name: String(name).trim(), potentialApe: r.potentialApe ? num(r.potentialApe) : null });
          }
        });
      });
      persist(); paint();
      showToast(`Imported ${rows.length} row(s) from Excel.`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  }

  paint();
}
