import { loadTab, markDirty } from '../store.js';
import { uid, escapeHtml, num, showToast, withScrollPreserved, wireMoneyInput } from '../utils.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';

const STAGE_NAMES = {
  'Prospecting': 'Level 1', 'Engagement & Approaching': 'Level 2', 'Set Appointment': 'Level 3',
  'Client Meeting': 'Level 4', 'Closing & Referrals': 'Level 5', 'Servicing': 'Level 6',
};

export async function render(container) {
  const d = await loadTab('selling-prospect-list');
  d.levels.forEach((lvl) => (d.board[lvl] || (d.board[lvl] = [])).forEach((c) => { if (!c._id) c._id = uid(); }));

  function paint() { withScrollPreserved(container, paintInner); }

  function paintInner() {
    container.innerHTML = `
      <div class="tab-header">
        <div>
          <h1 class="tab-title">🧲 Selling Journey</h1>
          <p class="tab-subtitle">Drag your prospecting funnel forward stage by stage — add or remove names any time.</p>
        </div>
        <div class="tab-actions">
          <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
          <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
        </div>
      </div>
      <div class="kanban">
        ${d.levels.map((lvl) => `
          <div class="kanban-col" data-lvl="${escapeHtml(lvl)}">
            <h4><span>${escapeHtml(lvl)}</span><span class="text-muted">${(d.board[lvl] || []).length}</span></h4>
            <div class="kanban-cards">
              ${(d.board[lvl] || []).map((c) => `
                <div class="kanban-card" data-id="${c._id}">
                  <input type="text" class="card-name" value="${escapeHtml(c.name)}" placeholder="Name" title="${escapeHtml(c.name)}" />
                  <div class="kanban-card-row2">
                    <input type="text" inputmode="decimal" class="card-ape money-input" value="${c.potentialApe ?? ''}" placeholder="APE" />
                    <button class="dg-del-btn card-del">✕</button>
                  </div>
                </div>`).join('')}
            </div>
            <button class="btn btn-light btn-sm kanban-add">+ Add</button>
          </div>`).join('')}
      </div>
    `;
    wire();
  }

  function persist() { markDirty('selling-prospect-list'); }

  function wire() {
    container.querySelectorAll('.kanban-col').forEach((col) => {
      const lvl = col.dataset.lvl;
      col.querySelectorAll('.kanban-card').forEach((card) => {
        const item = d.board[lvl].find((c) => c._id === card.dataset.id);
        card.querySelector('.card-name').addEventListener('input', (e) => { item.name = e.target.value; persist(); });
        card.querySelector('.card-ape').addEventListener('input', (e) => { item.potentialApe = e.target.value === '' ? null : num(e.target.value); persist(); });
        card.querySelector('.card-del').addEventListener('click', () => {
          d.board[lvl].splice(d.board[lvl].indexOf(item), 1);
          persist(); paint();
        });
      });
      col.querySelector('.kanban-add').addEventListener('click', () => {
        d.board[lvl].push({ _id: uid(), name: '', potentialApe: null });
        persist(); paint();
        const inputs = container.querySelectorAll(`.kanban-col[data-lvl="${CSS.escape(lvl)}"] .card-name`);
        if (inputs.length) inputs[inputs.length - 1].focus();
      });
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
