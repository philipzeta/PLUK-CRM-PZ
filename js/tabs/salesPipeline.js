import { loadTab, markDirty } from '../store.js';
import { uid, num, fmtMoney, fmtPercent, escapeHtml, toISODate, showToast, withScrollPreserved } from '../utils.js';
import { createDataGrid } from '../datagrid.js';
import { downloadWorkbook, readWorkbook, parseSheetRows } from '../excel.js';

const ACT_COLUMNS = [
  { key: 'activity', header: 'Activity', aliases: ['Activity', 'ACTIVITY'] },
  { key: 'target', header: 'Target' },
  { key: 'w1', header: 'W1' }, { key: 'w2', header: 'W2' }, { key: 'w3', header: 'W3' }, { key: 'w4', header: 'W4' }, { key: 'w5', header: 'W5' },
];
const PIPE_COLUMNS = [
  { key: 'clientName', header: 'Client Name', aliases: ['Client Name'] },
  { key: 'product', header: 'Product', aliases: ['Product'] },
  { key: 'ape', header: 'APE', aliases: ['APE'] },
  { key: 'chance', header: '% Chance of Closing', aliases: ['% Chance of Closing', 'Chance'] },
  { key: 'targetClosingDate', header: 'Target Closing Date', aliases: ['Target Closing Date'], type: 'date' },
  { key: 'remarks', header: 'Remarks', aliases: ['Remarks'] },
];

export async function render(container) {
  const d = await loadTab('sales-pipeline');
  d.months.forEach((m) => {
    m.activities.forEach((a) => { if (!a._id) a._id = uid(); });
    m.pipeline.forEach((p) => { if (!p._id) p._id = uid(); });
  });
  d.months.sort((a, b) => a.key.localeCompare(b.key));

  let activeKey = d.months.length ? d.months[d.months.length - 1].key : null;

  function activeMonth() { return d.months.find((m) => m.key === activeKey); }

  function persist() { markDirty('sales-pipeline'); }

  function paint() { withScrollPreserved(container, paintInner); }

  function paintInner() {
    const m = activeMonth();
    container.innerHTML = `
      <div class="tab-header">
        <div>
          <h1 class="tab-title">📈 Sales Pipeline</h1>
          <p class="tab-subtitle">Weekly activity targets and your live deal pipeline — one board per month, add a new one any time.</p>
        </div>
        <div class="tab-actions">
          <button class="btn btn-light" id="addMonthBtn">+ Add month</button>
          <button class="btn btn-light" id="exportAllBtn">⬇ Export all months</button>
        </div>
      </div>

      <div class="month-tabs" id="monthTabs">
        ${d.months.map((mm) => `<button class="month-tab-btn ${mm.key === activeKey ? 'active' : ''}" data-key="${mm.key}">${escapeHtml(mm.label)}</button>`).join('')}
      </div>

      ${!m ? `<div class="empty-state"><div class="es-emoji">📈</div>No months yet — add one to get started.</div>` : `
      <div class="card">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
          <h3 style="margin:0">${escapeHtml(m.label)} — Weekly Activity</h3>
          <div class="tab-actions">
            <button class="btn btn-light btn-sm" id="dupMonthBtn">⧉ Duplicate as next month</button>
            <button class="btn btn-light btn-sm" id="exportMonthBtn">⬇ Export this month</button>
            <label class="btn btn-light btn-sm">⬆ Import this month<input type="file" id="importMonthInput" accept=".xlsx,.xls" hidden /></label>
            <button class="btn btn-danger btn-sm" id="delMonthBtn">🗑 Delete month</button>
          </div>
        </div>
        <div class="week-grid" style="margin-top:12px;grid-template-columns:1.6fr repeat(7,1fr);" id="actGrid">
          <div class="hd">Activity</div><div class="hd">Target</div><div class="hd">W1</div><div class="hd">W2</div><div class="hd">W3</div><div class="hd">W4</div><div class="hd">W5</div><div class="hd">Total / To go</div>
          ${m.activities.map((a) => {
            const total = a.weeks.reduce((s, v) => s + num(v), 0);
            const toGo = num(a.target) - total;
            return `
            <div><input class="cell-input act-name" data-id="${a._id}" type="text" value="${escapeHtml(a.name)}" /></div>
            <div><input class="cell-input act-f" data-id="${a._id}" data-f="target" type="number" value="${a.target ?? 0}" /></div>
            ${a.weeks.map((w, wi) => `<div><input class="cell-input act-week" data-id="${a._id}" data-wi="${wi}" type="number" value="${w ?? 0}" /></div>`).join('')}
            <div class="text-muted act-total" data-id="${a._id}">${total} / ${toGo}</div>
            `;
          }).join('')}
        </div>
        <button class="btn btn-light btn-sm" id="addActBtn" style="margin-top:10px;">+ Add activity row</button>
      </div>

      <div class="card">
        <h3>${escapeHtml(m.label)} — Deal Pipeline</h3>
        <div id="pipeGridHost"></div>
      </div>
      `}
    `;

    wireStatic();
    if (m) { wireMonth(m); wirePipelineGrid(m); }
  }

  function wireStatic() {
    container.querySelectorAll('.month-tab-btn').forEach((b) => b.addEventListener('click', () => { activeKey = b.dataset.key; paint(); }));
    container.querySelector('#addMonthBtn').addEventListener('click', addMonth);
    container.querySelector('#exportAllBtn').addEventListener('click', exportAllMonths);
  }

  function addMonth() {
    const input = prompt('New month (YYYY-MM), e.g. 2026-10:', nextMonthKey());
    if (!input) return;
    const m = input.match(/^(\d{4})-(\d{2})$/);
    if (!m) { showToast('Please use YYYY-MM format.'); return; }
    if (d.months.some((mm) => mm.key === input)) { showToast('That month already exists.'); activeKey = input; paint(); return; }
    const label = monthLabel(input);
    const templateActs = d.months.length ? d.months[d.months.length - 1].activities : [];
    d.months.push({
      key: input, label,
      activities: templateActs.length
        ? templateActs.map((a) => ({ _id: uid(), name: a.name, target: a.target, weeks: [0,0,0,0,0] }))
        : [
            { _id: uid(), name: 'APPROACH', target: 0, weeks: [0,0,0,0,0] },
            { _id: uid(), name: 'APP SET', target: 0, weeks: [0,0,0,0,0] },
            { _id: uid(), name: 'PRESENTATION', target: 0, weeks: [0,0,0,0,0] },
            { _id: uid(), name: 'SUBMITTED CASE', target: 0, weeks: [0,0,0,0,0] },
          ],
      pipeline: [],
    });
    d.months.sort((a, b) => a.key.localeCompare(b.key));
    activeKey = input;
    persist(); paint();
  }

  function nextMonthKey() {
    if (!d.months.length) return new Date().toISOString().slice(0, 7);
    const last = d.months[d.months.length - 1].key;
    const [y, mo] = last.split('-').map(Number);
    const nd = new Date(y, mo, 1); // mo is 1-indexed already -> next month
    return nd.toISOString().slice(0, 7);
  }
  function monthLabel(key) {
    const [y, mo] = key.split('-').map(Number);
    return new Date(y, mo - 1, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' }).toUpperCase();
  }

  function wireMonth(m) {
    container.querySelector('#dupMonthBtn').addEventListener('click', () => {
      const key = nextMonthKeyAfter(m.key);
      if (d.months.some((mm) => mm.key === key)) { showToast('Next month already exists.'); return; }
      d.months.push({
        key, label: monthLabel(key),
        activities: m.activities.map((a) => ({ _id: uid(), name: a.name, target: a.target, weeks: [0,0,0,0,0] })),
        pipeline: m.pipeline.map((p) => ({ ...p, _id: uid() })),
      });
      d.months.sort((a, b) => a.key.localeCompare(b.key));
      activeKey = key;
      persist(); paint();
      showToast(`Created ${monthLabel(key)} from ${m.label}.`);
    });

    container.querySelector('#delMonthBtn').addEventListener('click', () => {
      if (!confirm(`Delete ${m.label}? This cannot be undone.`)) return;
      d.months = d.months.filter((mm) => mm.key !== m.key);
      activeKey = d.months.length ? d.months[d.months.length - 1].key : null;
      persist(); paint();
    });

    container.querySelector('#addActBtn').addEventListener('click', () => {
      m.activities.push({ _id: uid(), name: '', target: 0, weeks: [0,0,0,0,0] });
      persist(); paint();
    });

    const updateActTotal = (a) => {
      const cell = container.querySelector(`.act-total[data-id="${a._id}"]`);
      if (!cell) return;
      const total = a.weeks.reduce((s, v) => s + num(v), 0);
      cell.textContent = `${total} / ${num(a.target) - total}`;
    };

    container.querySelectorAll('.act-name').forEach((inp) => {
      inp.addEventListener('input', () => { m.activities.find((a) => a._id === inp.dataset.id).name = inp.value; persist(); });
    });
    container.querySelectorAll('.act-f').forEach((inp) => {
      inp.addEventListener('input', () => {
        const a = m.activities.find((a) => a._id === inp.dataset.id);
        a[inp.dataset.f] = num(inp.value);
        persist(); updateActTotal(a);
      });
    });
    container.querySelectorAll('.act-week').forEach((inp) => {
      inp.addEventListener('input', () => {
        const a = m.activities.find((a) => a._id === inp.dataset.id);
        a.weeks[+inp.dataset.wi] = num(inp.value);
        persist(); updateActTotal(a);
      });
    });

    container.querySelector('#exportMonthBtn').addEventListener('click', () => exportMonth(m));
    container.querySelector('#importMonthInput').addEventListener('change', (e) => importMonth(e, m));
  }

  function nextMonthKeyAfter(key) {
    const [y, mo] = key.split('-').map(Number);
    const nd = new Date(y, mo, 1);
    return nd.toISOString().slice(0, 7);
  }

  function wirePipelineGrid(m) {
    const cols = [
      { key: 'clientName', label: 'Client Name', type: 'text', editable: true, width: '220px' },
      { key: 'product', label: 'Product', type: 'text', editable: true, width: '140px' },
      { key: 'ape', label: 'APE', type: 'number', editable: true, width: '100px' },
      { key: 'chance', label: '% Chance (0-1)', type: 'number', editable: true, width: '110px' },
      { key: 'targetClosingDate', label: 'Target Closing Date', type: 'date', editable: true, width: '150px' },
      { key: 'remarks', label: 'Remarks', type: 'text', editable: true, width: '220px' },
    ];
    createDataGrid({
      container: container.querySelector('#pipeGridHost'),
      columns: cols,
      getRows: () => m.pipeline,
      onCellChange: (row, key, value) => { row[key] = value; persist(); },
      onAddRow: () => { m.pipeline.unshift({ _id: uid(), clientName: '', product: '', ape: 0, chance: 0, targetClosingDate: '', remarks: '' }); persist(); },
      onDeleteRow: (id) => { const i = m.pipeline.findIndex((r) => r._id === id); if (i > -1) m.pipeline.splice(i, 1); persist(); },
      idKey: '_id',
      pageSize: 40,
      emptyLabel: 'No deals in the pipeline for this month yet.',
    });
  }

  function exportMonth(m) {
    const actRows = m.activities.map((a) => ({ activity: a.name, target: a.target, w1: a.weeks[0], w2: a.weeks[1], w3: a.weeks[2], w4: a.weeks[3], w5: a.weeks[4] }));
    downloadWorkbook([
      { name: 'Activities', columns: ACT_COLUMNS, rows: actRows },
      { name: 'Pipeline', columns: PIPE_COLUMNS, rows: m.pipeline },
    ], `Sales Pipeline - ${m.label}.xlsx`);
  }

  async function importMonth(e, m) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      let importedAct = 0, importedPipe = 0;
      if (wb.SheetNames.includes('Activities')) {
        const rows = parseSheetRows(wb, ACT_COLUMNS, 'Activities');
        if (rows.length) {
          m.activities = rows.map((r) => ({ _id: uid(), name: r.activity || '', target: num(r.target), weeks: [num(r.w1), num(r.w2), num(r.w3), num(r.w4), num(r.w5)] }));
          importedAct = rows.length;
        }
      }
      if (wb.SheetNames.includes('Pipeline')) {
        const rows = parseSheetRows(wb, PIPE_COLUMNS, 'Pipeline');
        if (rows.length) {
          m.pipeline = rows.map((r) => ({ _id: uid(), clientName: r.clientName || '', product: r.product || '', ape: num(r.ape), chance: num(r.chance), targetClosingDate: r.targetClosingDate || '', remarks: r.remarks || '' }));
          importedPipe = rows.length;
        }
      }
      persist(); paint();
      showToast(`Imported ${importedAct} activity row(s), ${importedPipe} pipeline row(s).`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  }

  function exportAllMonths() {
    const sheets = d.months.map((mm) => ({ name: mm.label.slice(0, 31), columns: PIPE_COLUMNS, rows: mm.pipeline }));
    if (!sheets.length) { showToast('No months to export yet.'); return; }
    downloadWorkbook(sheets, 'Sales Pipeline - All Months.xlsx');
  }

  paint();
}
