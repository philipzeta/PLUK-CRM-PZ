import { loadTab, markDirty } from '../store.js';
import { uid, escapeHtml } from '../utils.js';
import { createDataGrid } from '../datagrid.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';
import { showToast } from '../utils.js';

const COLUMNS = [
  { key: 'dateAdded', header: 'Date Added', aliases: ['Date Added', 'Date'], type: 'date', width: '125px' },
  { key: 'recruitName', header: 'Recruit Name', aliases: ['Recruit Name', 'Name'], type: 'text', width: '270px' },
  { key: 'source', header: 'SOURCE', aliases: ['SOURCE', 'Source'], type: 'select', width: '150px' },
  { key: 'birthdate', header: 'BIRTHDATE', aliases: ['BIRTHDATE', 'Birthdate'], type: 'date', width: '125px' },
  { key: 'age', header: 'AGE', aliases: ['AGE', 'Age'], type: 'number', width: '80px' },
  { key: 'collegeCourse', header: 'COLLEGE COURSE', aliases: ['COLLEGE COURSE'], type: 'text', width: '190px' },
  { key: 'areaOfResidence', header: 'AREA OF RESIDENCE', aliases: ['AREA OF RESIDENCE (Must be within NCR)', 'AREA OF RESIDENCE'], type: 'text', width: '200px' },
  { key: 'mobileNo', header: 'MOBILE NO.', aliases: ['MOBILE NO.', 'Mobile No', 'Mobile'], type: 'text', width: '140px' },
  { key: 'plukGmail', header: 'PLUK GMAIL', aliases: ['PLUK GMAIL', 'Gmail'], type: 'text', width: '190px' },
  { key: 'bybSchedule', header: 'BYB SCHEDULE', aliases: ['BYB SCHEDULE'], type: 'date', width: '135px' },
  { key: 'attended', header: 'ATTENDED?', aliases: ['ATTENDED / DID NOT ATTEND', 'ATTENDED?'], type: 'checkbox', width: '90px' },
  { key: 'paidIcExam', header: 'PAID IC EXAM?', aliases: ['PAID IC EXAM?'], type: 'text', width: '140px' },
  { key: 'ilt1', header: 'ILT 1', aliases: ['ILT 1'], type: 'date', width: '115px' },
  { key: 'ilt2', header: 'ILT 2', aliases: ['ILT 2'], type: 'date', width: '115px' },
  { key: 'orientationDate', header: 'ORIENTATION DATE', aliases: ['ORIENTATION DATE'], type: 'date', width: '150px' },
  { key: 'dateCoded', header: 'DATE CODED', aliases: ['DATE CODED'], type: 'date', width: '125px' },
  { key: 'remarks', header: 'REMARKS', aliases: ['REMARKS'], type: 'text', width: '220px' },
  { key: 'tatBybToCoding', header: 'TAT (BYB→Coding, days)', aliases: ['TAT from BYB date to Coding', 'TAT'], type: 'number', width: '110px' },
];

export async function render(container, { focusRowId } = {}) {
  const d = await loadTab('recruit-prospect-list');
  d.rows.forEach((r) => { if (!r._id) r._id = uid(); });

  container.innerHTML = `
    <div class="tab-header">
      <div>
        <h1 class="tab-title">🧑‍🤝‍🧑 Recruit Prospect List</h1>
        <p class="tab-subtitle">Future teammates pipeline, from BYB invite through coding.</p>
      </div>
      <div class="tab-actions">
        <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
        <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
      </div>
    </div>
    <div class="card" style="padding:10px 14px;">
      <strong style="font-size:12.5px">Source options:</strong>
      <span class="text-muted" style="font-size:12.5px">${d.sourceLegend.map((s) => escapeHtml(s.source)).join(' · ')}</span>
    </div>
    <div id="gridHost"></div>
  `;

  const gridCols = COLUMNS.map((c) => ({
    key: c.key, label: c.header, type: c.type, editable: true, width: c.width,
    options: c.key === 'source' ? d.sourceOptions : undefined,
  }));

  const grid = createDataGrid({
    container: container.querySelector('#gridHost'),
    columns: gridCols,
    getRows: () => d.rows,
    onCellChange: (row, key, value) => { row[key] = value; markDirty('recruit-prospect-list'); },
    onAddRow: () => { const r = { _id: uid() }; COLUMNS.forEach((c) => (r[c.key] = c.type === 'checkbox' ? false : '')); d.rows.unshift(r); markDirty('recruit-prospect-list'); },
    onDeleteRow: (id) => { const i = d.rows.findIndex((r) => r._id === id); if (i > -1) d.rows.splice(i, 1); markDirty('recruit-prospect-list'); },
    idKey: '_id',
    emptyLabel: 'No recruit prospects yet. Add one, or import your list from Excel.',
  });

  container.querySelector('#exportBtn').addEventListener('click', () => {
    downloadSingleSheet('RECRUIT PROSPECT LIST', COLUMNS, d.rows, 'Recruit Prospect List.xlsx');
  });

  container.querySelector('#importInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const rows = parseSheetRows(wb, COLUMNS, 'RECRUIT PROSPECT LIST');
      rows.forEach((r) => (r._id = uid()));
      if (!confirm(`Import ${rows.length} row(s)? This will replace the current list in this tab.`)) return;
      d.rows = rows;
      markDirty('recruit-prospect-list');
      grid.resetPage(); grid.refresh();
      showToast(`Imported ${rows.length} row(s) from Excel.`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  });

  if (focusRowId) {
    const idx = d.rows.findIndex((r) => r._id === focusRowId);
    if (idx > -1) showToast('Jumped to Recruitment Prospect List — search to locate the row.');
  }
}
