import { loadTab, markDirty } from '../store.js';
import { uid, daysSince, todayISO, showToast } from '../utils.js';
import { createDataGrid } from '../datagrid.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';

const COLUMNS = [
  { key: 'dateAdded', header: 'Date Added', aliases: ['Date Added', 'Date'], type: 'date', width: '125px' },
  { key: 'task', header: 'Task', aliases: ['Task', 'To-Do', 'Item'], type: 'text', width: '320px' },
  { key: 'category', header: 'Category', aliases: ['Category'], type: 'select', width: '150px' },
  { key: 'priority', header: 'Priority', aliases: ['Priority'], type: 'select', width: '120px' },
  { key: 'dueDate', header: 'Due Date', aliases: ['Due Date', 'Due'], type: 'date', width: '125px' },
  { key: 'done', header: 'Done', aliases: ['Done', 'Status', 'Complete'], type: 'checkbox', width: '80px' },
  { key: 'notes', header: 'Notes', aliases: ['Notes', 'Remarks'], type: 'text', width: '260px' },
];

export async function render(container) {
  const d = await loadTab('todo');
  if (!d.categoryOptions) d.categoryOptions = ['Selling', 'Recruitment', 'Admin', 'Follow-up', 'Personal', 'Other'];
  if (!d.priorityOptions) d.priorityOptions = ['Low', 'Medium', 'High'];
  if (!d.rows) d.rows = [];
  d.rows.forEach((r) => { if (!r._id) r._id = uid(); });

  container.innerHTML = `
    <div class="tab-header">
      <div>
        <h1 class="tab-title">✅ To-Do List</h1>
        <p class="tab-subtitle">Anything you need to track — sales, recruitment, or personal — kept in this same file.</p>
      </div>
      <div class="tab-actions">
        <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
        <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
      </div>
    </div>
    <div style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;">
      <div class="stat" style="min-width:120px;"><div class="stat-label">Open</div><div class="stat-value" id="statOpen">0</div></div>
      <div class="stat" style="min-width:120px;"><div class="stat-label">Overdue</div><div class="stat-value" id="statOverdue">0</div></div>
      <div class="stat" style="min-width:120px;"><div class="stat-label">Done</div><div class="stat-value" id="statDone">0</div></div>
    </div>
    <p class="section-note">🔔 Rows highlight amber when due within 2 days, orange when due today, red once overdue — and dim with a strikethrough once checked Done.</p>
    <div id="gridHost"></div>
  `;

  const gridCols = COLUMNS.map((c) => ({
    key: c.key, label: c.header, type: c.type, editable: true, width: c.width,
    options: c.key === 'category' ? d.categoryOptions : c.key === 'priority' ? d.priorityOptions : undefined,
  }));

  const grid = createDataGrid({
    container: container.querySelector('#gridHost'),
    columns: gridCols,
    getRows: () => d.rows,
    onCellChange: (row, key, value) => {
      row[key] = value;
      markDirty('todo');
      // 'done' and 'dueDate' fire on discrete change events (not every
      // keystroke), so it's safe to refresh the grid to update row
      // highlighting/strikethrough immediately without breaking typing.
      if (key === 'done' || key === 'dueDate') refreshStats(true);
      else refreshStats(false);
    },
    onAddRow: () => {
      const r = { _id: uid(), dateAdded: todayISO() };
      COLUMNS.forEach((c) => { if (!(c.key in r)) r[c.key] = c.type === 'checkbox' ? false : ''; });
      d.rows.unshift(r);
      markDirty('todo');
      refreshStats(false);
    },
    onDeleteRow: (id) => {
      const i = d.rows.findIndex((r) => r._id === id);
      if (i > -1) d.rows.splice(i, 1);
      markDirty('todo');
      refreshStats(false);
    },
    rowClass: (row) => {
      if (row.done) return 'todo-done';
      const ds = daysSince(row.dueDate);
      if (ds === null) return '';
      if (ds > 0) return 'overdue-21';
      if (ds === 0) return 'overdue-14';
      if (ds >= -2) return 'overdue-7';
      return '';
    },
    idKey: '_id',
    pageSize: 60,
    emptyLabel: 'No to-dos yet. Add one, or import a list from Excel.',
  });

  function refreshStats(alsoRefreshGrid) {
    const open = d.rows.filter((r) => !r.done).length;
    const overdue = d.rows.filter((r) => !r.done && daysSince(r.dueDate) > 0).length;
    container.querySelector('#statOpen').textContent = open;
    container.querySelector('#statOverdue').textContent = overdue;
    container.querySelector('#statDone').textContent = d.rows.length - open;
    if (alsoRefreshGrid) grid.refresh();
  }
  refreshStats(false);

  container.querySelector('#exportBtn').addEventListener('click', () => {
    downloadSingleSheet('TODO', COLUMNS, d.rows, 'To-Do List.xlsx');
  });

  container.querySelector('#importInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const rows = parseSheetRows(wb, COLUMNS, 'TODO');
      rows.forEach((r) => {
        r._id = uid();
        r.done = r.done === true || r.done === 'TRUE' || r.done === 'true' || r.done === 1 || r.done === '1';
      });
      if (!confirm(`Import ${rows.length} row(s)? This will replace the current to-do list.`)) return;
      d.rows = rows;
      markDirty('todo');
      grid.resetPage(); grid.refresh();
      refreshStats(false);
      showToast(`Imported ${rows.length} row(s) from Excel.`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  });
}
