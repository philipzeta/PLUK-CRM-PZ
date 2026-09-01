import { loadTab, markDirty } from '../store.js';
import { uid, daysSince } from '../utils.js';
import { createDataGrid } from '../datagrid.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';
import { showToast } from '../utils.js';

const COLUMNS = [
  { key: 'response', header: 'RESPONSE', aliases: ['RESPONSE'], type: 'select', optKey: 'responseOptions', width: '150px' },
  { key: 'source', header: 'SOURCE', aliases: ['SOURCE'], type: 'select', optKey: 'sourceOptions', width: '150px' },
  { key: 'demographic', header: 'DEMOGRAPHIC: RELATIONSHIP', aliases: ['DEMOGRAPHIC: RELATIONSHIP', 'DEMOGRAPHIC'], type: 'select', optKey: 'demographicOptions', width: '180px' },
  { key: 'name', header: 'NAMES', aliases: ['NAMES', 'Name'], type: 'text', width: '270px' },
  { key: 'status', header: 'STATUS', aliases: ['STATUS'], type: 'text', width: '140px' },
  { key: 'occupation', header: 'OCCUPATION', aliases: ['OCCUPATION'], type: 'text', width: '150px' },
  { key: 'dateAdded', header: 'DATE ADDED', aliases: ['DATE ADDED'], type: 'date', width: '125px' },
  { key: 'lastApproachDate', header: 'LAST APPROACH DATE', aliases: ['LAST APPROACH DATE'], type: 'date', width: '150px' },
  { key: 'approachMethod', header: 'APPROACH METHOD', aliases: ['APPROACH METHOD'], type: 'text', width: '150px' },
  { key: 'remarks', header: 'REMARKS', aliases: ['REMARKS'], type: 'text', width: '230px' },
];

export function makePipelineTabRenderer({ tabId, title, subtitle, emoji }) {
  return async function render(container, { focusRowId } = {}) {
    const d = await loadTab(tabId);
    d.rows.forEach((r) => { if (!r._id) r._id = uid(); });

    container.innerHTML = `
      <div class="tab-header">
        <div>
          <h1 class="tab-title">${emoji} ${title}</h1>
          <p class="tab-subtitle">${subtitle}</p>
        </div>
        <div class="tab-actions">
          <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
          <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
        </div>
      </div>
      <p class="section-note">🔔 Rows highlight amber at 7 days, orange at 14, red at 21+ days since <em>Last Approach Date</em> — these feed the notification bell.</p>
      <div id="gridHost"></div>
    `;

    const gridCols = COLUMNS.map((c) => ({
      key: c.key, label: c.header, type: c.type, editable: true, width: c.width,
      options: c.optKey ? d[c.optKey] : undefined,
    }));
    gridCols.push({
      key: '_days', label: 'Days Since Contact', editable: false, width: '90px',
      computed: (row) => { const ds = daysSince(row.lastApproachDate); return ds === null ? '' : ds + 'd'; },
    });

    const grid = createDataGrid({
      container: container.querySelector('#gridHost'),
      columns: gridCols,
      getRows: () => d.rows,
      onCellChange: (row, key, value) => { row[key] = value; markDirty(tabId); },
      onAddRow: () => {
        const r = { _id: uid(), response: '', source: '', demographic: '', name: '', status: '', occupation: '', dateAdded: '', lastApproachDate: '', approachMethod: '', remarks: '' };
        d.rows.unshift(r); markDirty(tabId);
      },
      onDeleteRow: (id) => { const i = d.rows.findIndex((r) => r._id === id); if (i > -1) d.rows.splice(i, 1); markDirty(tabId); },
      rowClass: (row) => {
        const ds = daysSince(row.lastApproachDate);
        if (ds === null || row.response === 'CLIENT') return '';
        if (ds >= 21) return 'overdue-21';
        if (ds >= 14) return 'overdue-14';
        if (ds >= 7) return 'overdue-7';
        return '';
      },
      idKey: '_id',
      pageSize: 60,
      emptyLabel: `No ${title.toLowerCase()} rows yet. Add one, or import your list from Excel.`,
    });

    if (focusRowId) {
      // best-effort: try to surface the row by scrolling the search to its name
      const row = d.rows.find((r) => r._id === focusRowId);
      if (row) {
        showToast(`Showing ${title} — search "${row.name}" to jump to the row.`);
      }
    }

    container.querySelector('#exportBtn').addEventListener('click', () => {
      downloadSingleSheet(tabId.toUpperCase(), COLUMNS, d.rows, `${title}.xlsx`);
    });

    container.querySelector('#importInput').addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        const wb = await readWorkbook(file);
        const rows = parseSheetRows(wb, COLUMNS, wb.SheetNames.find((n) => n.toUpperCase() === tabId.toUpperCase()) || undefined);
        rows.forEach((r) => (r._id = uid()));
        if (!confirm(`Import ${rows.length} row(s)? This will replace the current ${title} list.`)) return;
        d.rows = rows;
        markDirty(tabId);
        grid.resetPage(); grid.refresh();
        showToast(`Imported ${rows.length} row(s) from Excel.`);
      } catch (err) {
        console.error(err);
        showToast('Could not read that file.');
      } finally {
        e.target.value = '';
      }
    });
  };
}
