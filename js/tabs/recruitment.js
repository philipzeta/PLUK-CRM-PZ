import { loadTab, markDirty } from '../store.js';
import { dbGet } from '../db.js';
import { uid, daysSince, escapeHtml, showToast } from '../utils.js';
import { createDataGrid } from '../datagrid.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';

// Merged tab: covers the whole recruiting journey in one row per person —
// first approach (source/response/last-approach-date, which feeds the
// 7/14/21-day follow-up alerts) through BYB, ILT, orientation and coding.
// This used to be two separate tabs ("Recruit Prospect List" and
// "Recruitment"); they're combined here to avoid entering the same person
// twice. The tab id stays 'recruitment' so the notification bell (which
// hardcodes this id) keeps working unchanged.

const COLUMNS = [
  { key: 'dateAdded', header: 'Date Added', aliases: ['Date Added', 'Date'], type: 'date', width: '115px' },
  { key: 'name', header: 'Name', aliases: ['Name', 'NAMES', 'Recruit Name'], type: 'text', width: '270px' },
  { key: 'source', header: 'Source', aliases: ['SOURCE', 'Source'], type: 'select', width: '150px' },
  { key: 'demographic', header: 'Demographic: Relationship', aliases: ['DEMOGRAPHIC: RELATIONSHIP', 'DEMOGRAPHIC', 'Demographic'], type: 'select', width: '170px' },
  { key: 'occupation', header: 'Occupation', aliases: ['OCCUPATION', 'Occupation'], type: 'text', width: '140px' },
  { key: 'approachMethod', header: 'Approach Method', aliases: ['APPROACH METHOD', 'Approach Method'], type: 'text', width: '140px' },
  { key: 'lastApproachDate', header: 'Last Approach Date', aliases: ['LAST APPROACH DATE', 'Last Approach Date'], type: 'date', width: '145px' },
  { key: 'response', header: 'Response', aliases: ['RESPONSE', 'Response'], type: 'select', width: '140px' },
  { key: 'birthdate', header: 'Birthdate', aliases: ['BIRTHDATE', 'Birthdate'], type: 'date', width: '120px' },
  { key: 'age', header: 'Age', aliases: ['AGE', 'Age'], type: 'number', width: '70px' },
  { key: 'collegeCourse', header: 'College Course', aliases: ['COLLEGE COURSE', 'College Course'], type: 'text', width: '170px' },
  { key: 'areaOfResidence', header: 'Area of Residence', aliases: ['AREA OF RESIDENCE (Must be within NCR)', 'AREA OF RESIDENCE', 'Area of Residence'], type: 'text', width: '190px' },
  { key: 'mobileNo', header: 'Mobile No.', aliases: ['MOBILE NO.', 'Mobile No', 'Mobile'], type: 'text', width: '130px' },
  { key: 'plukGmail', header: 'PRU Gmail', aliases: ['PLUK GMAIL', 'PRU GMAIL', 'Gmail'], type: 'text', width: '180px' },
  { key: 'bybSchedule', header: 'BYB Schedule', aliases: ['BYB SCHEDULE', 'BYB Schedule'], type: 'date', width: '130px' },
  { key: 'attended', header: 'Attended?', aliases: ['ATTENDED / DID NOT ATTEND', 'ATTENDED?', 'Attended'], type: 'checkbox', width: '85px' },
  { key: 'paidIcExam', header: 'Paid IC Exam?', aliases: ['PAID IC EXAM?', 'Paid IC Exam'], type: 'text', width: '120px' },
  { key: 'ilt1', header: 'ILT 1', aliases: ['ILT 1'], type: 'date', width: '110px' },
  { key: 'ilt2', header: 'ILT 2', aliases: ['ILT 2'], type: 'date', width: '110px' },
  { key: 'orientationDate', header: 'Orientation Date', aliases: ['ORIENTATION DATE', 'Orientation Date'], type: 'date', width: '140px' },
  { key: 'dateCoded', header: 'Date Coded', aliases: ['DATE CODED', 'Date Coded'], type: 'date', width: '120px' },
  { key: 'remarks', header: 'Remarks', aliases: ['REMARKS', 'Remarks'], type: 'text', width: '200px' },
];

// Fields copied in from a legacy "Recruit Prospect List" row during one-time migration.
const LEGACY_COPY_FIELDS = ['source', 'birthdate', 'age', 'collegeCourse', 'areaOfResidence', 'mobileNo', 'plukGmail', 'bybSchedule', 'attended', 'paidIcExam', 'ilt1', 'ilt2', 'orientationDate', 'dateCoded'];

// Returns null if migration already ran in a previous session (no-op this
// time). Otherwise always returns { mergedCount, addedCount } — even if both
// are 0 — as the caller's signal to persist the migrated-flag (and any
// options/legend merged in below) so this doesn't re-run every load.
async function migrateLegacyProspectList(d) {
  if (d._migratedFromProspectList) return null;
  d._migratedFromProspectList = true; // set first so this only ever runs once, even if something below throws
  let legacy;
  try { legacy = await dbGet('recruit-prospect-list'); } catch (e) { legacy = null; }
  if (!legacy) return { mergedCount: 0, addedCount: 0 };

  // Bring over source options/legend this tab didn't already have, so values
  // migrated from the old tab still show correctly in the Source dropdown.
  if (Array.isArray(legacy.sourceOptions)) {
    const existing = new Set(d.sourceOptions || (d.sourceOptions = []));
    legacy.sourceOptions.forEach((o) => { if (o && !existing.has(o)) { d.sourceOptions.push(o); existing.add(o); } });
  }
  if (!d.sourceLegend && Array.isArray(legacy.sourceLegend)) {
    d.sourceLegend = legacy.sourceLegend.filter((l) => l && l.source && l.source !== 'SOURCE');
  }

  if (!Array.isArray(legacy.rows) || !legacy.rows.length) return { mergedCount: 0, addedCount: 0 };

  let mergedCount = 0, addedCount = 0;
  legacy.rows.forEach((lr) => {
    const legacyName = String(lr.recruitName || '').trim();
    if (!legacyName) return;
    const match = d.rows.find((r) => String(r.name || '').trim().toLowerCase() === legacyName.toLowerCase());
    if (match) {
      LEGACY_COPY_FIELDS.forEach((f) => {
        const empty = match[f] === undefined || match[f] === null || match[f] === '';
        const legacyHasValue = lr[f] !== undefined && lr[f] !== null && lr[f] !== '';
        if (empty && legacyHasValue) match[f] = lr[f];
      });
      if (lr.remarks && String(lr.remarks).trim() && String(lr.remarks).trim() !== String(match.remarks || '').trim()) {
        match.remarks = [match.remarks, lr.remarks].filter(Boolean).join(' | ');
      }
      if (!match.dateAdded && lr.dateAdded) match.dateAdded = lr.dateAdded;
      mergedCount++;
    } else {
      const row = { _id: uid(), name: legacyName, dateAdded: lr.dateAdded || '', remarks: lr.remarks || '' };
      LEGACY_COPY_FIELDS.forEach((f) => { row[f] = lr[f] !== undefined && lr[f] !== null ? lr[f] : (f === 'attended' ? false : ''); });
      d.rows.push(row);
      addedCount++;
    }
  });
  return { mergedCount, addedCount };
}

export async function render(container, { focusRowId } = {}) {
  const d = await loadTab('recruitment');
  if (!d.rows) d.rows = [];
  if (!d.responseOptions) d.responseOptions = [];
  if (!d.sourceOptions) d.sourceOptions = [];
  if (!d.demographicOptions) d.demographicOptions = [];
  d.rows.forEach((r) => { if (!r._id) r._id = uid(); });

  const migration = await migrateLegacyProspectList(d);
  if (migration) markDirty('recruitment');

  container.innerHTML = `
    <div class="tab-header">
      <div>
        <h1 class="tab-title">👥 Recruitment</h1>
        <p class="tab-subtitle">The full recruiting journey in one place — first approach through BYB, ILT, orientation and coding.</p>
      </div>
      <div class="tab-actions">
        <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
        <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
      </div>
    </div>
    <p class="section-note">🔔 Rows highlight amber → orange → red at 7 / 14 / 21 days since <em>Last Approach Date</em> — these feed the notification bell. Set <em>Response</em> to CLIENT once someone's coded to stop their alerts.</p>
    ${d.sourceLegend && d.sourceLegend.length ? `
    <div class="card" style="padding:10px 14px;">
      <strong style="font-size:12.5px">Source options:</strong>
      <span class="text-muted" style="font-size:12.5px">${d.sourceLegend.map((s) => escapeHtml(s.source)).join(' · ')}</span>
    </div>` : ''}
    <div id="gridHost"></div>
  `;

  const gridCols = COLUMNS.map((c) => ({
    key: c.key, label: c.header, type: c.type, editable: true, width: c.width,
    options: c.key === 'response' ? d.responseOptions : c.key === 'source' ? d.sourceOptions : c.key === 'demographic' ? d.demographicOptions : undefined,
  }));
  gridCols.push({
    key: '_days', label: 'Days Since Contact', editable: false, width: '90px',
    computed: (row) => { const ds = daysSince(row.lastApproachDate); return ds === null ? '' : ds + 'd'; },
  });
  gridCols.push({
    key: '_bybToCoded', label: 'Days: BYB → Coded', editable: false, width: '100px',
    computed: (row) => {
      if (!row.bybSchedule || !row.dateCoded) return '';
      const diff = daysSince(row.bybSchedule) - daysSince(row.dateCoded);
      return diff >= 0 ? diff + 'd' : '';
    },
  });

  const grid = createDataGrid({
    container: container.querySelector('#gridHost'),
    columns: gridCols,
    getRows: () => d.rows,
    onCellChange: (row, key, value) => { row[key] = value; markDirty('recruitment'); },
    onAddRow: () => {
      const r = { _id: uid() };
      COLUMNS.forEach((c) => (r[c.key] = c.type === 'checkbox' ? false : ''));
      d.rows.unshift(r);
      markDirty('recruitment');
    },
    onDeleteRow: (id) => { const i = d.rows.findIndex((r) => r._id === id); if (i > -1) d.rows.splice(i, 1); markDirty('recruitment'); },
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
    emptyLabel: 'No recruits yet. Add one, or import your list from Excel.',
  });

  if (migration && (migration.mergedCount || migration.addedCount)) {
    grid.refresh();
    showToast(`Combined Recruit Prospect List into Recruitment: merged ${migration.mergedCount}, added ${migration.addedCount} row(s).`, 5000);
  }

  if (focusRowId) {
    const row = d.rows.find((r) => r._id === focusRowId);
    if (row) showToast(`Showing Recruitment — search "${row.name}" to jump to the row.`);
  }

  container.querySelector('#exportBtn').addEventListener('click', () => {
    downloadSingleSheet('RECRUITMENT', COLUMNS, d.rows, 'Recruitment.xlsx');
  });

  container.querySelector('#importInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const rows = parseSheetRows(wb, COLUMNS, wb.SheetNames.find((n) => n.toUpperCase() === 'RECRUITMENT') || undefined);
      rows.forEach((r) => (r._id = uid()));
      if (!confirm(`Import ${rows.length} row(s)? This will replace the current Recruitment list.`)) return;
      d.rows = rows;
      markDirty('recruitment');
      grid.resetPage(); grid.refresh();
      showToast(`Imported ${rows.length} row(s) from Excel.`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  });
}
