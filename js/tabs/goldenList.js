import { loadTab, markDirty } from '../store.js';
import { uid, num, showToast } from '../utils.js';
import { createDataGrid } from '../datagrid.js';
import { downloadWorkbook, readWorkbook, parseSheetRows } from '../excel.js';

const SALES_COLUMNS = [
  { key: 'prospectName', header: 'Prospect Name', aliases: ['Prospect Name', 'Client Name', 'Name'] },
  { key: 'product', header: 'Product', aliases: ['Product'] },
  { key: 'ape', header: 'APE', aliases: ['APE'] },
  { key: 'chancePct', header: '% Chance (1-100)', aliases: ['% Chance (1-100)', '% Chance of Closing', 'Chance'] },
  { key: 'targetClosingDate', header: 'Target Closing Date', aliases: ['Target Closing Date'], type: 'date' },
  { key: 'remarks', header: 'Remarks', aliases: ['Remarks'] },
];

const RECRUITMENT_COLUMNS = [
  { key: 'prospectName', header: 'Prospect Name', aliases: ['Prospect Name', 'Name'] },
  { key: 'source', header: 'Source', aliases: ['Source'] },
  { key: 'demographic', header: 'Demographic', aliases: ['Demographic'] },
  { key: 'lastApproachDate', header: 'Last Approach Date', aliases: ['Last Approach Date'], type: 'date' },
  { key: 'birthdate', header: 'Birthdate', aliases: ['Birthdate'], type: 'date' },
  { key: 'email', header: 'Email', aliases: ['Email'] },
  { key: 'contactNo', header: 'Contact No.', aliases: ['Contact No.', 'Contact No', 'Mobile No.'] },
  { key: 'bybTtSchedule', header: 'BYB/TT Schedule', aliases: ['BYB/TT Schedule', 'BYB Schedule'], type: 'date' },
  { key: 'attendedBybTt', header: 'Attended BYB/TT', aliases: ['Attended BYB/TT', 'Attended'], type: 'checkbox' },
  { key: 'remarks', header: 'Remarks', aliases: ['Remarks'] },
];

function emptySalesRow() {
  return { _id: uid(), prospectName: '', product: '', ape: null, chancePct: null, targetClosingDate: '', remarks: '' };
}
function emptyRecruitmentRow() {
  return {
    _id: uid(), prospectName: '', source: '', demographic: '', lastApproachDate: '', birthdate: '',
    email: '', contactNo: '', bybTtSchedule: '', attendedBybTt: false, remarks: '',
  };
}

export async function render(container) {
  const d = await loadTab('golden-list');
  if (!d.salesRows) d.salesRows = [];
  if (!d.recruitmentRows) d.recruitmentRows = [];
  d.salesRows.forEach((r) => { if (!r._id) r._id = uid(); });
  d.recruitmentRows.forEach((r) => { if (!r._id) r._id = uid(); });

  // One-time migration: the Deal Pipeline section that used to live inside
  // Sales Pipeline (now "Weekly Activity") moved here instead. Pull in
  // anything already entered there, across every month, so it isn't lost.
  if (!d._migratedDealPipeline) {
    d._migratedDealPipeline = true;
    const spd = await loadTab('sales-pipeline');
    let migrated = 0;
    (spd.months || []).forEach((m) => {
      (m.pipeline || []).forEach((p) => {
        if (!p.clientName || !String(p.clientName).trim()) return;
        const chanceRaw = num(p.chance, null);
        let chancePct = chanceRaw === null ? null : Math.round(chanceRaw * 100);
        if (chancePct !== null) chancePct = Math.max(0, Math.min(100, chancePct));
        d.salesRows.push({
          _id: uid(),
          prospectName: p.clientName,
          product: p.product || '',
          ape: p.ape ?? null,
          chancePct,
          targetClosingDate: p.targetClosingDate || '',
          remarks: p.remarks || '',
        });
        migrated++;
      });
    });
    markDirty('golden-list');
    if (migrated) showToast(`Brought ${migrated} deal(s) over from the old Sales Pipeline into the Golden List.`);
  }

  function persist() { markDirty('golden-list'); }

  container.innerHTML = `
    <div class="tab-header">
      <div>
        <h1 class="tab-title">🥇 Golden List</h1>
        <p class="tab-subtitle">Your shortlist of hottest prospects — for sales and for recruitment — kept separate from the full pipelines.</p>
      </div>
      <div class="tab-actions">
        <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
        <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
      </div>
    </div>

    <div class="card">
      <h3>💰 Sales Pipeline</h3>
      <div id="salesGridHost"></div>
    </div>

    <div class="card" style="margin-top:16px;">
      <h3>🤝 Recruitment Pipeline</h3>
      <div id="recruitmentGridHost"></div>
    </div>
  `;

  createDataGrid({
    container: container.querySelector('#salesGridHost'),
    columns: [
      { key: 'prospectName', label: 'Prospect Name', type: 'text', editable: true, width: '220px' },
      { key: 'product', label: 'Product', type: 'text', editable: true, width: '150px' },
      { key: 'ape', label: 'APE', type: 'money', editable: true, width: '120px' },
      { key: 'chancePct', label: '% Chance (1-100)', type: 'number', editable: true, width: '130px', min: 1, max: 100 },
      { key: 'targetClosingDate', label: 'Target Closing Date', type: 'date', editable: true, width: '150px' },
      { key: 'remarks', label: 'Remarks', type: 'text', editable: true, width: '240px' },
    ],
    getRows: () => d.salesRows,
    onCellChange: (row, key, value) => { row[key] = value; persist(); },
    onAddRow: () => { d.salesRows.unshift(emptySalesRow()); persist(); },
    onDeleteRow: (id) => { const i = d.salesRows.findIndex((r) => r._id === id); if (i > -1) d.salesRows.splice(i, 1); persist(); },
    idKey: '_id',
    pageSize: 60,
    emptyLabel: 'No sales prospects on the Golden List yet. Add one, or import from Excel.',
  });

  createDataGrid({
    container: container.querySelector('#recruitmentGridHost'),
    columns: [
      { key: 'prospectName', label: 'Prospect Name', type: 'text', editable: true, width: '210px' },
      { key: 'source', label: 'Source', type: 'select', editable: true, width: '150px', options: d.sourceOptions },
      { key: 'demographic', label: 'Demographic', type: 'select', editable: true, width: '170px', options: d.demographicOptions },
      { key: 'lastApproachDate', label: 'Last Approach Date', type: 'date', editable: true, width: '150px' },
      { key: 'birthdate', label: 'Birthdate', type: 'date', editable: true, width: '125px' },
      { key: 'email', label: 'Email', type: 'text', editable: true, width: '190px' },
      { key: 'contactNo', label: 'Contact No.', type: 'text', editable: true, width: '130px' },
      { key: 'bybTtSchedule', label: 'BYB/TT Schedule', type: 'date', editable: true, width: '145px' },
      { key: 'attendedBybTt', label: 'Attended BYB/TT', type: 'checkbox', editable: true, width: '120px' },
      { key: 'remarks', label: 'Remarks', type: 'text', editable: true, width: '220px' },
    ],
    getRows: () => d.recruitmentRows,
    onCellChange: (row, key, value) => { row[key] = value; persist(); },
    onAddRow: () => { d.recruitmentRows.unshift(emptyRecruitmentRow()); persist(); },
    onDeleteRow: (id) => { const i = d.recruitmentRows.findIndex((r) => r._id === id); if (i > -1) d.recruitmentRows.splice(i, 1); persist(); },
    idKey: '_id',
    pageSize: 60,
    emptyLabel: 'No recruitment prospects on the Golden List yet. Add one, or import from Excel.',
  });

  container.querySelector('#exportBtn').addEventListener('click', () => {
    downloadWorkbook([
      { name: 'GOLDEN LIST - SALES', columns: SALES_COLUMNS, rows: d.salesRows },
      { name: 'GOLDEN LIST - RECRUITMENT', columns: RECRUITMENT_COLUMNS, rows: d.recruitmentRows },
    ], 'Golden List.xlsx');
  });

  container.querySelector('#importInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      let newSales = null, newRecruitment = null;
      if (wb.SheetNames.includes('GOLDEN LIST - SALES')) {
        const rows = parseSheetRows(wb, SALES_COLUMNS, 'GOLDEN LIST - SALES');
        if (rows.length) {
          newSales = rows.map((r) => ({
            _id: uid(), prospectName: r.prospectName || '', product: r.product || '',
            ape: r.ape ? num(r.ape) : null, chancePct: r.chancePct ? num(r.chancePct) : null,
            targetClosingDate: r.targetClosingDate || '', remarks: r.remarks || '',
          }));
        }
      }
      if (wb.SheetNames.includes('GOLDEN LIST - RECRUITMENT')) {
        const rows = parseSheetRows(wb, RECRUITMENT_COLUMNS, 'GOLDEN LIST - RECRUITMENT');
        if (rows.length) {
          newRecruitment = rows.map((r) => ({
            _id: uid(), prospectName: r.prospectName || '', source: r.source || '', demographic: r.demographic || '',
            lastApproachDate: r.lastApproachDate || '', birthdate: r.birthdate || '', email: r.email || '',
            contactNo: r.contactNo || '', bybTtSchedule: r.bybTtSchedule || '',
            attendedBybTt: r.attendedBybTt === true || r.attendedBybTt === 'TRUE' || r.attendedBybTt === 'true' || r.attendedBybTt === 1 || r.attendedBybTt === '1',
            remarks: r.remarks || '',
          }));
        }
      }
      if (!newSales && !newRecruitment) { showToast('No matching sheets found in that file.'); return; }
      if (!confirm(`Import ${newSales ? newSales.length : 0} sales row(s) and ${newRecruitment ? newRecruitment.length : 0} recruitment row(s)? This will replace the current Golden List.`)) return;
      if (newSales) d.salesRows = newSales;
      if (newRecruitment) d.recruitmentRows = newRecruitment;
      persist();
      render(container);
      showToast(`Imported ${newSales ? newSales.length : 0} sales row(s), ${newRecruitment ? newRecruitment.length : 0} recruitment row(s).`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  });
}
