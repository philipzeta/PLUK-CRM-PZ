import { loadTab, markDirty } from '../store.js';
import { num, fmtMoney, fmtInt, escapeHtml, showToast, withScrollPreserved, wireMoneyInput } from '../utils.js';
import { computeAgape, computeActionPlan } from '../formulas.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';

function editableMonthRow(label, values, key, money = false) {
  const cells = values
    .map(
      (v, i) =>
        `<td class="num"><input class="cell-input${money ? ' money-input' : ''}" type="${money ? 'text' : 'number'}" ${money ? 'inputmode="decimal"' : 'step="any"'} data-i="${i}" data-key="${key}" value="${v ?? 0}" style="width:82px;text-align:right" /></td>`
    )
    .join('');
  return `<tr><td>${escapeHtml(label)}</td>${cells}</tr>`;
}

function readonlyMonthRow(label, values, rowId, fmt = fmtMoney) {
  const cells = values.map((v, i) => `<td class="num" data-i="${i}">${fmt(v)}</td>`).join('');
  return `<tr data-row="${rowId}"><td>${escapeHtml(label)}</td>${cells}</tr>`;
}

export async function render(container) {
  const agape = await loadTab('agape');
  const ap = await loadTab('action-plan');

  function paint() { withScrollPreserved(container, paintInner); }

  function paintInner() {
    const ac = computeAgape(agape);
    const c = computeActionPlan(ap, ac.annualPremiumTarget);
    const months = ap.months;
    const monthHead = months.map((m) => `<th class="num">${escapeHtml(m.slice(0, 3))}</th>`).join('');

    container.innerHTML = `
      <div class="tab-header">
        <div>
          <h1 class="tab-title">🗓️ Action Plan</h1>
          <p class="tab-subtitle">Monthly sales & recruitment targets. "Target APE" flows automatically from your AGAPE annual premium target.</p>
        </div>
        <div class="tab-actions">
          <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
          <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
        </div>
      </div>

      <div class="card">
        <h3>Sales</h3>
        <div class="dg-scroll" style="max-height:none">
          <table class="dg-table"><thead><tr><th style="position:sticky;left:0;background:#faf7fb">Metric</th>${monthHead}</tr></thead>
          <tbody id="salesRows">
            ${readonlyMonthRow('Target APE', c.targetApe, 'targetApe')}
            ${editableMonthRow('Average case size', ap.sales.averageCaseSize, 'averageCaseSize', true)}
            ${editableMonthRow('Number of cases', ap.sales.numberOfCases, 'numberOfCases')}
            ${readonlyMonthRow('Actual APE', c.actualApe, 'actualApe')}
            ${readonlyMonthRow('Excess / Lacking', c.excess, 'excess')}
            ${readonlyMonthRow('Income (35%)', c.income, 'income')}
          </tbody></table>
        </div>
      </div>

      <div class="card">
        <h3>Recruitment</h3>
        <div style="max-width:260px;margin-bottom:10px;">
          <label class="text-muted">Beginning manpower (January)</label>
          <input type="number" id="beginMP" value="${ap.recruitment.beginningManpowerJan ?? 0}" />
        </div>
        <div class="dg-scroll" style="max-height:none">
          <table class="dg-table"><thead><tr><th style="position:sticky;left:0;background:#faf7fb">Metric</th>${monthHead}</tr></thead>
          <tbody id="recruitRows">
            ${readonlyMonthRow('Beginning manpower', c.beginMP, 'beginMP', fmtInt)}
            ${editableMonthRow('Target new recruits', ap.recruitment.targetNewRecruits, 'targetNewRecruits')}
            ${editableMonthRow('Provision for attrition', ap.recruitment.provisionForAttrition, 'provisionForAttrition')}
            ${readonlyMonthRow('Ending manpower', c.endMP, 'endMP', fmtInt)}
          </tbody></table>
        </div>
      </div>

      <div class="card">
        <h3>Planning my sales activities</h3>
        <p class="section-note">${escapeHtml(ap.successFormulaSales || '')}</p>
        <div class="dg-scroll" style="max-height:none">
          <table class="dg-table"><thead><tr><th style="position:sticky;left:0;background:#faf7fb">Metric</th>${monthHead}</tr></thead>
          <tbody id="salesActRows">
            ${editableMonthRow('Target cases for the month', ap.salesActivity.targetCasesForMonth, 'targetCasesForMonth')}
            ${readonlyMonthRow('No. of prospects (×15)', c.prospects15, 'prospects15', fmtInt)}
            ${readonlyMonthRow('No. of appointments (×5)', c.appts5, 'appts5', fmtInt)}
            ${readonlyMonthRow('No. of client meetings (×3)', c.meetings3, 'meetings3', fmtInt)}
          </tbody></table>
        </div>

        <h4>Recruitment activity</h4>
        <div class="dg-scroll" style="max-height:none">
          <table class="dg-table"><thead><tr><th style="position:sticky;left:0;background:#faf7fb">Metric</th>${monthHead}</tr></thead>
          <tbody id="recruitActRows">
            ${editableMonthRow('Target no. of coded (20% conversion)', ap.recruitmentActivity.targetCoded, 'targetCoded')}
            ${readonlyMonthRow('No. BYB show-ups (50% show-up)', c.showUps, 'showUps', fmtInt)}
            ${readonlyMonthRow('No. of confirmed BYB guests', c.confirmedBYB, 'confirmedBYB', fmtInt)}
            ${readonlyMonthRow('No. of prospects', c.recProspects, 'recProspects', fmtInt)}
          </tbody></table>
        </div>
      </div>

      <div class="card">
        <h3>Sample monthly calendar <span class="text-muted" style="font-weight:400;font-size:12px">(personal weekly habit template — edit freely)</span></h3>
        <div id="calendarWrap"></div>
      </div>
    `;

    renderCalendar();
    wire();
  }

  function renderCalendar() {
    const wrap = container.querySelector('#calendarWrap');
    const days = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];
    wrap.innerHTML = ap.weeklyCalendar.map((wk, wi) => `
      <h4>Week ${wk.week}</h4>
      <table class="cal-table"><thead><tr>${days.map((d) => `<th>${d}</th>`).join('')}</tr></thead>
      <tbody><tr>
        ${days.map((d) => `
          <td data-week="${wi}" data-day="${d}">
            ${(wk.days[d] || []).map((e, ei) => `
              <div class="cal-entry" data-ei="${ei}">
                <input type="time" class="cal-time" value="${e.time || ''}" />
                <input type="text" class="cal-act" value="${escapeHtml(e.activity || '')}" placeholder="Activity" />
                <button class="dg-del-btn cal-del">✕</button>
              </div>`).join('')}
            <button class="btn btn-light btn-sm cal-add">+ Add</button>
          </td>`).join('')}
      </tr></tbody></table>
    `).join('');

    wrap.querySelectorAll('td[data-week]').forEach((td) => {
      const wi = +td.dataset.week, day = td.dataset.day;
      const entries = ap.weeklyCalendar[wi].days[day] || (ap.weeklyCalendar[wi].days[day] = []);
      td.querySelectorAll('.cal-entry').forEach((row) => {
        const ei = +row.dataset.ei;
        row.querySelector('.cal-time').addEventListener('change', (e) => { entries[ei].time = e.target.value; markDirty('action-plan'); });
        row.querySelector('.cal-act').addEventListener('input', (e) => { entries[ei].activity = e.target.value; markDirty('action-plan'); });
        row.querySelector('.cal-del').addEventListener('click', () => { entries.splice(ei, 1); markDirty('action-plan'); renderCalendar(); });
      });
      td.querySelector('.cal-add').addEventListener('click', () => { entries.push({ time: '', activity: '' }); markDirty('action-plan'); renderCalendar(); });
    });
  }

  // Structural (none currently needed here, kept for parity/future use + import/restore).
  function persist() { markDirty('action-plan'); paint(); }

  // Recompute and patch only the readonly computed cells — no input elements
  // are touched, so typing keeps focus/cursor/scroll intact.
  function updateComputed() {
    const ac = computeAgape(agape);
    const c = computeActionPlan(ap, ac.annualPremiumTarget);
    const patchRow = (rowId, values, fmt = fmtMoney) => {
      const tr = container.querySelector(`tr[data-row="${rowId}"]`);
      if (!tr) return;
      tr.querySelectorAll('td[data-i]').forEach((td) => { td.textContent = fmt(values[+td.dataset.i]); });
    };
    patchRow('targetApe', c.targetApe);
    patchRow('actualApe', c.actualApe);
    patchRow('excess', c.excess);
    patchRow('income', c.income);
    patchRow('beginMP', c.beginMP, fmtInt);
    patchRow('endMP', c.endMP, fmtInt);
    patchRow('prospects15', c.prospects15, fmtInt);
    patchRow('appts5', c.appts5, fmtInt);
    patchRow('meetings3', c.meetings3, fmtInt);
    patchRow('showUps', c.showUps, fmtInt);
    patchRow('confirmedBYB', c.confirmedBYB, fmtInt);
    patchRow('recProspects', c.recProspects, fmtInt);
  }

  function wire() {
    container.querySelectorAll('#salesRows input, #recruitRows input, #salesActRows input, #recruitActRows input').forEach((input) => {
      input.addEventListener('input', () => {
        const i = +input.dataset.i, key = input.dataset.key, v = num(input.value);
        if (key in ap.sales) ap.sales[key][i] = v;
        else if (key in ap.recruitment) ap.recruitment[key][i] = v;
        else if (key in ap.salesActivity) ap.salesActivity[key][i] = v;
        else if (key in ap.recruitmentActivity) ap.recruitmentActivity[key][i] = v;
        markDirty('action-plan');
        updateComputed();
      });
    });
    container.querySelector('#beginMP').addEventListener('input', (e) => {
      ap.recruitment.beginningManpowerJan = num(e.target.value);
      markDirty('action-plan');
      updateComputed();
    });
    container.querySelector('#exportBtn').addEventListener('click', doExport);
    container.querySelector('#importInput').addEventListener('change', doImport);

    container.querySelectorAll('.money-input').forEach((el) => wireMoneyInput(el));
  }

  function flatten() {
    const rows = [];
    ap.months.forEach((m, i) => {
      rows.push({
        month: m,
        averageCaseSize: num(ap.sales.averageCaseSize[i]),
        numberOfCases: num(ap.sales.numberOfCases[i]),
        targetNewRecruits: num(ap.recruitment.targetNewRecruits[i]),
        provisionForAttrition: num(ap.recruitment.provisionForAttrition[i]),
        targetCasesForMonth: num(ap.salesActivity.targetCasesForMonth[i]),
        targetCoded: num(ap.recruitmentActivity.targetCoded[i]),
      });
    });
    return rows;
  }

  function doExport() {
    const columns = [
      { key: 'month', header: 'Month' },
      { key: 'averageCaseSize', header: 'Average Case Size' },
      { key: 'numberOfCases', header: 'Number of Cases' },
      { key: 'targetNewRecruits', header: 'Target New Recruits' },
      { key: 'provisionForAttrition', header: 'Provision for Attrition' },
      { key: 'targetCasesForMonth', header: 'Target Cases for Month' },
      { key: 'targetCoded', header: 'Target No. Coded' },
    ];
    downloadSingleSheet('ACTION PLAN', columns, flatten(), 'Action Plan.xlsx');
  }

  async function doImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const columns = [
        { key: 'month', header: 'Month', aliases: ['Month'] },
        { key: 'averageCaseSize', header: 'Average Case Size' },
        { key: 'numberOfCases', header: 'Number of Cases' },
        { key: 'targetNewRecruits', header: 'Target New Recruits' },
        { key: 'provisionForAttrition', header: 'Provision for Attrition' },
        { key: 'targetCasesForMonth', header: 'Target Cases for Month' },
        { key: 'targetCoded', header: 'Target No. Coded' },
      ];
      const rows = parseSheetRows(wb, columns, 'ACTION PLAN');
      let matched = 0;
      rows.forEach((r) => {
        const i = ap.months.findIndex((m) => m.toLowerCase() === String(r.month || '').toLowerCase());
        if (i === -1) return;
        matched++;
        if (r.averageCaseSize !== null) ap.sales.averageCaseSize[i] = num(r.averageCaseSize);
        if (r.numberOfCases !== null) ap.sales.numberOfCases[i] = num(r.numberOfCases);
        if (r.targetNewRecruits !== null) ap.recruitment.targetNewRecruits[i] = num(r.targetNewRecruits);
        if (r.provisionForAttrition !== null) ap.recruitment.provisionForAttrition[i] = num(r.provisionForAttrition);
        if (r.targetCasesForMonth !== null) ap.salesActivity.targetCasesForMonth[i] = num(r.targetCasesForMonth);
        if (r.targetCoded !== null) ap.recruitmentActivity.targetCoded[i] = num(r.targetCoded);
      });
      persist();
      showToast(`Imported ${matched} month${matched === 1 ? '' : 's'} from Excel.`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  }

  paint();
}
