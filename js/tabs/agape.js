import { loadTab, markDirty } from '../store.js';
import { num, fmtMoney, uid, escapeHtml } from '../utils.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';
import { showToast } from '../utils.js';
import { computeAgape as compute } from '../formulas.js';

function itemRow(label, amount, onLabel, onAmount, onDel, extra) {
  return `
    <tr>
      <td><input class="cell-input item-label" type="text" value="${escapeHtml(label)}" /></td>
      ${extra !== undefined ? `<td style="width:120px"><input class="cell-input item-extra" type="text" value="${escapeHtml(extra ?? '')}" /></td>` : ''}
      <td class="num" style="width:150px"><input class="cell-input item-amount" type="number" step="any" value="${amount ?? ''}" style="text-align:right" /></td>
      <td style="width:34px"><button class="dg-del-btn item-del" title="Remove">✕</button></td>
    </tr>`;
}

export async function render(container) {
  const d = await loadTab('agape');
  d.sections.forEach((s) => s.groups.forEach((g) => g.items.forEach((it) => { if (!it._id) it._id = uid(); })));
  d.goals.forEach((it) => { if (!it._id) it._id = uid(); });
  d.ftJobItems.forEach((it) => { if (!it._id) it._id = uid(); });

  function paint() {
    const c = compute(d);

    const groupHtml = (sec) => sec.groups.map((g) => `
      <h4>${escapeHtml(g.title)}</h4>
      <table class="kv-table" data-sec="${sec.id}" data-grp="${escapeHtml(g.title)}">
        <tbody>
          ${g.items.map((it) => itemRow(it.label, it.amount).replace('<tr>', `<tr data-id="${it._id}">`)).join('')}
        </tbody>
      </table>
      <button class="btn btn-light btn-sm add-item" data-sec="${sec.id}" data-grp="${escapeHtml(g.title)}" style="margin:6px 0 4px;">+ Add expense line</button>
    `).join('');

    container.innerHTML = `
      <div class="tab-header">
        <div>
          <h1 class="tab-title">🌱 AGAPE</h1>
          <p class="tab-subtitle">Annual Goals and Action Plans toward Excellence — fill in your monthly budget and it flows straight through to your annual premium target.</p>
        </div>
        <div class="tab-actions">
          <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
          <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
        </div>
      </div>

      <div class="grid-2">
        <div class="card">
          <h3>I. Planning my monthly budget</h3>
          ${d.sections.map((sec) => sec.id === 'living' ? `<h4 style="margin-top:0">${escapeHtml(sec.title)}</h4>${groupHtml(sec)}` : '').join('')}
          <table class="kv-table"><tbody><tr class="total-row"><td>Total monthly for living expenses</td><td class="num" style="width:150px">₱ ${fmtMoney(c.livingTotal)}</td><td style="width:34px"></td></tr></tbody></table>
        </div>

        <div class="card">
          ${d.sections.map((sec) => sec.id === 'business' ? `<h3>${escapeHtml(sec.title)}</h3>${groupHtml(sec)}<table class="kv-table"><tbody><tr class="total-row"><td>Total monthly for business expenses</td><td class="num" style="width:150px">₱ ${fmtMoney(c.businessTotal)}</td><td style="width:34px"></td></tr></tbody></table>` : '').join('')}
          ${d.sections.map((sec) => sec.id === 'savings' ? `<h3 style="margin-top:16px">${escapeHtml(sec.title)}</h3>${groupHtml(sec)}<table class="kv-table"><tbody><tr class="total-row"><td>Total monthly for savings & investments</td><td class="num" style="width:150px">₱ ${fmtMoney(c.savingsTotal)}</td><td style="width:34px"></td></tr></tbody></table>` : '').join('')}
        </div>
      </div>

      <div class="grid-2">
        <div class="card">
          <h3>II. Other financial goals (long term)</h3>
          <table class="kv-table" data-goals="1">
            <thead><tr><th>Goal</th><th style="width:120px">Time</th><th style="width:150px">Amount / month</th><th></th></tr></thead>
            <tbody>
              ${d.goals.map((it) => itemRow(it.label, it.amount, null, null, null, it.time).replace('<tr>', `<tr data-id="${it._id}">`)).join('')}
            </tbody>
          </table>
          <button class="btn btn-light btn-sm add-goal" style="margin:6px 0 4px;">+ Add goal</button>
          <table class="kv-table"><tbody><tr class="total-row"><td>Total monthly for future & goals</td><td class="num" style="width:150px">₱ ${fmtMoney(c.goalsTotal)}</td><td style="width:34px"></td></tr></tbody></table>
        </div>

        <div class="card">
          <h3>Expenses related to my full-time job</h3>
          <table class="kv-table" data-ftjob="1">
            <tbody>
              ${d.ftJobItems.map((it) => itemRow(it.label, it.amount).replace('<tr>', `<tr data-id="${it._id}">`)).join('')}
            </tbody>
          </table>
          <button class="btn btn-light btn-sm add-ftjob" style="margin:6px 0 4px;">+ Add line</button>
          <table class="kv-table"><tbody><tr class="total-row"><td>Total</td><td class="num" style="width:150px">₱ ${fmtMoney(c.ftJobTotal)}</td><td style="width:34px"></td></tr></tbody></table>
          <div style="margin-top:12px"><label class="text-muted">Income from employment (monthly)</label>
            <input type="number" step="any" id="incomeFromEmployment" value="${d.incomeFromEmployment ?? 0}" />
          </div>
        </div>
      </div>

      <div class="card">
        <h3>III. Planning my annual income goal</h3>
        <div class="grid-3">
          <div class="stat"><div class="stat-label">Monthly income requirement</div><div class="stat-value">₱ ${fmtMoney(c.incomeRequirement)}</div></div>
          <div class="stat"><div class="stat-label">Target monthly income from PRU LIFE UK</div><div class="stat-value">₱ ${fmtMoney(c.targetFromBusiness)}</div></div>
          <div class="stat"><div class="stat-label">Annual income requirement</div><div class="stat-value">₱ ${fmtMoney(c.annualIncomeRequirement)}</div></div>
        </div>
        <div class="grid-3" style="margin-top:14px">
          <div>
            <label class="text-muted">Average first-year commission (e.g. 0.30 = 30%)</label>
            <input type="number" step="0.01" id="avgFYC" value="${d.avgFirstYearCommission ?? 0.3}" />
          </div>
          <div>
            <label class="text-muted">Average APE per case</label>
            <input type="number" step="any" id="avgApe" value="${d.avgApePerCase ?? 36000}" />
          </div>
          <div class="stat"><div class="stat-label">Annual premium target</div><div class="stat-value">₱ ${fmtMoney(c.annualPremiumTarget)}</div></div>
        </div>
        <div class="stat" style="margin-top:14px; max-width:260px;"><div class="stat-label">Number of cases in one year</div><div class="stat-value">${c.numCases.toFixed(1)}</div></div>
      </div>

      <div class="card">
        <h3>IV. Planning my sales activity</h3>
        <p class="section-note">Success formula: ${escapeHtml('15 Prospects → 5 Appointments → 3 Closed Sales')}</p>
        <table class="kv-table">
          <thead><tr><th>Sales activity</th><th class="num">Annual</th><th class="num">Monthly (÷12)</th></tr></thead>
          <tbody>
            <tr><td>No. of prospects (× 15)</td><td class="num">${c.prospects.toFixed(1)}</td><td class="num">${(c.prospects/12).toFixed(1)}</td></tr>
            <tr><td>No. of appointments (× 5)</td><td class="num">${c.appointments.toFixed(1)}</td><td class="num">${(c.appointments/12).toFixed(1)}</td></tr>
            <tr><td>No. of client meetings (× 3)</td><td class="num">${c.meetings.toFixed(1)}</td><td class="num">${(c.meetings/12).toFixed(1)}</td></tr>
            <tr><td>No. of closed sales (× 1)</td><td class="num">${c.closed.toFixed(1)}</td><td class="num">${(c.closed/12).toFixed(1)}</td></tr>
          </tbody>
        </table>
      </div>
    `;

    wire();
  }

  function persist() { markDirty('agape'); paint(); }

  function wire() {
    container.querySelectorAll('table[data-sec] tr[data-id]').forEach((tr) => {
      const sec = d.sections.find((s) => s.id === tr.closest('table').dataset.sec);
      const grp = sec.groups.find((g) => g.title === tr.closest('table').dataset.grp);
      const it = grp.items.find((x) => x._id === tr.dataset.id);
      tr.querySelector('.item-label').addEventListener('input', (e) => { it.label = e.target.value; markDirty('agape'); });
      tr.querySelector('.item-amount').addEventListener('input', (e) => { it.amount = num(e.target.value); persist(); });
      tr.querySelector('.item-del').addEventListener('click', () => { grp.items.splice(grp.items.indexOf(it), 1); persist(); });
    });

    container.querySelectorAll('.add-item').forEach((btn) => {
      btn.addEventListener('click', () => {
        const sec = d.sections.find((s) => s.id === btn.dataset.sec);
        const grp = sec.groups.find((g) => g.title === btn.dataset.grp);
        grp.items.push({ _id: uid(), label: '', amount: 0 });
        persist();
      });
    });

    container.querySelectorAll('table[data-goals] tr[data-id]').forEach((tr) => {
      const it = d.goals.find((x) => x._id === tr.dataset.id);
      tr.querySelector('.item-label').addEventListener('input', (e) => { it.label = e.target.value; markDirty('agape'); });
      tr.querySelector('.item-extra').addEventListener('input', (e) => { it.time = e.target.value; markDirty('agape'); });
      tr.querySelector('.item-amount').addEventListener('input', (e) => { it.amount = num(e.target.value); persist(); });
      tr.querySelector('.item-del').addEventListener('click', () => { d.goals.splice(d.goals.indexOf(it), 1); persist(); });
    });
    const addGoal = container.querySelector('.add-goal');
    if (addGoal) addGoal.addEventListener('click', () => { d.goals.push({ _id: uid(), label: '', amount: 0, time: '' }); persist(); });

    container.querySelectorAll('table[data-ftjob] tr[data-id]').forEach((tr) => {
      const it = d.ftJobItems.find((x) => x._id === tr.dataset.id);
      tr.querySelector('.item-label').addEventListener('input', (e) => { it.label = e.target.value; markDirty('agape'); });
      tr.querySelector('.item-amount').addEventListener('input', (e) => { it.amount = num(e.target.value); persist(); });
      tr.querySelector('.item-del').addEventListener('click', () => { d.ftJobItems.splice(d.ftJobItems.indexOf(it), 1); persist(); });
    });
    const addFt = container.querySelector('.add-ftjob');
    if (addFt) addFt.addEventListener('click', () => { d.ftJobItems.push({ _id: uid(), label: '', amount: 0 }); persist(); });

    container.querySelector('#incomeFromEmployment').addEventListener('input', (e) => { d.incomeFromEmployment = num(e.target.value); persist(); });
    container.querySelector('#avgFYC').addEventListener('input', (e) => { d.avgFirstYearCommission = num(e.target.value); persist(); });
    container.querySelector('#avgApe').addEventListener('input', (e) => { d.avgApePerCase = num(e.target.value); persist(); });

    container.querySelector('#exportBtn').addEventListener('click', doExport);
    container.querySelector('#importInput').addEventListener('change', doImport);
  }

  function flatten() {
    const rows = [];
    d.sections.forEach((sec) => sec.groups.forEach((g) => g.items.forEach((it) => {
      rows.push({ section: sec.title, group: g.title, label: it.label, time: '', amount: num(it.amount) });
    })));
    d.goals.forEach((it) => rows.push({ section: 'II. PLANNING MY OTHER FINANCIAL GOALS (LONG TERM)', group: 'GOAL', label: it.label, time: it.time || '', amount: num(it.amount) }));
    d.ftJobItems.forEach((it) => rows.push({ section: 'EXPENSES RELATED TO MY FULL TIME JOB', group: '', label: it.label, time: '', amount: num(it.amount) }));
    rows.push({ section: 'SETTINGS', group: '', label: 'Income from employment (monthly)', time: '', amount: num(d.incomeFromEmployment) });
    rows.push({ section: 'SETTINGS', group: '', label: 'Average first-year commission (0-1)', time: '', amount: num(d.avgFirstYearCommission, 0.3) });
    rows.push({ section: 'SETTINGS', group: '', label: 'Average APE per case', time: '', amount: num(d.avgApePerCase, 36000) });
    return rows;
  }

  function doExport() {
    const columns = [
      { key: 'section', header: 'Section' }, { key: 'group', header: 'Group' },
      { key: 'label', header: 'Label' }, { key: 'time', header: 'Time' }, { key: 'amount', header: 'Amount' },
    ];
    downloadSingleSheet('AGAPE', columns, flatten(), 'AGAPE.xlsx');
  }

  async function doImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const columns = [
        { key: 'section', header: 'Section' }, { key: 'group', header: 'Group' },
        { key: 'label', header: 'Label' }, { key: 'time', header: 'Time' }, { key: 'amount', header: 'Amount' },
      ];
      const rows = parseSheetRows(wb, columns, 'AGAPE');
      let matched = 0;
      rows.forEach((r) => {
        if (r.section === 'SETTINGS') {
          if (r.label && r.label.startsWith('Income from employment')) d.incomeFromEmployment = num(r.amount);
          if (r.label && r.label.startsWith('Average first-year')) d.avgFirstYearCommission = num(r.amount);
          if (r.label && r.label.startsWith('Average APE')) d.avgApePerCase = num(r.amount);
          return;
        }
        if (r.section && r.section.includes('FINANCIAL GOALS')) {
          const g = d.goals.find((x) => x.label === r.label);
          if (g) { g.amount = num(r.amount); g.time = r.time || g.time; matched++; }
          else if (r.label) { d.goals.push({ _id: uid(), label: r.label, amount: num(r.amount), time: r.time || '' }); matched++; }
          return;
        }
        if (r.section && r.section.includes('FULL TIME JOB')) {
          const it = d.ftJobItems.find((x) => x.label === r.label);
          if (it) { it.amount = num(r.amount); matched++; }
          else if (r.label) { d.ftJobItems.push({ _id: uid(), label: r.label, amount: num(r.amount) }); matched++; }
          return;
        }
        const sec = d.sections.find((s) => s.title === r.section);
        if (!sec) return;
        const grp = sec.groups.find((g) => g.title === r.group);
        if (!grp) return;
        const it = grp.items.find((x) => x.label === r.label);
        if (it) { it.amount = num(r.amount); matched++; }
      });
      persist();
      showToast(`Imported ${matched} value${matched === 1 ? '' : 's'} from Excel.`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  }

  paint();
}
