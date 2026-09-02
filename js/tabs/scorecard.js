import { loadTab, markDirty } from '../store.js';
import { num, fmtMoney, uid, escapeHtml, showToast, withScrollPreserved, wireMoneyInput } from '../utils.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';

function computeTripApe(trip, monthlyApe) {
  if (Array.isArray(trip.apeFypRange)) {
    const [s, e] = trip.apeFypRange;
    return monthlyApe.slice(s, e + 1).reduce((a, b) => a + num(b), 0);
  }
  if (trip.apeFypRange === 'total') return monthlyApe.reduce((a, b) => a + num(b), 0);
  return null;
}

export async function render(container) {
  const d = await loadTab('scorecard');
  d.achieversClub.forEach((r) => { if (!r._id) r._id = uid(); });
  d.tripQualifications.forEach((r) => { if (!r._id) r._id = uid(); });

  function paint() { withScrollPreserved(container, paintInner); }

  function paintInner() {
    const totalApe = d.monthlyApe.reduce((a, b) => a + num(b), 0);
    const totalCc = d.monthlyCc.reduce((a, b) => a + num(b), 0);

    container.innerHTML = `
      <div class="tab-header">
        <div>
          <h1 class="tab-title">🏆 Personal Scorecard</h1>
          <p class="tab-subtitle">Achiever's Club, trip qualifications and QPB — targets and reward text stay as your leaders set them; "TO GO" and reward figures recompute live from your monthly APE & case count.</p>
        </div>
        <div class="tab-actions">
          <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
          <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
        </div>
      </div>

      <div class="card">
        <h3>Monthly APE & case count <span class="text-muted" style="font-weight:400">— ${escapeHtml(d.asOf)}</span></h3>
        <div class="dg-scroll" style="max-height:none">
          <table class="dg-table">
            <thead><tr><th style="position:sticky;left:0;background:#faf7fb">Month</th>${d.months.map((m) => `<th class="num">${m.slice(0,3)}</th>`).join('')}<th class="num">Total</th></tr></thead>
            <tbody>
              <tr><td style="position:sticky;left:0;background:#fff">APE</td>${d.monthlyApe.map((v, i) => `<td class="num"><input class="cell-input money-input" type="text" inputmode="decimal" data-arr="monthlyApe" data-i="${i}" value="${v ?? 0}" style="width:78px;text-align:right" /></td>`).join('')}<td class="num" id="apeTotalCell" style="font-weight:700">₱ ${fmtMoney(totalApe)}</td></tr>
              <tr><td style="position:sticky;left:0;background:#fff">Case count</td>${d.monthlyCc.map((v, i) => `<td class="num"><input class="cell-input" type="number" step="any" data-arr="monthlyCc" data-i="${i}" value="${v ?? 0}" style="width:78px;text-align:right" /></td>`).join('')}<td class="num" id="ccTotalCell" style="font-weight:700">${totalCc}</td></tr>
            </tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <h3>Achiever's Club 2026</h3>
        <table class="dg-table" style="width:100%">
          <thead><tr><th>Goal</th><th class="num" style="width:130px">Target</th><th>Reward</th><th style="width:110px">Period</th><th class="num" style="width:120px">To go</th><th style="width:34px"></th></tr></thead>
          <tbody id="achieversBody">
            ${d.achieversClub.map((r) => `
              <tr data-id="${r._id}">
                <td><input class="cell-input" type="text" data-f="goal" value="${escapeHtml(r.goal)}" /></td>
                <td class="num"><input class="cell-input money-input" type="text" inputmode="decimal" data-f="target" value="${r.target ?? 0}" style="text-align:right" /></td>
                <td><input class="cell-input" type="text" data-f="reward" value="${escapeHtml(r.reward || '')}" /></td>
                <td><input class="cell-input" type="text" data-f="period" value="${escapeHtml(r.period || '')}" /></td>
                <td class="num ach-togo">₱ ${fmtMoney(num(r.target) - totalApe)}</td>
                <td><button class="dg-del-btn ach-del">✕</button></td>
              </tr>`).join('')}
          </tbody>
        </table>
        <button class="btn btn-light btn-sm" id="addAchiever" style="margin-top:8px;">+ Add goal</button>

        <h4>Rookie Producers Bonus</h4>
        <p class="text-muted" style="font-size:12.5px">${d.rookieProducersBonus.tiers.map(escapeHtml).join('<br>')}</p>
        <label class="text-muted">Reward per case count</label>
        <input type="text" inputmode="decimal" class="money-input" id="rewardPerCase" value="${d.rookieProducersBonus.rewardPerCase ?? 0}" style="max-width:160px" />
      </div>

      <div class="card">
        <h3>Trip Qualifications 2026</h3>
        <table class="dg-table" style="width:100%">
          <thead><tr><th>Trip</th><th style="width:160px">Target</th><th>Reward</th><th class="num" style="width:120px">APE / FYP</th><th style="width:34px"></th></tr></thead>
          <tbody id="tripsBody">
            ${d.tripQualifications.map((r) => {
              const ape = computeTripApe(r, d.monthlyApe);
              return `
              <tr data-id="${r._id}">
                <td><input class="cell-input" type="text" data-f="trip" value="${escapeHtml(r.trip)}" /></td>
                <td><input class="cell-input" type="text" data-f="target" value="${escapeHtml(r.target ?? '')}" /></td>
                <td><input class="cell-input" type="text" data-f="reward" value="${escapeHtml(r.reward || '')}" /></td>
                <td class="num trip-ape">${ape === null ? '<span class="text-muted">—</span>' : '₱ ' + fmtMoney(ape)}</td>
                <td><button class="dg-del-btn trip-del">✕</button></td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
        <button class="btn btn-light btn-sm" id="addTrip" style="margin-top:8px;">+ Add trip</button>
      </div>

      <div class="card">
        <h3>QPB 2026</h3>
        <p class="section-note">Reward figures here are planning estimates recomputed from "FYP for quarter" × tier % — treat them as a guide, not a guaranteed payout.</p>
        ${d.qpb.quarters.map((q, qi) => {
          const qApe = q.monthIdx.reduce((s, mi) => s + num(d.monthlyApe[mi]), 0);
          return `
          <h4>${escapeHtml(q.label)} <span class="text-muted" style="font-weight:400">— APE so far: <span id="qpbApe-${qi}">₱ ${fmtMoney(qApe)}</span></span></h4>
          <div style="max-width:220px;margin-bottom:8px;">
            <label class="text-muted">FYP for quarter</label>
            <input type="text" inputmode="decimal" class="qpb-fyp money-input" data-qi="${qi}" value="${q.fypForQuarter ?? 0}" />
          </div>
          <table class="dg-table" style="width:100%;margin-bottom:14px;">
            <thead><tr><th class="num" style="width:140px">Target</th><th class="num" style="width:80px">%</th><th class="num" style="width:130px">To go</th><th class="num" style="width:130px">Reward</th></tr></thead>
            <tbody>
              ${q.tiers.map((t, ti) => `
                <tr>
                  <td class="num"><input class="cell-input qpb-target money-input" data-qi="${qi}" data-ti="${ti}" type="text" inputmode="decimal" value="${t.target}" style="text-align:right" /></td>
                  <td class="num"><input class="cell-input qpb-pct" data-qi="${qi}" data-ti="${ti}" type="number" step="0.01" value="${t.percent}" style="text-align:right" /></td>
                  <td class="num qpb-togo" data-qi="${qi}" data-ti="${ti}">₱ ${fmtMoney(t.target - qApe)}</td>
                  <td class="num qpb-reward" data-qi="${qi}" data-ti="${ti}">₱ ${fmtMoney((num(q.fypForQuarter)) * num(t.percent))}</td>
                </tr>`).join('')}
            </tbody>
          </table>`;
        }).join('')}
      </div>
    `;
    wire();
  }

  // Structural changes (add/remove a goal or trip) still need a full repaint.
  function persist() { markDirty('scorecard'); paint(); }

  // Recompute + patch only the computed output cells — no input elements are
  // touched, so typing keeps focus/cursor/scroll intact.
  function updateComputed() {
    const totalApe = d.monthlyApe.reduce((a, b) => a + num(b), 0);
    const totalCc = d.monthlyCc.reduce((a, b) => a + num(b), 0);

    const apeCell = container.querySelector('#apeTotalCell');
    if (apeCell) apeCell.textContent = '₱ ' + fmtMoney(totalApe);
    const ccCell = container.querySelector('#ccTotalCell');
    if (ccCell) ccCell.textContent = totalCc;

    container.querySelectorAll('#achieversBody tr').forEach((tr) => {
      const r = d.achieversClub.find((x) => x._id === tr.dataset.id);
      const cell = tr.querySelector('.ach-togo');
      if (r && cell) cell.textContent = '₱ ' + fmtMoney(num(r.target) - totalApe);
    });

    container.querySelectorAll('#tripsBody tr').forEach((tr) => {
      const r = d.tripQualifications.find((x) => x._id === tr.dataset.id);
      const cell = tr.querySelector('.trip-ape');
      if (!r || !cell) return;
      const ape = computeTripApe(r, d.monthlyApe);
      cell.innerHTML = ape === null ? '<span class="text-muted">—</span>' : '₱ ' + fmtMoney(ape);
    });

    d.qpb.quarters.forEach((q, qi) => {
      const qApe = q.monthIdx.reduce((s, mi) => s + num(d.monthlyApe[mi]), 0);
      const apeSpan = container.querySelector(`#qpbApe-${qi}`);
      if (apeSpan) apeSpan.textContent = '₱ ' + fmtMoney(qApe);
      q.tiers.forEach((t, ti) => {
        const togoCell = container.querySelector(`.qpb-togo[data-qi="${qi}"][data-ti="${ti}"]`);
        if (togoCell) togoCell.textContent = '₱ ' + fmtMoney(t.target - qApe);
        const rewardCell = container.querySelector(`.qpb-reward[data-qi="${qi}"][data-ti="${ti}"]`);
        if (rewardCell) rewardCell.textContent = '₱ ' + fmtMoney(num(q.fypForQuarter) * num(t.percent));
      });
    });
  }

  function wire() {
    container.querySelectorAll('[data-arr]').forEach((input) => {
      input.addEventListener('input', () => { d[input.dataset.arr][+input.dataset.i] = num(input.value); markDirty('scorecard'); updateComputed(); });
    });
    container.querySelectorAll('#achieversBody tr').forEach((tr) => {
      const r = d.achieversClub.find((x) => x._id === tr.dataset.id);
      tr.querySelectorAll('[data-f]').forEach((inp) => {
        inp.addEventListener('input', () => { r[inp.dataset.f] = inp.dataset.f === 'target' ? num(inp.value) : inp.value; markDirty('scorecard'); updateComputed(); });
      });
      tr.querySelector('.ach-del').addEventListener('click', () => { d.achieversClub.splice(d.achieversClub.indexOf(r), 1); persist(); });
    });
    const addAch = container.querySelector('#addAchiever');
    if (addAch) addAch.addEventListener('click', () => { d.achieversClub.push({ _id: uid(), goal: '', target: 0, reward: '', period: '' }); persist(); });

    container.querySelector('#rewardPerCase').addEventListener('input', (e) => { d.rookieProducersBonus.rewardPerCase = num(e.target.value); markDirty('scorecard'); });

    container.querySelectorAll('#tripsBody tr').forEach((tr) => {
      const r = d.tripQualifications.find((x) => x._id === tr.dataset.id);
      tr.querySelectorAll('[data-f]').forEach((inp) => {
        inp.addEventListener('input', () => { r[inp.dataset.f] = inp.value; markDirty('scorecard'); updateComputed(); });
      });
      tr.querySelector('.trip-del').addEventListener('click', () => { d.tripQualifications.splice(d.tripQualifications.indexOf(r), 1); persist(); });
    });
    const addTrip = container.querySelector('#addTrip');
    if (addTrip) addTrip.addEventListener('click', () => { d.tripQualifications.push({ _id: uid(), trip: '', target: '', reward: '', apeFypRange: null }); persist(); });

    container.querySelectorAll('.qpb-fyp').forEach((inp) => {
      inp.addEventListener('input', () => {
        const q = d.qpb.quarters[+inp.dataset.qi];
        q.fypForQuarter = num(inp.value);
        markDirty('scorecard'); updateComputed();
      });
    });
    container.querySelectorAll('.qpb-target, .qpb-pct').forEach((inp) => {
      inp.addEventListener('input', () => {
        const q = d.qpb.quarters[+inp.dataset.qi];
        const t = q.tiers[+inp.dataset.ti];
        if (inp.classList.contains('qpb-target')) t.target = num(inp.value);
        else t.percent = num(inp.value);
        markDirty('scorecard'); updateComputed();
      });
    });

    container.querySelector('#exportBtn').addEventListener('click', doExport);
    container.querySelector('#importInput').addEventListener('change', doImport);

    container.querySelectorAll('.money-input').forEach((el) => wireMoneyInput(el));
  }

  function doExport() {
    const columns = [{ key: 'month', header: 'Month' }, { key: 'ape', header: 'APE' }, { key: 'cc', header: 'Case Count' }];
    const rows = d.months.map((m, i) => ({ month: m, ape: num(d.monthlyApe[i]), cc: num(d.monthlyCc[i]) }));
    downloadSingleSheet('SCORECARD', columns, rows, 'Scorecard.xlsx');
  }

  async function doImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const columns = [{ key: 'month', header: 'Month', aliases: ['Month'] }, { key: 'ape', header: 'APE' }, { key: 'cc', header: 'Case Count' }];
      const rows = parseSheetRows(wb, columns, 'SCORECARD');
      let matched = 0;
      rows.forEach((r) => {
        const i = d.months.findIndex((m) => m.toLowerCase() === String(r.month || '').toLowerCase());
        if (i === -1) return;
        matched++;
        if (r.ape !== null) d.monthlyApe[i] = num(r.ape);
        if (r.cc !== null) d.monthlyCc[i] = num(r.cc);
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
