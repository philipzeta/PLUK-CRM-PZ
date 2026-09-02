import { loadTab, markDirty } from '../store.js';
import { uid, escapeHtml, showToast, withScrollPreserved } from '../utils.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';

export async function render(container) {
  const d = await loadTab('product-knowledge');
  d.categories.forEach((cat) => { if (!cat._id) cat._id = uid(); cat.products.forEach((p) => { if (!p._id) p._id = uid(); }); });

  function paint() { withScrollPreserved(container, paintInner); }

  function paintInner() {
    const allRatings = d.categories.flatMap((c) => c.products.map((p) => p.rating).filter((r) => r));
    const avg = allRatings.length ? (allRatings.reduce((a, b) => a + b, 0) / allRatings.length).toFixed(1) : '—';

    container.innerHTML = `
      <div class="tab-header">
        <div>
          <h1 class="tab-title">📦 Product Knowledge</h1>
          <p class="tab-subtitle">${escapeHtml(d.scaleLabel)} Rate yourself 1 (needs work) to 5 (expert).</p>
        </div>
        <div class="tab-actions">
          <button class="btn btn-light" id="exportBtn">⬇ Export to Excel</button>
          <label class="btn btn-light">⬆ Import from Excel<input type="file" id="importInput" accept=".xlsx,.xls" hidden /></label>
        </div>
      </div>
      <div class="stat" style="max-width:220px;margin-bottom:16px;"><div class="stat-label">Average self-rating</div><div class="stat-value">${avg} / 5</div></div>

      ${d.categories.map((cat) => `
        <div class="card" data-cat="${cat._id}">
          <h3><input class="cell-input cat-title" type="text" value="${escapeHtml(cat.title)}" style="font-weight:700;color:var(--brand-dark);font-size:14.5px" /></h3>
          <table class="kv-table">
            <tbody>
              ${cat.products.map((p) => `
                <tr data-pid="${p._id}">
                  <td><input class="cell-input prod-name" type="text" value="${escapeHtml(p.name)}" /></td>
                  <td style="width:190px">
                    <div class="rating-scale">
                      ${[1,2,3,4,5].map((n) => `<div class="rating-dot ${p.rating === n ? 'active' : ''}" data-n="${n}">${n}</div>`).join('')}
                    </div>
                  </td>
                  <td style="width:34px"><button class="dg-del-btn prod-del">✕</button></td>
                </tr>`).join('')}
            </tbody>
          </table>
          <button class="btn btn-light btn-sm add-prod" style="margin-top:8px;">+ Add product</button>
        </div>
      `).join('')}
      <button class="btn btn-light" id="addCat">+ Add category</button>
    `;
    wire();
  }

  function persist() { markDirty('product-knowledge'); }

  function wire() {
    container.querySelectorAll('.card[data-cat]').forEach((card) => {
      const cat = d.categories.find((c) => c._id === card.dataset.cat);
      card.querySelector('.cat-title').addEventListener('input', (e) => { cat.title = e.target.value; persist(); });
      card.querySelectorAll('tr[data-pid]').forEach((tr) => {
        const p = cat.products.find((x) => x._id === tr.dataset.pid);
        tr.querySelector('.prod-name').addEventListener('input', (e) => { p.name = e.target.value; persist(); });
        tr.querySelectorAll('.rating-dot').forEach((dot) => {
          dot.addEventListener('click', () => {
            const n = +dot.dataset.n;
            p.rating = p.rating === n ? null : n;
            persist(); paint();
          });
        });
        tr.querySelector('.prod-del').addEventListener('click', () => { cat.products.splice(cat.products.indexOf(p), 1); persist(); paint(); });
      });
      card.querySelector('.add-prod').addEventListener('click', () => { cat.products.push({ _id: uid(), name: '', rating: null }); persist(); paint(); });
    });
    container.querySelector('#addCat').addEventListener('click', () => { d.categories.push({ _id: uid(), title: 'NEW CATEGORY', products: [] }); persist(); paint(); });
    container.querySelector('#exportBtn').addEventListener('click', doExport);
    container.querySelector('#importInput').addEventListener('change', doImport);
  }

  function flatten() {
    const rows = [];
    d.categories.forEach((c) => c.products.forEach((p) => rows.push({ category: c.title, product: p.name, rating: p.rating ?? '' })));
    return rows;
  }
  function doExport() {
    const columns = [{ key: 'category', header: 'Category' }, { key: 'product', header: 'Product' }, { key: 'rating', header: 'Rating (1-5)' }];
    downloadSingleSheet('PRODUCT KNOWLEDGE', columns, flatten(), 'Product Knowledge.xlsx');
  }
  async function doImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = await readWorkbook(file);
      const columns = [{ key: 'category', header: 'Category', aliases: ['Category'] }, { key: 'product', header: 'Product', aliases: ['Product'] }, { key: 'rating', header: 'Rating (1-5)', aliases: ['Rating (1-5)', 'Rating'] }];
      const rows = parseSheetRows(wb, columns, 'PRODUCT KNOWLEDGE');
      let matched = 0;
      rows.forEach((r) => {
        let cat = d.categories.find((c) => c.title === r.category);
        if (!cat) { cat = { _id: uid(), title: r.category || 'OTHER', products: [] }; d.categories.push(cat); }
        let p = cat.products.find((x) => x.name === r.product);
        const rating = r.rating === null || r.rating === '' ? null : parseInt(r.rating, 10);
        if (p) p.rating = rating;
        else cat.products.push({ _id: uid(), name: r.product, rating });
        matched++;
      });
      persist(); paint();
      showToast(`Imported ${matched} row(s) from Excel.`);
    } catch (err) {
      console.error(err);
      showToast('Could not read that file.');
    } finally {
      e.target.value = '';
    }
  }

  paint();
}
