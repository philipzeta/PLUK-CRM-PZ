import { escapeHtml, toISODate, wireMoneyInput } from './utils.js';

// A lightweight, dependency-free editable data grid: search, sort, pagination,
// inline editing (text/number/date/select/checkbox), add row, delete row,
// and optional per-row CSS class (used for overdue-follow-up highlighting).
export function createDataGrid(opts) {
  const {
    container,
    columns,
    getRows,
    onCellChange,
    onAddRow,
    onDeleteRow,
    rowClass,
    pageSize = 60,
    idKey = '_id',
    emptyLabel = 'No rows yet. Add one, or import from Excel.',
  } = opts;

  let search = '';
  // Multi-level sort: an ordered list of { key, dir } — the first entry is the
  // primary sort, later entries break ties left over from earlier ones (like
  // Excel/Sheets "Sort by ... then by ..."). Plain-click a header to sort by
  // just that column; Shift-click a header to add/toggle it as another level.
  let sortLevels = [];
  const MAX_SORT_LEVELS = 4;
  let page = 0;

  function filteredRows() {
    let rows = getRows();
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      rows = rows.filter((r) =>
        columns.some((c) => {
          const v = c.computed ? c.computed(r) : r[c.key];
          return v !== null && v !== undefined && String(v).toLowerCase().includes(q);
        })
      );
    }
    if (sortLevels.length) {
      const levels = sortLevels
        .map((lvl) => ({ ...lvl, col: columns.find((c) => c.key === lvl.key) }))
        .filter((lvl) => lvl.col);
      if (levels.length) {
        rows = rows.slice().sort((a, b) => {
          for (const { col, dir } of levels) {
            const va = col.computed ? col.computed(a) : a[col.key];
            const vb = col.computed ? col.computed(b) : b[col.key];
            if (va === vb) continue;
            if (va === null || va === undefined || va === '') return 1;
            if (vb === null || vb === undefined || vb === '') return -1;
            const cmp = (va > vb ? 1 : -1) * dir;
            if (cmp !== 0) return cmp;
          }
          return 0;
        });
      }
    }
    return rows;
  }

  function render() {
    const prevScroll = container.querySelector('.dg-scroll');
    const prevScrollTop = prevScroll ? prevScroll.scrollTop : 0;
    const all = filteredRows();
    const totalPages = Math.max(1, Math.ceil(all.length / pageSize));
    page = Math.min(page, totalPages - 1);
    const start = page * pageSize;
    const pageRows = all.slice(start, start + pageSize);

    const defaultWidth = (c) => {
      if (c.width) return c.width;
      if (c.type === 'select') return '150px';
      if (c.type === 'date') return '125px';
      if (c.type === 'number' || c.type === 'percent') return '95px';
      if (c.type === 'money') return '115px';
      if (c.type === 'checkbox') return '80px';
      return '160px'; // text
    };
    const colWidths = columns.map(defaultWidth);
    const delColWidth = onDeleteRow ? 36 : 0;
    const totalWidth = colWidths.reduce((s, w) => s + parseInt(w, 10), 0) + delColWidth;

    const colgroup = '<colgroup>' + colWidths.map((w) => `<col style="width:${w}">`).join('') + (onDeleteRow ? '<col style="width:36px">' : '') + '</colgroup>';

    const head = columns
      .map((c, i) => {
        const lvlIdx = sortLevels.findIndex((s) => s.key === c.key);
        let indicator = '';
        if (lvlIdx > -1) {
          const arrow = sortLevels[lvlIdx].dir === 1 ? '▲' : '▼';
          indicator = ' ' + arrow + (sortLevels.length > 1 ? lvlIdx + 1 : '');
        }
        return `<th data-key="${c.key}" style="width:${colWidths[i]}" title="Click to sort. Shift-click to add as another sort level.">${escapeHtml(c.label)}${indicator}</th>`;
      })
      .join('') + (onDeleteRow ? '<th style="width:36px"></th>' : '');

    let bodyHtml;
    if (pageRows.length === 0) {
      bodyHtml = `<tr><td colspan="${columns.length + 1}"><div class="empty-state"><div class="es-emoji">🗂️</div>${escapeHtml(
        emptyLabel
      )}</div></td></tr>`;
    } else {
      bodyHtml = pageRows
        .map((row, i) => {
          const cls = rowClass ? rowClass(row) || '' : '';
          const cells = columns
            .map((c) => renderCell(row, c))
            .join('');
          const delCell = onDeleteRow
            ? `<td><button class="dg-del-btn" data-del="${row[idKey]}" title="Delete row">✕</button></td>`
            : '';
          return `<tr data-rid="${row[idKey]}" class="${cls}">${cells}${delCell}</tr>`;
        })
        .join('');
    }

    container.innerHTML = `
      <div class="dg-toolbar">
        <input type="text" class="dg-search" placeholder="Search..." value="${escapeHtml(search)}" />
        <span class="dg-count">${all.length} row${all.length === 1 ? '' : 's'}</span>
        ${sortLevels.length ? `<span class="dg-sort-summary">Sorted by ${sortLevels.map((s) => `${columns.find((c) => c.key === s.key)?.label || s.key} ${s.dir === 1 ? '▲' : '▼'}`).join(', then ')}</span><button class="btn btn-light btn-sm dg-clearsort" title="Clear sorting">✕ Clear sort</button>` : ''}
        <span class="topbar-spacer" style="flex:1"></span>
        ${onAddRow ? '<button class="btn btn-light btn-sm dg-add">+ Add row</button>' : ''}
      </div>
      <div class="dg-sort-hint text-muted">Tip: click a column header to sort by it; Shift-click another header to sort by that too (e.g. Priority, then Due Date).</div>
      <div class="dg-scroll">
        <table class="dg-table" style="width:${totalWidth}px; min-width:100%;">
          ${colgroup}
          <thead><tr>${head}</tr></thead>
          <tbody>${bodyHtml}</tbody>
        </table>
      </div>
      <div class="dg-pagination">
        <button class="dg-prev" ${page === 0 ? 'disabled' : ''}>‹ Prev</button>
        <span>Page ${page + 1} of ${totalPages}</span>
        <button class="dg-next" ${page >= totalPages - 1 ? 'disabled' : ''}>Next ›</button>
      </div>
    `;

    const newScroll = container.querySelector('.dg-scroll');
    if (newScroll && prevScrollTop) newScroll.scrollTop = prevScrollTop;

    wireEvents();
  }

  function renderCell(row, c) {
    const raw = c.computed ? c.computed(row) : row[c.key];
    if (!c.editable) {
      return `<td>${escapeHtml(c.format ? c.format(raw, row) : raw)}</td>`;
    }
    if (c.type === 'select') {
      const opts = (c.options || [])
        .map((o) => `<option value="${escapeHtml(o)}" ${o === raw ? 'selected' : ''}>${escapeHtml(o)}</option>`)
        .join('');
      return `<td><select class="cell-input" data-key="${c.key}"><option value=""></option>${opts}</select></td>`;
    }
    if (c.type === 'checkbox') {
      return `<td style="text-align:center"><input type="checkbox" data-key="${c.key}" ${raw ? 'checked' : ''} /></td>`;
    }
    if (c.type === 'date') {
      return `<td><input class="cell-input" type="date" data-key="${c.key}" value="${toISODate(raw) || ''}" /></td>`;
    }
    if (c.type === 'number' || c.type === 'percent') {
      return `<td><input class="cell-input" type="number" step="any" data-key="${c.key}" value="${raw === null || raw === undefined ? '' : raw}" style="text-align:right" /></td>`;
    }
    if (c.type === 'money') {
      // A plain type="number" input can't display comma grouping, so money
      // columns use a text input instead — comma-formatted (e.g. "1,234,567")
      // while not focused, plain digits while typing (wired below).
      return `<td><input class="cell-input money-input" type="text" inputmode="decimal" data-key="${c.key}" value="${raw === null || raw === undefined ? '' : raw}" style="text-align:right" /></td>`;
    }
    return `<td><input class="cell-input" type="text" data-key="${c.key}" value="${escapeHtml(raw ?? '')}" /></td>`;
  }

  function wireEvents() {
    const searchInput = container.querySelector('.dg-search');
    searchInput.addEventListener('input', (e) => {
      search = e.target.value;
      page = 0;
      render();
      // restore focus + caret
      const ni = container.querySelector('.dg-search');
      ni.focus();
      ni.setSelectionRange(ni.value.length, ni.value.length);
    });

    container.querySelectorAll('thead th[data-key]').forEach((th) => {
      th.addEventListener('click', (e) => {
        const key = th.dataset.key;
        const idx = sortLevels.findIndex((s) => s.key === key);
        if (e.shiftKey) {
          if (idx > -1) {
            sortLevels[idx].dir *= -1;
          } else {
            if (sortLevels.length >= MAX_SORT_LEVELS) sortLevels.shift();
            sortLevels.push({ key, dir: 1 });
          }
        } else if (sortLevels.length === 1 && idx === 0) {
          sortLevels[0].dir *= -1;
        } else {
          sortLevels = [{ key, dir: 1 }];
        }
        page = 0;
        render();
      });
    });

    const clearSortBtn = container.querySelector('.dg-clearsort');
    if (clearSortBtn) clearSortBtn.addEventListener('click', () => { sortLevels = []; page = 0; render(); });

    const addBtn = container.querySelector('.dg-add');
    if (addBtn) addBtn.addEventListener('click', () => { onAddRow(); page = 0; render(); });

    container.querySelectorAll('[data-del]').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (!confirm('Delete this row?')) return;
        onDeleteRow(btn.dataset.del);
        render();
      });
    });

    container.querySelectorAll('.dg-prev').forEach((b) => b.addEventListener('click', () => { page = Math.max(0, page - 1); render(); }));
    container.querySelectorAll('.dg-next').forEach((b) => b.addEventListener('click', () => { page += 1; render(); }));

    container.querySelectorAll('tbody tr[data-rid]').forEach((tr) => {
      const rid = tr.dataset.rid;
      const row = getRows().find((r) => String(r[idKey]) === String(rid));
      if (!row) return;
      tr.querySelectorAll('[data-key]').forEach((input) => {
        const key = input.dataset.key;
        const evt = input.tagName === 'SELECT' || input.type === 'checkbox' || input.type === 'date' ? 'change' : 'input';
        input.addEventListener(evt, () => {
          let val = input.type === 'checkbox' ? input.checked : input.value;
          if (input.type === 'number' || input.classList.contains('money-input')) {
            val = val === '' ? null : parseFloat(String(val).replace(/,/g, ''));
          }
          onCellChange(row, key, val);
        });
      });
    });

    container.querySelectorAll('.money-input').forEach((el) => wireMoneyInput(el));
  }

  render();
  return { refresh: render, resetPage: () => { page = 0; } };
}
