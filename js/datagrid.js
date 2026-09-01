import { escapeHtml, toISODate } from './utils.js';

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
  let sortKey = null;
  let sortDir = 1;
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
    if (sortKey) {
      const col = columns.find((c) => c.key === sortKey);
      rows = rows.slice().sort((a, b) => {
        const va = col.computed ? col.computed(a) : a[sortKey];
        const vb = col.computed ? col.computed(b) : b[sortKey];
        if (va === vb) return 0;
        if (va === null || va === undefined || va === '') return 1;
        if (vb === null || vb === undefined || vb === '') return -1;
        return (va > vb ? 1 : -1) * sortDir;
      });
    }
    return rows;
  }

  function render() {
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
      if (c.type === 'checkbox') return '80px';
      return '160px'; // text
    };
    const colWidths = columns.map(defaultWidth);
    const delColWidth = onDeleteRow ? 36 : 0;
    const totalWidth = colWidths.reduce((s, w) => s + parseInt(w, 10), 0) + delColWidth;

    const colgroup = '<colgroup>' + colWidths.map((w) => `<col style="width:${w}">`).join('') + (onDeleteRow ? '<col style="width:36px">' : '') + '</colgroup>';

    const head = columns
      .map(
        (c, i) =>
          `<th data-key="${c.key}" style="width:${colWidths[i]}">${escapeHtml(c.label)}${
            sortKey === c.key ? (sortDir === 1 ? ' ▲' : ' ▼') : ''
          }</th>`
      )
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
        <span class="topbar-spacer" style="flex:1"></span>
        ${onAddRow ? '<button class="btn btn-light btn-sm dg-add">+ Add row</button>' : ''}
      </div>
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
      th.addEventListener('click', () => {
        const key = th.dataset.key;
        if (sortKey === key) sortDir *= -1;
        else { sortKey = key; sortDir = 1; }
        render();
      });
    });

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
          if (input.type === 'number') val = val === '' ? null : parseFloat(val);
          onCellChange(row, key, val);
        });
      });
    });
  }

  render();
  return { refresh: render, resetPage: () => { page = 0; } };
}
