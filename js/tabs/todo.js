import { loadTab, markDirty } from '../store.js';
import { uid, daysSince, todayISO, showToast, escapeHtml, toISODate } from '../utils.js';
import { downloadSingleSheet, readWorkbook, parseSheetRows } from '../excel.js';

// The board's columns — task progress, not category. Category, priority,
// dates and notes all live as fields on the card itself instead.
const STATUSES = ['Pending', 'In Progress', 'Done'];
const CATEGORIES = ['Mindset', 'Trainings', 'Selling', 'Recruitment', 'Follow Ups', 'Admin', 'FINSEM', 'Other'];
// Old category values (from before this list changed) get mapped onto the
// closest new one so nothing already typed in gets lost or hidden.
const CATEGORY_ALIASES = { 'Follow-up': 'Follow Ups', 'Follow-Up': 'Follow Ups', 'Followup': 'Follow Ups', 'Personal': 'Other' };
const PRIORITIES = ['Low', 'Medium', 'High'];

// Column defs used only for Excel export/import — the board itself renders
// every field on the card, not from this list.
const COLUMNS = [
  { key: 'dateAdded', header: 'Date Added', aliases: ['Date Added', 'Date'], type: 'date' },
  { key: 'task', header: 'Task', aliases: ['Task', 'To-Do', 'Item'], type: 'text' },
  { key: 'category', header: 'Category', aliases: ['Category'], type: 'text' },
  { key: 'priority', header: 'Priority', aliases: ['Priority'], type: 'text' },
  { key: 'dueDate', header: 'Due Date', aliases: ['Due Date', 'Due'], type: 'date' },
  { key: 'status', header: 'Status', aliases: ['Status', 'Done', 'Complete'], type: 'text' },
  { key: 'notes', header: 'Notes', aliases: ['Notes', 'Remarks'], type: 'text' },
];

function normalizeCategory(cat) {
  const aliased = CATEGORY_ALIASES[cat];
  if (aliased) return aliased;
  return CATEGORIES.includes(cat) ? cat : (cat ? 'Other' : '');
}

// Reads either a modern status string or an older row's boolean/legacy
// "done" value (from before Status replaced the Done checkbox) and settles
// on one of the three board columns.
function normalizeStatus(r) {
  if (STATUSES.includes(r.status)) return r.status;
  if (typeof r.status === 'string' && r.status.trim()) {
    const s = r.status.trim().toLowerCase();
    if (['in progress', 'ongoing', 'doing', 'started', 'in-progress'].includes(s)) return 'In Progress';
    if (['done', 'complete', 'completed', 'true', '1', 'yes'].includes(s)) return 'Done';
    if (['pending', 'open', 'not started', 'todo', 'to do', 'false', '0', 'no'].includes(s)) return 'Pending';
  }
  if (r.done === true || r.done === 'true' || r.done === 'TRUE' || r.done === 1 || r.done === '1') return 'Done';
  return 'Pending';
}

export async function render(container) {
  const d = await loadTab('todo');
  if (!d.rows) d.rows = [];
  d.rows.forEach((r) => { if (!r._id) r._id = uid(); });

  // One-time-per-load normalization: the board now groups by Status
  // (Pending / In Progress / Done) instead of by Category, and the old
  // boolean "Done" checkbox is now a 3-way status. Remap every row rather
  // than dropping anything.
  let changed = false;
  d.rows.forEach((r) => {
    const nextCat = normalizeCategory(r.category);
    if (nextCat !== r.category) { r.category = nextCat; changed = true; }
    const nextStatus = normalizeStatus(r);
    if (nextStatus !== r.status) { r.status = nextStatus; changed = true; }
    if ('done' in r) { delete r.done; changed = true; }
  });
  if (JSON.stringify(d.categoryOptions) !== JSON.stringify(CATEGORIES)) { d.categoryOptions = CATEGORIES.slice(); changed = true; }
  if (JSON.stringify(d.statusOptions) !== JSON.stringify(STATUSES)) { d.statusOptions = STATUSES.slice(); changed = true; }
  if (!d.priorityOptions) { d.priorityOptions = PRIORITIES.slice(); changed = true; }
  if (changed) markDirty('todo');

  let search = '';
  let draggingId = null;

  function persist() { markDirty('todo'); }

  function findRow(id) { return d.rows.find((r) => r._id === id); }

  function visibleRows() {
    if (!search.trim()) return d.rows;
    const q = search.trim().toLowerCase();
    return d.rows.filter((r) => [r.task, r.notes, r.category, r.priority, r.status].some((v) => v && String(v).toLowerCase().includes(q)));
  }

  function cardClasses(r) {
    const cls = ['trello-card'];
    if (r.status === 'Done') { cls.push('todo-done'); return cls.join(' '); }
    const ds = daysSince(r.dueDate);
    if (ds !== null) {
      if (ds > 0) cls.push('overdue-21');
      else if (ds === 0) cls.push('overdue-14');
      else if (ds >= -2) cls.push('overdue-7');
    }
    return cls.join(' ');
  }

  function renderCard(r) {
    const pr = (r.priority || '').toLowerCase();
    return `
      <div class="${cardClasses(r)}" draggable="true" data-id="${r._id}">
        <div class="trello-card-top">
          <select class="trello-priority-select trello-priority-${pr}" data-id="${r._id}" title="Priority">
            <option value="" ${!r.priority ? 'selected' : ''}>No priority</option>
            ${PRIORITIES.map((p) => `<option value="${p}" ${r.priority === p ? 'selected' : ''}>${p}</option>`).join('')}
          </select>
          <button class="dg-del-btn trello-del" data-id="${r._id}" title="Delete this to-do">✕</button>
        </div>
        <textarea class="trello-task" data-id="${r._id}" placeholder="Task...">${escapeHtml(r.task || '')}</textarea>
        <div class="trello-card-fields">
          <select class="trello-status-select" data-id="${r._id}" title="Status">
            ${STATUSES.map((s) => `<option value="${s}" ${r.status === s ? 'selected' : ''}>${s}</option>`).join('')}
          </select>
          <select class="trello-category-select" data-id="${r._id}" title="Category">
            <option value="" ${!r.category ? 'selected' : ''}>No category</option>
            ${CATEGORIES.map((c) => `<option value="${escapeHtml(c)}" ${r.category === c ? 'selected' : ''}>${escapeHtml(c)}</option>`).join('')}
          </select>
        </div>
        <div class="trello-dates">
          <label>Added<input type="date" class="trello-date-added" data-id="${r._id}" value="${toISODate(r.dateAdded) || ''}" /></label>
          <label>Due<input type="date" class="trello-due-date" data-id="${r._id}" value="${toISODate(r.dueDate) || ''}" /></label>
        </div>
        <textarea class="trello-notes" data-id="${r._id}" placeholder="Notes...">${escapeHtml(r.notes || '')}</textarea>
      </div>`;
  }

  function renderColumn(status, rowsInStatus) {
    const slug = status.toLowerCase().replace(/\s+/g, '-');
    return `
      <div class="trello-col trello-col-${slug}">
        <div class="trello-col-header">
          <span>${escapeHtml(status)}</span>
          <span class="trello-count">${rowsInStatus.length}</span>
        </div>
        <div class="trello-cards" data-status="${escapeHtml(status)}">
          ${rowsInStatus.length === 0 ? '<div class="trello-empty">No tasks yet.</div>' : ''}
          ${rowsInStatus.map(renderCard).join('')}
        </div>
        <button class="btn btn-light btn-sm trello-add" data-status="${escapeHtml(status)}">+ Add a card</button>
      </div>`;
  }

  function paint() {
    const rows = visibleRows();
    const open = d.rows.filter((r) => r.status !== 'Done').length;
    const overdue = d.rows.filter((r) => r.status !== 'Done' && daysSince(r.dueDate) > 0).length;
    const done = d.rows.filter((r) => r.status === 'Done').length;

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
      <div class="todo-statbar">
        <div class="stat" style="min-width:110px;"><div class="stat-label">Open</div><div class="stat-value">${open}</div></div>
        <div class="stat" style="min-width:110px;"><div class="stat-label">Overdue</div><div class="stat-value">${overdue}</div></div>
        <div class="stat" style="min-width:110px;"><div class="stat-label">Done</div><div class="stat-value">${done}</div></div>
        <input type="text" id="todoSearch" class="dg-search" placeholder="Search tasks & notes..." value="${escapeHtml(search)}" />
      </div>
      <p class="section-note">🔔 Cards tint amber when due within 2 days, orange when due today, red once overdue — and dim with a strikethrough once marked Done. Drag a card into another column to change its status, or use the Status dropdown on the card. Category, priority, dates and notes are all right there on each card.</p>
      <div class="trello-board">
        ${STATUSES.map((s) => renderColumn(s, rows.filter((r) => r.status === s))).join('')}
      </div>
    `;
    wire();
  }

  function autoGrow(el) {
    el.style.height = 'auto';
    el.style.height = el.scrollHeight + 'px';
  }

  function wire() {
    const searchInput = container.querySelector('#todoSearch');
    searchInput.addEventListener('input', (e) => {
      search = e.target.value;
      paint();
      const ni = container.querySelector('#todoSearch');
      ni.focus();
      ni.setSelectionRange(ni.value.length, ni.value.length);
    });

    container.querySelectorAll('.trello-task, .trello-notes').forEach((el) => {
      autoGrow(el);
      el.addEventListener('input', () => {
        const row = findRow(el.dataset.id);
        if (!row) return;
        row[el.classList.contains('trello-task') ? 'task' : 'notes'] = el.value;
        persist();
        autoGrow(el);
      });
    });

    container.querySelectorAll('.trello-priority-select').forEach((sel) => {
      sel.addEventListener('change', () => {
        const row = findRow(sel.dataset.id);
        if (!row) return;
        row.priority = sel.value;
        persist();
        sel.className = `trello-priority-select trello-priority-${sel.value.toLowerCase()}`;
      });
    });

    container.querySelectorAll('.trello-category-select').forEach((sel) => {
      sel.addEventListener('change', () => {
        const row = findRow(sel.dataset.id);
        if (!row) return;
        row.category = sel.value;
        persist();
        // Category no longer decides which column a card lives in, so no
        // repaint is needed here — this keeps focus/scroll undisturbed.
      });
    });

    container.querySelectorAll('.trello-status-select').forEach((sel) => {
      sel.addEventListener('change', () => {
        const row = findRow(sel.dataset.id);
        if (!row) return;
        row.status = sel.value;
        persist();
        paint(); // status decides the column, so the card needs to move
      });
    });

    container.querySelectorAll('.trello-date-added').forEach((inp) => {
      inp.addEventListener('change', () => {
        const row = findRow(inp.dataset.id);
        if (!row) return;
        row.dateAdded = inp.value;
        persist();
      });
    });

    container.querySelectorAll('.trello-due-date').forEach((inp) => {
      inp.addEventListener('change', () => {
        const row = findRow(inp.dataset.id);
        if (!row) return;
        row.dueDate = inp.value;
        persist();
        paint(); // due-date shift can change the overdue tint, so refresh
      });
    });

    container.querySelectorAll('.trello-del').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (!confirm('Delete this to-do?')) return;
        const i = d.rows.findIndex((r) => r._id === btn.dataset.id);
        if (i > -1) d.rows.splice(i, 1);
        persist();
        paint();
      });
    });

    container.querySelectorAll('.trello-add').forEach((btn) => {
      btn.addEventListener('click', () => {
        const status = btn.dataset.status;
        const r = { _id: uid(), dateAdded: todayISO(), task: '', category: '', priority: '', dueDate: '', status, notes: '' };
        d.rows.unshift(r);
        persist();
        paint();
        const ta = container.querySelector(`.trello-card[data-id="${r._id}"] .trello-task`);
        if (ta) ta.focus();
      });
    });

    // Drag a card into another column to reassign its status — the classic
    // Trello move. The Status dropdown on each card does the same thing for
    // anyone who'd rather not drag (or is on a device where dragging is
    // awkward).
    container.querySelectorAll('.trello-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        draggingId = card.dataset.id;
        card.classList.add('dragging');
        if (e.dataTransfer) {
          e.dataTransfer.effectAllowed = 'move';
          e.dataTransfer.setData('text/plain', card.dataset.id);
        }
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        draggingId = null;
        container.querySelectorAll('.trello-cards.drag-over').forEach((z) => z.classList.remove('drag-over'));
      });
    });

    container.querySelectorAll('.trello-cards').forEach((zone) => {
      zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
      zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        const id = (e.dataTransfer && e.dataTransfer.getData('text/plain')) || draggingId;
        const row = findRow(id);
        const newStatus = zone.dataset.status;
        if (row && newStatus && row.status !== newStatus) {
          row.status = newStatus;
          persist();
          paint();
        }
      });
    });

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
          r.category = normalizeCategory(r.category);
          r.status = normalizeStatus(r);
        });
        if (!confirm(`Import ${rows.length} row(s)? This will replace the current to-do list.`)) return;
        d.rows = rows;
        persist();
        paint();
        showToast(`Imported ${rows.length} row(s) from Excel.`);
      } catch (err) {
        console.error(err);
        showToast('Could not read that file.');
      } finally {
        e.target.value = '';
      }
    });
  }

  paint();
}
