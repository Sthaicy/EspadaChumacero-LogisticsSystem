// ── TOAST ──────────────────────────────────────────────────────────────────────
function toast(msg, type = 'success') {
  let c = document.getElementById('toast-container');
  if (!c) { c = document.createElement('div'); c.id = 'toast-container'; document.body.appendChild(c); }
  const t = document.createElement('div');
  const icons = { success: '✓', error: '✕', info: 'ℹ' };
  t.className = `toast ${type}`;
  t.innerHTML = `<span style="font-size:16px">${icons[type]||'•'}</span><span>${msg}</span>`;
  c.appendChild(t);
  setTimeout(() => {
    t.style.opacity = '0'; t.style.transform = 'translateX(20px)'; t.style.transition = 'all .3s';
    setTimeout(() => t.remove(), 300);
  }, 3200);
}

// ── MODAL ──────────────────────────────────────────────────────────────────────
function openModal(id) { document.getElementById(id).classList.add('open'); }

function closeModal(id) {
  document.getElementById(id).classList.remove('open');
  const form = document.querySelector(`#${id} form`);
  if (form) form.reset();
}

document.addEventListener('click', e => {
  if (e.target.classList.contains('modal-overlay')) {
    e.target.classList.remove('open');
    const form = e.target.querySelector('form');
    if (form) form.reset();
  }
});

document.addEventListener('keydown', e => {
  if (e.key === 'Escape')
    document.querySelectorAll('.modal-overlay.open').forEach(m => {
      m.classList.remove('open');
      const form = m.querySelector('form');
      if (form) form.reset();
    });
});

// ── AJAX ───────────────────────────────────────────────────────────────────────
async function submitForm(url, formId, successMsg, reloadOnSuccess = true) {
  const form = document.getElementById(formId);
  const fd = new FormData(form);
  try {
    const r = await fetch(url, { method: 'POST', body: fd });
    const j = await r.json();
    if (j.ok) {
      toast(successMsg, 'success');
      form.reset();
      if (reloadOnSuccess) setTimeout(() => location.reload(), 700);
    } else {
      toast(j.msg || 'Error al procesar', 'error');
    }
    return j;
  } catch { toast('Error de conexión', 'error'); }
}

async function deleteRow(url, msg = '¿Eliminar este registro?') {
  if (!confirm(msg)) return;
  try {
    const r = await fetch(url, { method: 'POST' });
    const j = await r.json();
    if (j.ok) { toast('Registro eliminado', 'success'); setTimeout(() => location.reload(), 600); }
    else toast('Error al eliminar', 'error');
  } catch { toast('Error de conexión', 'error'); }
}

// ── TABLE FILTER ───────────────────────────────────────────────────────────────
function initTableFilter(tableId) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const thead = table.querySelector('thead');
  const headers = thead.querySelectorAll('th:not(.no-filter)');
  const filterRow = document.createElement('tr');
  filterRow.className = 'filter-row';

  // Build one th per original th (including no-filter ones)
  const allTh = thead.querySelectorAll('tr:first-child th');
  allTh.forEach((th, i) => {
    const td = document.createElement('th');
    if (!th.classList.contains('no-filter')) {
      const inp = document.createElement('input');
      inp.type = 'text'; inp.placeholder = '🔍'; inp.dataset.col = i;
      inp.addEventListener('input', () => filterTable(tableId));
      td.appendChild(inp);
    }
    filterRow.appendChild(td);
  });
  thead.appendChild(filterRow);

  // Sort on click
  allTh.forEach((th, i) => {
    if (!th.classList.contains('no-filter'))
      th.addEventListener('click', () => sortTable(tableId, i));
  });
}

function filterTable(tableId) {
  const table = document.getElementById(tableId);
  const inputs = table.querySelectorAll('.filter-row input');
  table.querySelectorAll('tbody tr').forEach(row => {
    if (row.classList.contains('grupo-header')) return;
    let visible = true;
    inputs.forEach(inp => {
      const col = parseInt(inp.dataset.col);
      const cell = row.cells[col];
      if (cell && inp.value && !cell.textContent.toLowerCase().includes(inp.value.toLowerCase()))
        visible = false;
    });
    row.style.display = visible ? '' : 'none';
  });
  // Hide group headers if all children hidden
  table.querySelectorAll('.grupo-header').forEach(hdr => {
    const cod = hdr.dataset.grupo;
    if (!cod) return;
    const hijos = table.querySelectorAll(`tbody tr[data-cod="${cod}"]`);
    const alguno = Array.from(hijos).some(r => r.style.display !== 'none');
    hdr.style.display = alguno ? '' : 'none';
  });
}

let sortState = {};
function sortTable(tableId, col) {
  const table = document.getElementById(tableId);
  const tbody = table.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr:not(.grupo-header)'));
  const key = `${tableId}-${col}`;
  sortState[key] = sortState[key] === 'asc' ? 'desc' : 'asc';
  const dir = sortState[key];
  rows.sort((a, b) => {
    const va = a.cells[col]?.textContent.trim() || '';
    const vb = b.cells[col]?.textContent.trim() || '';
    const na = parseFloat(va.replace(/[^0-9.-]/g, ''));
    const nb = parseFloat(vb.replace(/[^0-9.-]/g, ''));
    if (!isNaN(na) && !isNaN(nb)) return dir === 'asc' ? na - nb : nb - na;
    return dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
  });
  rows.forEach(r => tbody.appendChild(r));
}

function globalSearch(inputId, tableId) {
  const q = document.getElementById(inputId).value.toLowerCase();
  document.querySelectorAll(`#${tableId} tbody tr`).forEach(r => {
    if (r.classList.contains('grupo-header')) return;
    r.style.display = r.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
  // re-hide grupo headers
  document.querySelectorAll(`#${tableId} .grupo-header`).forEach(hdr => {
    const cod = hdr.dataset.grupo;
    if (!cod) return;
    const hijos = document.querySelectorAll(`#${tableId} tbody tr[data-cod="${cod}"]`);
    const alguno = Array.from(hijos).some(r => r.style.display !== 'none');
    hdr.style.display = alguno ? '' : 'none';
  });
}

// ── POPULATE EDIT ──────────────────────────────────────────────────────────────
function populateEdit(modalId, data) {
  const form = document.querySelector(`#${modalId} form`);
  if (!form) return;
  Object.entries(data).forEach(([k, v]) => {
    const el = form.querySelector(`[name="${k}"]`);
    if (el) el.value = v ?? '';
  });
  openModal(modalId);
}

// ── DATE ───────────────────────────────────────────────────────────────────────
function updateDateTime() {
  const el = document.getElementById('topbar-date');
  if (!el) return;
  const now = new Date();
  el.textContent = now.toLocaleDateString('es-BO', { weekday:'short', year:'numeric', month:'short', day:'numeric' });
}
updateDateTime();

// ── SIDEBAR ACTIVE ─────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  const path = window.location.pathname.split('/')[1] || '';
  document.querySelectorAll('.nav-item').forEach(a => {
    const href = (a.getAttribute('href') || '').replace(/^\//, '');
    if (href && path.startsWith(href)) a.classList.add('active');
    if (href === '' && path === '') a.classList.add('active');
  });
});

// ── CALC TOTAL ─────────────────────────────────────────────────────────────────
function calcTotal(cantId, precioId, totalId) {
  const cant  = parseFloat(document.getElementById(cantId)?.value) || 0;
  const precio = parseFloat(document.getElementById(precioId)?.value) || 0;
  const el = document.getElementById(totalId);
  if (el) el.value = (cant * precio).toFixed(2);
}

// ── RECALC TOTALS HELPER ───────────────────────────────────────────────────────
function recalcTotals(config) {
  // config: { tableId, cols: [{id, colIndex}], countId, badgeId, badgeLabel }
  const { tableId, cols, countId, badgeId, badgeLabel } = config;
  const sums = {};
  cols.forEach(c => sums[c.id] = 0);
  let count = 0;

  document.querySelectorAll(`#${tableId} tbody tr`).forEach(row => {
    if (row.classList.contains('grupo-header') || row.style.display === 'none') return;
    const cells = row.querySelectorAll('td');
    if (cells.length < 2) return;
    cols.forEach(c => {
      sums[c.id] += parseFloat(cells[c.colIndex]?.textContent.trim()) || 0;
    });
    count++;
  });

  cols.forEach(c => {
    const el = document.getElementById(c.id);
    if (el) el.textContent = 'Bs. ' + sums[c.id].toFixed(2);
  });
  if (countId) { const el = document.getElementById(countId); if (el) el.textContent = count; }
  if (badgeId)  { const el = document.getElementById(badgeId);  if (el) el.textContent = count + ' ' + (badgeLabel||'registros'); }
}

// ── TOM SELECT — autocomplete en todos los <select> ───────────────────────────
// Llama initSelects() después de que el DOM cargue para activar búsqueda
// en todos los <select> del documento (excepto los del filtro de tabla).
function initSelects(root) {
  const scope = root || document;
  scope.querySelectorAll('select:not(.ts-hidden-accessible):not(.filter-row select)').forEach(sel => {
    if (sel.tomselect) return; // ya inicializado
    new TomSelect(sel, {
      allowEmptyOption: true,
      placeholder: sel.options[0]?.text || '— Seleccionar —',
      maxOptions: 200,
      render: {
        no_results: () => '<div class="no-results" style="padding:10px 14px;color:#999">Sin resultados</div>'
      }
    });
  });
}

// Inicializar al abrir cualquier modal
const _origOpenModal = openModal;
openModal = function(id) {
  _origOpenModal(id);
  const modal = document.getElementById(id);
  if (modal) setTimeout(() => initSelects(modal), 50);
};

// Inicializar los que están visibles en la página (no en modales)
document.addEventListener('DOMContentLoaded', () => {
  // No aplicar a selects dentro de modales (se inicia al abrir)
  document.querySelectorAll('.page-content select, .toolbar select').forEach(sel => {
    if (!sel.closest('.modal-overlay') && !sel.tomselect) {
      new TomSelect(sel, {
        allowEmptyOption: true,
        maxOptions: 200,
        render: {
          no_results: () => '<div class="no-results" style="padding:10px 14px;color:#999">Sin resultados</div>'
        }
      });
    }
  });
});
