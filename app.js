// ── Estado global ────────────────────────────────────────
let orders = [];
let currentDate = new Date();
const today = new Date();
today.setHours(0, 0, 0, 0);

const activeFilters = { ok: true, late: true, today: true };
let searchQuery   = '';
let tableFilter   = 'all'; // 'all' | 'late' | 'ok'
let currentView   = 'calendario';

const MONTHS = [
  'Enero','Febrero','Marzo','Abril','Mayo','Junio',
  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
];
const DAYS_OF_WEEK = ['D','L','M','MI','J','V','S'];

const DONE_KEYWORDS = [
  'completado','completada','entregado','entregada',
  'finalizado','finalizada','done','complete',
  'cerrado','cerrada','terminado','terminada'
];

// ── Navegación ───────────────────────────────────────────
function switchView(view) {
  currentView = view;
  document.getElementById('view-calendario').classList.toggle('hidden', view !== 'calendario');
  document.getElementById('view-ordenes').classList.toggle('hidden', view !== 'ordenes');
  document.getElementById('tab-calendario').classList.toggle('active', view === 'calendario');
  document.getElementById('tab-ordenes').classList.toggle('active', view === 'ordenes');
  if (view === 'ordenes') renderTable();
}

// ── Utilidades de fecha ──────────────────────────────────
function parseDate(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    const parsed = XLSX.SSF.parse_date_code(val);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const s = String(val).trim();
  const formats = [
    { re: /^(\d{4})-(\d{2})-(\d{2})$/,        fn: m => new Date(+m[1], +m[2]-1, +m[3]) },
    { re: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,  fn: m => new Date(+m[3], +m[2]-1, +m[1]) },
    { re: /^(\d{2})-(\d{2})-(\d{4})$/,         fn: m => new Date(+m[3], +m[2]-1, +m[1]) },
  ];
  for (const { re, fn } of formats) {
    const m = s.match(re);
    if (m) return fn(m);
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function formatDate(date) {
  if (!date) return '—';
  return date.toLocaleDateString('es', { day: '2-digit', month: 'short', year: 'numeric' });
}

function toMidnight(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

// ── Lógica de atraso ─────────────────────────────────────
function isLate(order) {
  if (!order.fecha) return false;
  const d   = toMidnight(order.fecha);
  const est = String(order.estado || '').toLowerCase();
  return !DONE_KEYWORDS.some(k => est.includes(k)) && d < today;
}

// ── Normalizar filas del Excel ───────────────────────────
function normalizeOrders(rows) {
  if (!rows.length) return [];
  const KEY_MAP = {
    'numero de orden':'num','número de orden':'num','num_orden':'num',
    'n_orden':'num','numero':'num','no. orden':'num','no orden':'num',
    'numero orden':'num','nro orden':'num','nro. orden':'num',
    'estado de orden':'estado','estado_orden':'estado','estado':'estado',
    'status':'estado','estado orden':'estado',
    'orden':'fecha','fecha':'fecha','fecha orden':'fecha',
    'fecha_orden':'fecha','date':'fecha','delivery':'fecha',
    'fecha entrega':'fecha','fecha de orden':'fecha',
  };
  const firstRow = rows[0];
  const mapping  = {};
  Object.keys(firstRow).forEach(h => {
    const key = KEY_MAP[h.toLowerCase().trim()];
    if (key && !mapping[key]) mapping[key] = h;
  });
  const allH = Object.keys(firstRow);
  if (!mapping.num)    mapping.num    = allH[0];
  if (!mapping.estado) mapping.estado = allH[1];
  if (!mapping.fecha)  mapping.fecha  = allH[2];

  return rows.map((row, i) => ({
    id:     i,
    num:    row[mapping.num]    || `ORD-${i+1}`,
    estado: row[mapping.estado] || 'Pendiente',
    fecha:  parseDate(row[mapping.fecha]),
  })).filter(o => o.fecha);
}

// ── Filtros de leyenda (calendario) ─────────────────────
function toggleFilter(type) {
  const others = Object.keys(activeFilters).filter(k => k !== type);
  if (!others.some(k => activeFilters[k]) && activeFilters[type]) return;
  activeFilters[type] = !activeFilters[type];
  const btn = document.getElementById('filter-' + type);
  btn.classList.toggle('active', activeFilters[type]);
  btn.setAttribute('aria-pressed', activeFilters[type]);
  renderCalendar();
}

// ── Filtro rápido tabla ───────────────────────────────────
function setTableFilter(f) {
  tableFilter = f;
  ['all','late','ok'].forEach(id => {
    document.getElementById('tf-' + id).classList.toggle('active', id === f);
  });
  renderTable();
}

// ── Renderizar calendario ────────────────────────────────
function renderCalendar() {
  const y = currentDate.getFullYear();
  const m = currentDate.getMonth();

  document.getElementById('cal-month-lbl').textContent = MONTHS[m];
  document.getElementById('cal-year-lbl').textContent  = y;

  const monthOrders = orders.filter(o =>
    o.fecha && o.fecha.getFullYear() === y && o.fecha.getMonth() === m
  );
  document.getElementById('cal-units-lbl').textContent =
    monthOrders.length ? `· ${monthOrders.length} órdenes` : '';

  const grid = document.getElementById('cal-grid');
  grid.innerHTML = '';

  DAYS_OF_WEEK.forEach(d => {
    const el = document.createElement('div');
    el.className = 'cal-dow';
    el.textContent = d;
    grid.appendChild(el);
  });

  const firstDay      = new Date(y, m, 1).getDay();
  const daysInMonth   = new Date(y, m+1, 0).getDate();
  const prevMonthDays = new Date(y, m, 0).getDate();

  for (let i = 0; i < firstDay; i++) {
    const el = document.createElement('div');
    el.className = 'cal-day empty';
    const num = document.createElement('div');
    num.className = 'day-num';
    num.textContent = prevMonthDays - firstDay + 1 + i;
    el.appendChild(num);
    grid.appendChild(el);
  }

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(y, m, d);
    date.setHours(0, 0, 0, 0);
    const isToday = date.getTime() === today.getTime();

    const allDayOrders = orders.filter(o =>
      o.fecha &&
      o.fecha.getFullYear() === y &&
      o.fecha.getMonth()    === m &&
      o.fecha.getDate()     === d
    );

    const hasLate = allDayOrders.some(isLate);
    const hasOk   = allDayOrders.some(o => !isLate(o));

    let dayVisible = !allDayOrders.length; // días vacíos siempre visibles
    if (isToday && activeFilters.today) dayVisible = true;
    if (hasLate  && activeFilters.late) dayVisible = true;
    if (hasOk    && activeFilters.ok)   dayVisible = true;

    const el  = document.createElement('div');
    let   cls = 'cal-day';

    if (!dayVisible) {
      cls += ' filtered-out';
    } else {
      if (isToday) cls += ' today';
      if      (hasLate && activeFilters.late) cls += ' has-late';
      else if (hasOk   && activeFilters.ok)   cls += ' has-orders';
    }
    el.className = cls;

    const numEl = document.createElement('div');
    numEl.className = 'day-num';
    numEl.textContent = d;
    el.appendChild(numEl);

    if (dayVisible && allDayOrders.length) {
      const visible = allDayOrders.filter(o => isLate(o) ? activeFilters.late : activeFilters.ok);
      visible.slice(0, 3).forEach(o => {
        const dot = document.createElement('div');
        dot.className = 'day-dot ' + (isLate(o) ? 'dot-red' : 'dot-green');
        dot.textContent = o.num;
        el.appendChild(dot);
      });
      if (visible.length > 3) {
        const more = document.createElement('div');
        more.className = 'day-dot dot-more';
        more.textContent = `+${visible.length - 3} más`;
        el.appendChild(more);
      }
    }

    el.addEventListener('click', () => {
      if (el.classList.contains('filtered-out')) return;
      showDayDetail(d, y, m, allDayOrders);
    });
    grid.appendChild(el);
  }

  const remaining = (7 - ((firstDay + daysInMonth) % 7)) % 7;
  for (let i = 1; i <= remaining; i++) {
    const el = document.createElement('div');
    el.className = 'cal-day empty';
    const num = document.createElement('div');
    num.className = 'day-num';
    num.textContent = i;
    el.appendChild(num);
    grid.appendChild(el);
  }
}

// ── Detalle al hacer clic en un día ─────────────────────
function showDayDetail(d, y, m, dayOrders) {
  document.getElementById('detail-title').textContent = `${d} de ${MONTHS[m]} ${y}`;
  const wrap = document.getElementById('detail-wrap');
  if (!dayOrders.length) {
    wrap.innerHTML = '<div class="no-data" style="padding:1rem">No hay órdenes este día</div>';
    return;
  }
  wrap.innerHTML = dayOrders.map(o => {
    const late = isLate(o);
    const cls  = late ? 'late' : 'ok';
    return `<div class="detail-order ${cls}">
      <div class="detail-order-num ${cls}">${escHtml(o.num)}</div>
      <div class="detail-order-est ${cls}">${escHtml(o.estado)}</div>
      ${late ? '<div class="detail-order-tag"><i class="ti ti-clock"></i> Atrasada</div>' : ''}
    </div>`;
  }).join('');
}

// ── Panel de alertas ─────────────────────────────────────
function renderAlerts() {
  const lateOrders = orders.filter(isLate);
  const byDay = {};
  lateOrders.forEach(o => {
    const key = formatDate(o.fecha);
    if (!byDay[key]) byDay[key] = [];
    byDay[key].push(o);
  });
  const wrap = document.getElementById('alerts-wrap');
  if (!lateOrders.length) {
    wrap.innerHTML = `<div class="alert-empty">
      <i class="ti ti-check" style="font-size:20px;display:block;margin-bottom:4px"></i>
      Sin atrasos detectados</div>`;
    return;
  }
  wrap.innerHTML = Object.entries(byDay)
    .sort((a, b) => new Date(a[0]) - new Date(b[0]))
    .map(([day, ords]) => `<div class="alert-item">
      <div class="alert-item-top">
        <i class="ti ti-alert-circle" style="font-size:14px;color:#A32D2D"></i>
        <span class="alert-item-label">${escHtml(day)}</span>
        <span class="badge badge-red" style="margin-left:auto">${ords.length} ord.</span>
      </div>
      <div class="alert-item-sub">
        ${ords.slice(0,3).map(o => escHtml(o.num)).join(', ')}
        ${ords.length > 3 ? ` +${ords.length-3} más` : ''}
      </div>
    </div>`).join('');
}

// ── Búsqueda ─────────────────────────────────────────────
function onSearch(val) {
  searchQuery = val.trim().toLowerCase();
  document.getElementById('search-clear').style.display = searchQuery ? 'flex' : 'none';
  renderTable();
}

function clearSearch() {
  searchQuery = '';
  document.getElementById('search-input').value = '';
  document.getElementById('search-clear').style.display = 'none';
  renderTable();
}

// ── Tabla ────────────────────────────────────────────────
function highlight(text, query) {
  if (!query) return escHtml(text);
  const escaped = escHtml(text);
  const idx = escaped.toLowerCase().indexOf(query.toLowerCase());
  if (idx === -1) return escaped;
  return escaped.slice(0, idx) +
    `<mark class="hl">${escaped.slice(idx, idx + query.length)}</mark>` +
    escaped.slice(idx + query.length);
}

function renderTable() {
  let fo = orders.slice();

  // Filtro rápido de la vista órdenes
  if (tableFilter === 'late') fo = fo.filter(isLate);
  if (tableFilter === 'ok')   fo = fo.filter(o => !isLate(o));

  // Búsqueda de texto
  if (searchQuery) {
    fo = fo.filter(o => {
      return String(o.num    || '').toLowerCase().includes(searchQuery) ||
             String(o.estado || '').toLowerCase().includes(searchQuery) ||
             formatDate(o.fecha).toLowerCase().includes(searchQuery);
    });
  }

  const label = searchQuery
    ? `${fo.length} resultado${fo.length !== 1 ? 's' : ''} de ${orders.length}`
    : `Todas las órdenes (${fo.length})`;
  document.getElementById('table-title').textContent = label;

  if (!fo.length) {
    document.getElementById('table-wrap').innerHTML = searchQuery
      ? `<div class="no-data"><i class="ti ti-search-off no-data-icon"></i>Sin resultados para "<strong>${escHtml(searchQuery)}</strong>"</div>`
      : `<div class="no-data"><i class="ti ti-table-off no-data-icon"></i>Carga un archivo Excel para ver las órdenes</div>`;
    return;
  }

  const rows = fo.slice(0, 200).map(o => {
    const late   = isLate(o);
    const est    = String(o.estado || '').toLowerCase();
    const isDone = DONE_KEYWORDS.some(k => est.includes(k));
    const pillCls = late ? 'pill-late' : (isDone ? 'pill-ok' : 'pill-pend');
    return `<tr class="${late ? 'late-row' : ''}">
      <td>${highlight(String(o.num), searchQuery)}</td>
      <td><span class="status-pill ${pillCls}">${highlight(o.estado, searchQuery)}</span></td>
      <td>${highlight(formatDate(o.fecha), searchQuery)}</td>
      <td>${late
        ? '<span class="status-pill pill-late">⚠ Atrasada</span>'
        : '<span class="status-pill pill-ok">✓ Al día</span>'}</td>
    </tr>`;
  }).join('');

  document.getElementById('table-wrap').innerHTML = `
    <table class="orders-table">
      <thead><tr>
        <th style="width:28%">N° orden</th>
        <th style="width:24%">Estado</th>
        <th style="width:24%">Fecha</th>
        <th style="width:24%">Alerta</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    ${fo.length > 200 ? `<div class="table-hint">Mostrando 200 de ${fo.length} órdenes</div>` : ''}`;
}

// ── Resumen navbar ───────────────────────────────────────
function renderSummary() {
  const late = orders.filter(isLate).length;
  const ok   = orders.length - late;
  document.getElementById('summary-badges').innerHTML = `
    <span class="badge badge-green"><i class="ti ti-check"></i> ${ok} al día</span>
    ${late ? `<span class="badge badge-red"><i class="ti ti-alert-circle"></i> ${late} atrasadas</span>` : ''}`;

  const badge = document.getElementById('nav-badge-total');
  badge.textContent = orders.length;
  badge.style.display = orders.length ? 'inline-flex' : 'none';
}

// ── Procesar órdenes ─────────────────────────────────────
function processOrders(data) {
  orders = normalizeOrders(data);
  renderCalendar();
  renderAlerts();
  renderSummary();
  if (currentView === 'ordenes') renderTable();
}

// ── Cambiar mes ──────────────────────────────────────────
function changeMonth(dir) {
  currentDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + dir, 1);
  renderCalendar();
}

// ── Escape HTML ──────────────────────────────────────────
function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Leer archivo Excel ───────────────────────────────────
document.getElementById('file-inp').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById('file-status');
  statusEl.textContent = 'Cargando…';
  statusEl.style.color = 'var(--text-tertiary)';

  const reader = new FileReader();
  reader.onload = function (ev) {
    try {
      const wb   = XLSX.read(ev.target.result, { type: 'array', cellDates: false });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { raw: true, defval: '' });
      processOrders(data);
      statusEl.textContent = `✓ ${file.name} (${data.length} filas)`;
      statusEl.style.color = 'var(--green-dark)';
    } catch (err) {
      statusEl.textContent = 'Error al leer el archivo.';
      statusEl.style.color = 'var(--red-dark)';
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
  this.value = '';
});

// ── Drag & Drop ──────────────────────────────────────────
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('drop-area-mini').classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (!file) return;
  const dt = new DataTransfer();
  dt.items.add(file);
  document.getElementById('file-inp').files = dt.files;
  document.getElementById('file-inp').dispatchEvent(new Event('change'));
}

// ── Inicializar ──────────────────────────────────────────
renderCalendar();
