// ── Estado global ────────────────────────────────────────
let orders = [];
let currentDate = new Date();
const today = new Date();
today.setHours(0, 0, 0, 0);

// Filtros activos: qué tipos de días se muestran en el calendario
const activeFilters = { ok: true, late: true, today: true };

const MONTHS = [
  'Enero','Febrero','Marzo','Abril','Mayo','Junio',
  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
];
const DAYS_OF_WEEK = ['D','L','M','MI','J','V','S'];

// ── Utilidades de fecha ──────────────────────────────────
function parseDate(val) {
  if (!val) return null;

  // Número serial de Excel
  if (typeof val === 'number') {
    const parsed = XLSX.SSF.parse_date_code(val);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const s = String(val).trim();

  // Formatos: YYYY-MM-DD / DD/MM/YYYY / DD-MM-YYYY
  const formats = [
    { re: /^(\d{4})-(\d{2})-(\d{2})$/, fn: m => new Date(+m[1], +m[2] - 1, +m[3]) },
    { re: /^(\d{2})\/(\d{2})\/(\d{4})$/, fn: m => new Date(+m[3], +m[2] - 1, +m[1]) },
    { re: /^(\d{2})-(\d{2})-(\d{4})$/, fn: m => new Date(+m[3], +m[2] - 1, +m[1]) },
    { re: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, fn: m => new Date(+m[3], +m[2] - 1, +m[1]) },
  ];

  for (const { re, fn } of formats) {
    const m = s.match(re);
    if (m) return fn(m);
  }

  // Intento genérico
  const d = new Date(s);
  if (!isNaN(d)) return d;
  return null;
}

function formatDate(date) {
  if (!date) return '—';
  return date.toLocaleDateString('es', { day: '2-digit', month: 'short', year: 'numeric' });
}

function toMidnight(date) {
  if (!date) return null;
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

// ── Lógica de atraso ─────────────────────────────────────
const DONE_KEYWORDS = [
  'completado','completada','entregado','entregada',
  'finalizado','finalizada','done','complete',
  'cerrado','cerrada','terminado','terminada'
];

function isLate(order) {
  if (!order.fecha) return false;
  const d = toMidnight(order.fecha);
  const est = String(order.estado || '').toLowerCase();
  const isDone = DONE_KEYWORDS.some(k => est.includes(k));
  return !isDone && d < today;
}

// ── Normalizar filas del Excel ───────────────────────────
function normalizeOrders(rows) {
  if (!rows.length) return [];

  // Mapa de columnas conocidas → clave interna
  const KEY_MAP = {
    'numero de orden': 'num', 'número de orden': 'num',
    'num_orden': 'num', 'n_orden': 'num',
    'numero': 'num', 'no. orden': 'num', 'no orden': 'num',
    'numero orden': 'num', 'nro orden': 'num', 'nro. orden': 'num',

    'estado de orden': 'estado', 'estado_orden': 'estado',
    'estado': 'estado', 'status': 'estado', 'estado orden': 'estado',

    'orden': 'fecha', 'fecha': 'fecha',
    'fecha orden': 'fecha', 'fecha_orden': 'fecha',
    'date': 'fecha', 'delivery': 'fecha', 'fecha entrega': 'fecha',
    'fecha de orden': 'fecha',
  };

  // Mapear headers reales → clave interna
  const firstRow = rows[0];
  const mapping = {}; // { num: 'Numero de Orden', estado: 'Estado', fecha: 'Orden' }

  Object.keys(firstRow).forEach(header => {
    const key = KEY_MAP[header.toLowerCase().trim()];
    if (key && !mapping[key]) mapping[key] = header;
  });

  // Fallback: tomar columnas por posición si no se detectaron
  const allHeaders = Object.keys(firstRow);
  if (!mapping.num   && allHeaders[0]) mapping.num    = allHeaders[0];
  if (!mapping.estado && allHeaders[1]) mapping.estado = allHeaders[1];
  if (!mapping.fecha  && allHeaders[2]) mapping.fecha  = allHeaders[2];

  return rows
    .map((row, i) => {
      const fecha = parseDate(row[mapping.fecha]);
      return {
        id:     i,
        num:    row[mapping.num]    || `ORD-${i + 1}`,
        estado: row[mapping.estado] || 'Pendiente',
        fecha,
      };
    })
    .filter(o => o.fecha); // descartar filas sin fecha válida
}

// ── Filtros de leyenda ───────────────────────────────────
function toggleFilter(type) {
  // Al menos un filtro debe quedar activo
  const others = Object.keys(activeFilters).filter(k => k !== type);
  const anyOtherActive = others.some(k => activeFilters[k]);
  if (!anyOtherActive && activeFilters[type]) return; // no desactivar el último

  activeFilters[type] = !activeFilters[type];

  // Actualizar aspecto del botón
  const btn = document.getElementById('filter-' + type);
  btn.classList.toggle('active', activeFilters[type]);
  btn.setAttribute('aria-pressed', activeFilters[type]);

  // Re-renderizar calendario y tabla con el nuevo filtro
  renderCalendar();
  renderTable();
}

// ── Renderizar calendario ────────────────────────────────
function renderCalendar() {
  const y = currentDate.getFullYear();
  const m = currentDate.getMonth();

  document.getElementById('cal-month-lbl').textContent = MONTHS[m];
  document.getElementById('cal-year-lbl').textContent = y;

  const monthOrders = orders.filter(o =>
    o.fecha && o.fecha.getFullYear() === y && o.fecha.getMonth() === m
  );
  document.getElementById('cal-units-lbl').textContent =
    monthOrders.length ? `${monthOrders.length} órdenes` : '';

  const grid = document.getElementById('cal-grid');
  grid.innerHTML = '';

  // Encabezados días semana
  DAYS_OF_WEEK.forEach(d => {
    const el = document.createElement('div');
    el.className = 'cal-dow';
    el.textContent = d;
    grid.appendChild(el);
  });

  const firstDay      = new Date(y, m, 1).getDay();
  const daysInMonth   = new Date(y, m + 1, 0).getDate();
  const prevMonthDays = new Date(y, m, 0).getDate();

  // Días del mes anterior (relleno)
  for (let i = 0; i < firstDay; i++) {
    const el = document.createElement('div');
    el.className = 'cal-day empty';
    const num = document.createElement('div');
    num.className = 'day-num';
    num.textContent = prevMonthDays - firstDay + 1 + i;
    el.appendChild(num);
    grid.appendChild(el);
  }

  // Días del mes actual
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(y, m, d);
    date.setHours(0, 0, 0, 0);

    const isToday = date.getTime() === today.getTime();

    // Todas las órdenes de este día
    const allDayOrders = orders.filter(o =>
      o.fecha &&
      o.fecha.getFullYear() === y &&
      o.fecha.getMonth()    === m &&
      o.fecha.getDate()     === d
    );

    const hasLate = allDayOrders.some(isLate);
    const hasOk   = allDayOrders.some(o => !isLate(o));

    // Determinar si este día debe mostrarse según filtros activos
    // Un día es visible si cumple ALGUNA condición activa que le aplique
    let dayVisible = false;
    if (isToday && activeFilters.today) dayVisible = true;
    if (hasLate  && activeFilters.late) dayVisible = true;
    if (hasOk    && activeFilters.ok)   dayVisible = true;
    // Días sin órdenes: se muestran siempre (sólo como celda vacía)
    if (!allDayOrders.length) dayVisible = true;

    const el = document.createElement('div');
    let cls = 'cal-day';

    if (!dayVisible) {
      // Día oculto por filtro: mostrarlo apagado
      cls += ' filtered-out';
    } else {
      if (isToday) cls += ' today';
      if (hasLate && activeFilters.late)     cls += ' has-late';
      else if (hasOk && activeFilters.ok)    cls += ' has-orders';
    }
    el.className = cls;

    // Número del día
    const numEl = document.createElement('div');
    numEl.className = 'day-num';
    numEl.textContent = d;
    el.appendChild(numEl);

    // Puntos de órdenes — filtrar según filtros activos
    if (dayVisible && allDayOrders.length) {
      const visibleOrders = allDayOrders.filter(o => {
        if (isLate(o)) return activeFilters.late;
        return activeFilters.ok;
      });

      visibleOrders.slice(0, 2).forEach(o => {
        const dot = document.createElement('div');
        dot.className = 'day-dot ' + (isLate(o) ? 'dot-red' : 'dot-green');
        dot.textContent = o.num;
        el.appendChild(dot);
      });
      if (visibleOrders.length > 2) {
        const more = document.createElement('div');
        more.className = 'day-dot dot-more';
        more.textContent = `+${visibleOrders.length - 2} más`;
        el.appendChild(more);
      }
    }

    el.addEventListener('click', () => {
      if (el.classList.contains('filtered-out')) return;
      showDayDetail(d, y, m, allDayOrders);
    });
    grid.appendChild(el);
  }

  // Días del mes siguiente (relleno)
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
  document.getElementById('detail-title').textContent =
    `${d} de ${MONTHS[m]} ${y}`;

  const wrap = document.getElementById('detail-wrap');

  if (!dayOrders.length) {
    wrap.innerHTML = '<div class="no-data" style="padding:1rem">No hay órdenes este día</div>';
    return;
  }

  wrap.innerHTML = dayOrders.map(o => {
    const late = isLate(o);
    const cls  = late ? 'late' : 'ok';
    return `
      <div class="detail-order ${cls}">
        <div class="detail-order-num ${cls}">${escHtml(o.num)}</div>
        <div class="detail-order-est ${cls}">${escHtml(o.estado)}</div>
        ${late ? '<div class="detail-order-tag"><i class="ti ti-clock"></i> Atrasada</div>' : ''}
      </div>`;
  }).join('');
}

// ── Panel de alertas ─────────────────────────────────────
function renderAlerts() {
  const lateOrders = orders.filter(isLate);

  // Agrupar por día
  const byDay = {};
  lateOrders.forEach(o => {
    const key = formatDate(o.fecha);
    if (!byDay[key]) byDay[key] = [];
    byDay[key].push(o);
  });

  const wrap = document.getElementById('alerts-wrap');

  if (!lateOrders.length) {
    wrap.innerHTML = `
      <div class="alert-empty">
        <i class="ti ti-check" style="font-size:20px;display:block;margin-bottom:4px"></i>
        Sin atrasos detectados
      </div>`;
    return;
  }

  wrap.innerHTML = Object.entries(byDay)
    .sort((a, b) => new Date(a[0]) - new Date(b[0]))
    .map(([day, ords]) => `
      <div class="alert-item">
        <div class="alert-item-top">
          <i class="ti ti-alert-circle" style="font-size:14px;color:#A32D2D"></i>
          <span class="alert-item-label">${escHtml(day)}</span>
          <span class="badge badge-red" style="margin-left:auto">${ords.length} ord.</span>
        </div>
        <div class="alert-item-sub">
          ${ords.slice(0, 3).map(o => escHtml(o.num)).join(', ')}
          ${ords.length > 3 ? ` +${ords.length - 3} más` : ''}
        </div>
      </div>`)
    .join('');
}

// ── Búsqueda ─────────────────────────────────────────────
let searchQuery = '';

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

// ── Tabla de órdenes ─────────────────────────────────────
function highlight(text, query) {
  if (!query) return escHtml(text);
  const escaped = escHtml(text);
  const idx = escaped.toLowerCase().indexOf(query.toLowerCase());
  if (idx === -1) return escaped;
  return (
    escaped.slice(0, idx) +
    `<mark class="hl">${escaped.slice(idx, idx + query.length)}</mark>` +
    escaped.slice(idx + query.length)
  );
}

function renderTable(filtered) {
  // Aplicar filtros de leyenda además del filtro externo
  let fo = filtered || orders;
  fo = fo.filter(o => {
    if (isLate(o)) return activeFilters.late;
    return activeFilters.ok;
  });

  // Aplicar búsqueda de texto
  if (searchQuery) {
    fo = fo.filter(o => {
      const num    = String(o.num    || '').toLowerCase();
      const estado = String(o.estado || '').toLowerCase();
      const fecha  = formatDate(o.fecha).toLowerCase();
      return num.includes(searchQuery) || estado.includes(searchQuery) || fecha.includes(searchQuery);
    });
  }

  const total = (filtered || orders).filter(o => isLate(o) ? activeFilters.late : activeFilters.ok).length;
  const label = searchQuery
    ? `${fo.length} resultado${fo.length !== 1 ? 's' : ''} de ${total}`
    : `Todas las órdenes (${fo.length})`;
  document.getElementById('table-title').textContent = label;

  if (!fo.length) {
    document.getElementById('table-wrap').innerHTML = searchQuery
      ? `<div class="no-data"><i class="ti ti-search-off no-data-icon"></i>Sin resultados para "<strong>${escHtml(searchQuery)}</strong>"</div>`
      : '<div class="no-data"><i class="ti ti-table-off no-data-icon"></i>Carga un Excel para ver las órdenes</div>';
    return;
  }

  const rows = fo.slice(0, 100).map(o => {
    const late = isLate(o);
    const est  = String(o.estado || '').toLowerCase();
    const isDone = DONE_KEYWORDS.some(k => est.includes(k));
    const pillCls = late ? 'pill-late' : (isDone ? 'pill-ok' : 'pill-pend');
    const rowCls  = late ? 'late-row' : 'ok-row';
    return `
      <tr class="${rowCls}">
        <td>${highlight(String(o.num), searchQuery)}</td>
        <td><span class="status-pill ${pillCls}">${highlight(o.estado, searchQuery)}</span></td>
        <td>${highlight(formatDate(o.fecha), searchQuery)}</td>
        <td>${late
          ? '<span class="status-pill pill-late">Atrasada</span>'
          : '<span class="status-pill pill-ok">Al día</span>'}</td>
      </tr>`;
  }).join('');

  document.getElementById('table-wrap').innerHTML = `
    <table class="orders-table">
      <thead>
        <tr>
          <th style="width:30%">N° orden</th>
          <th style="width:25%">Estado</th>
          <th style="width:25%">Fecha</th>
          <th style="width:20%">Alerta</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
    ${fo.length > 100 ? `<div class="table-hint">Mostrando 100 de ${fo.length} órdenes</div>` : ''}`;
}

// ── Resumen en header ────────────────────────────────────
function renderSummary() {
  const late = orders.filter(isLate).length;
  const ok   = orders.length - late;
  document.getElementById('summary-badges').innerHTML = `
    <span class="badge badge-green"><i class="ti ti-check"></i> ${ok} al día</span>
    ${late ? `<span class="badge badge-red"><i class="ti ti-alert-circle"></i> ${late} atrasadas</span>` : ''}`;
}

// ── Procesar órdenes y re-renderizar todo ────────────────
function processOrders(data) {
  orders = normalizeOrders(data);
  renderCalendar();
  renderAlerts();
  renderTable();
  renderSummary();
}

// ── Cambiar mes ──────────────────────────────────────────
function changeMonth(dir) {
  currentDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + dir, 1);
  renderCalendar();
}

// ── Escape HTML básico ───────────────────────────────────
function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Leer archivo Excel / CSV ─────────────────────────────
document.getElementById('file-inp').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;

  document.getElementById('file-status').textContent = `Cargando: ${file.name}…`;

  const reader = new FileReader();
  reader.onload = function (ev) {
    try {
      const wb   = XLSX.read(ev.target.result, { type: 'array', cellDates: false });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { raw: true, defval: '' });

      processOrders(data);
      document.getElementById('file-status').textContent =
        `✓ ${file.name} — ${data.length} filas cargadas`;
    } catch (err) {
      document.getElementById('file-status').textContent =
        'Error al leer el archivo. Verifica que sea un Excel válido.';
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
  // Limpiar input para permitir volver a subir el mismo archivo
  this.value = '';
});

// ── Drag & Drop ──────────────────────────────────────────
const dropArea = document.getElementById('drop-area');

dropArea.addEventListener('dragover', e => {
  e.preventDefault();
  dropArea.classList.add('dragover');
});

dropArea.addEventListener('dragleave', () => {
  dropArea.classList.remove('dragover');
});

dropArea.addEventListener('drop', e => {
  e.preventDefault();
  dropArea.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (!file) return;

  // Simular selección
  const dt  = new DataTransfer();
  dt.items.add(file);
  document.getElementById('file-inp').files = dt.files;
  document.getElementById('file-inp').dispatchEvent(new Event('change'));
});

// ── Datos de ejemplo ─────────────────────────────────────
function loadSample() {
  const y = today.getFullYear();
  const m = today.getMonth();

  const fmt = dt =>
    `${String(dt.getDate()).padStart(2, '0')}/${String(dt.getMonth() + 1).padStart(2, '0')}/${dt.getFullYear()}`;

  const d = offset => new Date(y, m, today.getDate() + offset);

  const sample = [
    { 'numero de orden': 'ORD-001', 'estado de orden': 'Pendiente',   orden: fmt(d(-12)) },
    { 'numero de orden': 'ORD-002', 'estado de orden': 'En proceso',  orden: fmt(d(-8))  },
    { 'numero de orden': 'ORD-003', 'estado de orden': 'Completado',  orden: fmt(d(-6))  },
    { 'numero de orden': 'ORD-004', 'estado de orden': 'Pendiente',   orden: fmt(d(-4))  },
    { 'numero de orden': 'ORD-005', 'estado de orden': 'En revisión', orden: fmt(d(-2))  },
    { 'numero de orden': 'ORD-006', 'estado de orden': 'Entregado',   orden: fmt(d(-1))  },
    { 'numero de orden': 'ORD-007', 'estado de orden': 'Pendiente',   orden: fmt(d(0))   },
    { 'numero de orden': 'ORD-008', 'estado de orden': 'En proceso',  orden: fmt(d(2))   },
    { 'numero de orden': 'ORD-009', 'estado de orden': 'Pendiente',   orden: fmt(d(4))   },
    { 'numero de orden': 'ORD-010', 'estado de orden': 'En proceso',  orden: fmt(d(5))   },
    { 'numero de orden': 'ORD-011', 'estado de orden': 'Completado',  orden: fmt(d(-3))  },
    { 'numero de orden': 'ORD-012', 'estado de orden': 'Pendiente',   orden: fmt(d(-9))  },
    { 'numero de orden': 'ORD-013', 'estado de orden': 'En revisión', orden: fmt(d(-5))  },
    { 'numero de orden': 'ORD-014', 'estado de orden': 'Pendiente',   orden: fmt(d(7))   },
    { 'numero de orden': 'ORD-015', 'estado de orden': 'En proceso',  orden: fmt(d(10))  },
  ];

  processOrders(sample);
  document.getElementById('file-status').textContent =
    '✓ Datos de ejemplo cargados (15 órdenes)';
}

// ── Inicializar ──────────────────────────────────────────
renderCalendar();
