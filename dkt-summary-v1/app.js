const META_COLUMNS = [
  { id: 'id_material', label: 'SKU', sticky: 'sticky-1' },
  { id: 'descricao', label: 'Description', sticky: 'sticky-2' },
  { id: 'fornecedor', label: 'Supplier', sticky: 'sticky-3' },
  { id: 'categoria', label: 'Category', className: 'meta-col' },
  { id: 'abc', label: 'ABC', className: 'meta-col' },
  { id: 'custo_unit', label: 'Unit cost', numeric: true, className: 'meta-col' },
  { id: 'lote', label: 'Lot', numeric: true, className: 'meta-col' },
  { id: 'flag_validate', label: 'Validate', className: 'meta-col' },
  { id: 'total', label: 'Total', numeric: true, className: 'meta-col' },
];

const MONTH_FIELDS = Array.from({ length: 13 }, (_, index) => `m${String(index).padStart(2, '0')}`);
const RESIZABLE_COLUMN_IDS = [...META_COLUMNS.map((column) => column.id), ...MONTH_FIELDS];
const DEFAULT_COLUMN_WIDTHS = {
  id_material: 118,
  descricao: 430,
  fornecedor: 200,
  categoria: 140,
  abc: 90,
  custo_unit: 120,
  lote: 96,
  flag_validate: 96,
  total: 120,
  ...Object.fromEntries(MONTH_FIELDS.map((fieldId) => [fieldId, 108])),
};
const MIN_COLUMN_WIDTHS = {
  id_material: 90,
  descricao: 220,
  fornecedor: 150,
  categoria: 120,
  abc: 72,
  custo_unit: 100,
  lote: 80,
  flag_validate: 90,
  total: 100,
  ...Object.fromEntries(MONTH_FIELDS.map((fieldId) => [fieldId, 84])),
};
const MAX_COLUMN_WIDTH = 960;
const REQUESTED_COLUMNS = [
  ...META_COLUMNS.map((column) => ({ name: column.id, optional: true })),
  ...MONTH_FIELDS.map((fieldId) => ({ name: fieldId, optional: true })),
];
const MODE_CONFIG = {
  emissao: {
    kicker: 'DKT inventory equation · emissão',
    title: 'Monthly order emission review',
    subtitle: 'Planner-facing presentation sourced directly from SaidaAnalise_PedidosEmissao.',
  },
  recebimento: {
    kicker: 'DKT inventory equation · recebimento',
    title: 'Monthly order arrival review',
    subtitle: 'Planner-facing presentation sourced directly from SaidaAnalise_PedidosRecebimento.',
  },
};

const params = new URLSearchParams(window.location.search);
const mode = params.get('mode') === 'recebimento' ? 'recebimento' : 'emissao';
const modeConfig = MODE_CONFIG[mode];
const COLUMN_WIDTH_STORAGE_KEY = `dkt-summary-widths:${mode}`;

document.body.classList.add(`mode-${mode}`);
document.getElementById('mode-kicker').textContent = modeConfig.kicker;
document.getElementById('page-title').textContent = modeConfig.title;
document.getElementById('page-subtitle').textContent = modeConfig.subtitle;

const statusEl = document.getElementById('status');
const matrixShellEl = document.getElementById('matrix-shell');
const summaryTableEl = document.getElementById('summary-table');
const summaryPillEl = document.getElementById('summary-pill');

let latestMappings = null;
let columnWidths = loadColumnWidths();
let resizeState = null;

const numberFormatter = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
const decimalFormatter = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function setStatus(message) {
  statusEl.textContent = message;
}

function formatCellValue(columnId, value) {
  if (value === null || value === undefined || value === '') {
    return '<span class="muted">—</span>';
  }
  if (columnId === 'flag_validate') {
    const isOk = Number(value) === 1;
    return `<span class="flag-cell ${isOk ? 'flag-ok' : 'flag-warn'}">${isOk ? '1' : '0'}</span>`;
  }
  if (typeof value === 'number') {
    return Number.isInteger(value) ? numberFormatter.format(value) : decimalFormatter.format(value);
  }
  return escapeHtml(String(value));
}

function escapeHtml(value) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function getDefaultWidth(columnId) {
  return DEFAULT_COLUMN_WIDTHS[columnId] ?? 108;
}

function getMinWidth(columnId) {
  return MIN_COLUMN_WIDTHS[columnId] ?? 84;
}

function clampColumnWidth(columnId, width) {
  const numericWidth = Number(width);
  if (!Number.isFinite(numericWidth)) {
    return getDefaultWidth(columnId);
  }
  return Math.min(MAX_COLUMN_WIDTH, Math.max(getMinWidth(columnId), Math.round(numericWidth)));
}

function getColumnWidth(columnId) {
  return clampColumnWidth(columnId, columnWidths[columnId] ?? getDefaultWidth(columnId));
}

function loadColumnWidths() {
  try {
    const raw = window.localStorage.getItem(COLUMN_WIDTH_STORAGE_KEY);
    if (!raw) {
      return {};
    }
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') {
      return {};
    }
    return Object.fromEntries(
      Object.entries(parsed)
        .filter(([columnId]) => RESIZABLE_COLUMN_IDS.includes(columnId))
        .map(([columnId, width]) => [columnId, clampColumnWidth(columnId, width)]),
    );
  } catch {
    return {};
  }
}

function persistColumnWidths() {
  try {
    window.localStorage.setItem(COLUMN_WIDTH_STORAGE_KEY, JSON.stringify(columnWidths));
  } catch {
    // Ignore storage failures.
  }
}

function setColumnWidth(columnId, width, { persist = false } = {}) {
  columnWidths = {
    ...columnWidths,
    [columnId]: clampColumnWidth(columnId, width),
  };
  applyColumnWidths();
  if (persist) {
    persistColumnWidths();
  }
}

function resetColumnWidth(columnId) {
  const nextWidths = { ...columnWidths };
  delete nextWidths[columnId];
  columnWidths = nextWidths;
  applyColumnWidths();
  persistColumnWidths();
}

function applyColumnWidths() {
  const skuWidth = getColumnWidth('id_material');
  const descriptionWidth = getColumnWidth('descricao');

  summaryTableEl.style.setProperty('--sticky-left-1', '0px');
  summaryTableEl.style.setProperty('--sticky-left-2', `${skuWidth}px`);
  summaryTableEl.style.setProperty('--sticky-left-3', `${skuWidth + descriptionWidth}px`);

  for (const columnId of RESIZABLE_COLUMN_IDS) {
    const width = `${getColumnWidth(columnId)}px`;
    summaryTableEl.querySelectorAll(`[data-column-id="${columnId}"]`).forEach((cell) => {
      cell.style.width = width;
      cell.style.minWidth = width;
      cell.style.maxWidth = width;
    });
  }

  summaryTableEl.querySelectorAll('.sticky-3').forEach((cell) => {
    cell.style.boxShadow = `10px 0 16px rgba(18, 32, 51, 0.08)`;
  });
}

function buildHeader(labelsRow) {
  const headerRow = document.createElement('tr');

  for (const column of META_COLUMNS) {
    const th = document.createElement('th');
    th.className = [column.sticky, column.numeric ? 'numeric' : '', column.className || ''].filter(Boolean).join(' ');
    th.dataset.columnId = column.id;

    const labelEl = document.createElement('span');
    labelEl.className = 'header-label';
    labelEl.textContent = column.label;
    th.appendChild(labelEl);
    th.appendChild(buildResizeHandle(column.id, column.label));
    headerRow.appendChild(th);
  }

  for (const fieldId of MONTH_FIELDS) {
    const th = document.createElement('th');
    th.className = 'numeric month-col';
    th.dataset.columnId = fieldId;

    const headerLabelEl = document.createElement('span');
    headerLabelEl.className = 'header-label';

    const monthCodeEl = document.createElement('span');
    monthCodeEl.className = 'month-code';
    monthCodeEl.textContent = fieldId.toUpperCase();

    const monthLabelEl = document.createElement('span');
    monthLabelEl.className = 'month-label';
    monthLabelEl.textContent = labelsRow?.[fieldId] ? String(labelsRow[fieldId]) : fieldId.toUpperCase();

    headerLabelEl.append(monthCodeEl, monthLabelEl);
    th.appendChild(headerLabelEl);
    th.appendChild(buildResizeHandle(fieldId, fieldId.toUpperCase()));
    headerRow.appendChild(th);
  }

  const thead = document.createElement('thead');
  thead.appendChild(headerRow);
  return thead;
}

function buildResizeHandle(columnId, label) {
  const handle = document.createElement('button');
  handle.type = 'button';
  handle.className = 'resize-handle';
  handle.dataset.columnId = columnId;
  handle.setAttribute('aria-label', `Resize ${label} column`);
  handle.addEventListener('mousedown', startColumnResize);
  handle.addEventListener('dblclick', (event) => {
    event.preventDefault();
    event.stopPropagation();
    resetColumnWidth(columnId);
  });
  return handle;
}

function buildBody(itemRows, totalRow) {
  const tbody = document.createElement('tbody');

  for (const row of itemRows) {
    tbody.appendChild(buildDataRow(row, false));
  }

  if (totalRow) {
    tbody.appendChild(buildDataRow(totalRow, true));
  }

  return tbody;
}

function buildDataRow(row, isTotal) {
  const tr = document.createElement('tr');
  if (isTotal) {
    tr.classList.add('total-row');
  }

  for (const column of META_COLUMNS) {
    const td = document.createElement('td');
    td.className = [column.sticky, column.numeric ? 'numeric' : '', column.className || ''].filter(Boolean).join(' ');
    td.dataset.columnId = column.id;
    td.innerHTML = formatCellValue(column.id, row[column.id]);
    tr.appendChild(td);
  }

  for (const fieldId of MONTH_FIELDS) {
    const td = document.createElement('td');
    td.className = 'numeric month-col';
    td.dataset.columnId = fieldId;
    td.innerHTML = formatCellValue(fieldId, row[fieldId]);
    tr.appendChild(td);
  }

  return tr;
}

function normalizeRecords(records) {
  return (records || []).map((record) => record?.fields ?? record ?? {});
}

function mapRecords(records, mappings = latestMappings) {
  if (!window.grist?.mapColumnNames) {
    return normalizeRecords(records);
  }
  const mapped = window.grist.mapColumnNames(records, { mappings, columns: REQUESTED_COLUMNS });
  return mapped ? normalizeRecords(mapped) : normalizeRecords(records);
}

function renderTable(records) {
  const normalizedRecords = normalizeRecords(records);
  const labelRow = normalizedRecords.find((row) => !row.id_material);
  const totalRow = normalizedRecords.find((row) => String(row.id_material || '').trim().toUpperCase() === 'TOTAL GERAL');
  const itemRows = normalizedRecords.filter((row) => row !== labelRow && row !== totalRow);

  if (!itemRows.length) {
    matrixShellEl.hidden = true;
    summaryPillEl.hidden = true;
    setStatus('No summary rows were returned by the selected Grist backing table.');
    return;
  }

  summaryTableEl.replaceChildren(buildHeader(labelRow), buildBody(itemRows, totalRow));
  applyColumnWidths();
  matrixShellEl.hidden = false;

  const firstRowKeys = Object.keys(normalizedRecords[0] || {});
  const hasRecognizedFields = itemRows.some((row) => row.id_material || row.descricao || row.total !== undefined);
  if (hasRecognizedFields) {
    setStatus(`Loaded ${itemRows.length} SKU rows directly from the selected Grist backing table. Drag any header edge to resize columns. First row keys: ${firstRowKeys.join(', ') || '(none)'}.`);
  } else {
    setStatus(`Loaded ${itemRows.length} rows, but the widget did not recognize the record shape. First row: ${JSON.stringify(normalizedRecords[0] || {}).slice(0, 240)}.`);
  }

  if (totalRow && totalRow.total !== undefined && totalRow.total !== null && totalRow.total !== '') {
    summaryPillEl.hidden = false;
    summaryPillEl.textContent = `Workbook total: ${formatPillValue(totalRow.total)}`;
  } else {
    summaryPillEl.hidden = true;
  }
}

function formatPillValue(value) {
  if (typeof value === 'number') {
    return Number.isInteger(value) ? numberFormatter.format(value) : decimalFormatter.format(value);
  }
  return String(value);
}

function startColumnResize(event) {
  event.preventDefault();
  event.stopPropagation();

  const columnId = event.currentTarget?.dataset?.columnId;
  if (!columnId) {
    return;
  }

  resizeState = {
    columnId,
    startX: event.clientX,
    startWidth: getColumnWidth(columnId),
  };
  document.body.classList.add('is-resizing-columns');
}

function handleColumnResizeMove(event) {
  if (!resizeState) {
    return;
  }

  const deltaX = event.clientX - resizeState.startX;
  setColumnWidth(resizeState.columnId, resizeState.startWidth + deltaX);
}

function stopColumnResize() {
  if (!resizeState) {
    return;
  }

  persistColumnWidths();
  resizeState = null;
  document.body.classList.remove('is-resizing-columns');
}

document.addEventListener('mousemove', handleColumnResizeMove);
document.addEventListener('mouseup', stopColumnResize);
document.addEventListener('mouseleave', stopColumnResize);
window.addEventListener('blur', stopColumnResize);

async function fetchSelectedBackingTableRows() {
  return await window.grist.fetchSelectedTable({
    format: 'rows',
  });
}

async function refreshFromSelectedBackingTable() {
  try {
    setStatus('Loading summary rows from the selected Grist backing table...');
    const rows = await fetchSelectedBackingTableRows();
    renderTable(mapRecords(rows));
  } catch (error) {
    matrixShellEl.hidden = true;
    summaryPillEl.hidden = true;
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Failed to load the selected Grist backing table: ${message}`);
    console.error(error);
  }
}

if (window.grist) {
  window.grist.ready({ requiredAccess: 'read table', columns: REQUESTED_COLUMNS });

  window.grist.onRecords((records, mappings) => {
    latestMappings = mappings || latestMappings;
    renderTable(mapRecords(records, latestMappings));
  }, { format: 'rows' });

  let lastTableId = null;
  window.grist.on('message', (message) => {
    const nextTableId = message?.tableId || null;
    const tableChanged = Boolean(nextTableId && nextTableId !== lastTableId);
    if (nextTableId) {
      lastTableId = nextTableId;
    }
    if (message?.mappingsChange) {
      latestMappings = null;
    }

    if (tableChanged || message?.dataChange || message?.mappingsChange) {
      void refreshFromSelectedBackingTable();
    }
  });

  void refreshFromSelectedBackingTable();
} else {
  setStatus('This widget must be opened inside Grist.');
}
