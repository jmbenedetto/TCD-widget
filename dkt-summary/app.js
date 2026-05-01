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

document.body.classList.add(`mode-${mode}`);
document.getElementById('mode-kicker').textContent = modeConfig.kicker;
document.getElementById('page-title').textContent = modeConfig.title;
document.getElementById('page-subtitle').textContent = modeConfig.subtitle;

const statusEl = document.getElementById('status');
const matrixShellEl = document.getElementById('matrix-shell');
const summaryTableEl = document.getElementById('summary-table');
const summaryPillEl = document.getElementById('summary-pill');

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

function buildHeader(labelsRow) {
  const headerRow = document.createElement('tr');

  for (const column of META_COLUMNS) {
    const th = document.createElement('th');
    th.className = [column.sticky, column.numeric ? 'numeric' : '', column.className || ''].filter(Boolean).join(' ');
    th.textContent = column.label;
    headerRow.appendChild(th);
  }

  for (const fieldId of MONTH_FIELDS) {
    const th = document.createElement('th');
    th.className = 'numeric month-col';
    const monthCode = fieldId.toUpperCase();
    const monthLabel = labelsRow?.[fieldId] ? escapeHtml(String(labelsRow[fieldId])) : monthCode;
    th.innerHTML = `${monthCode}<span class="month-label">${monthLabel}</span>`;
    headerRow.appendChild(th);
  }

  const thead = document.createElement('thead');
  thead.appendChild(headerRow);
  return thead;
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
    td.innerHTML = formatCellValue(column.id, row[column.id]);
    tr.appendChild(td);
  }

  for (const fieldId of MONTH_FIELDS) {
    const td = document.createElement('td');
    td.className = 'numeric month-col';
    td.innerHTML = formatCellValue(fieldId, row[fieldId]);
    tr.appendChild(td);
  }

  return tr;
}

function normalizeRecords(records) {
  return (records || []).map((record) => record?.fields ?? record ?? {});
}

function tableDataToRows(table) {
  const columnIds = Array.isArray(table?.column_metadata)
    ? table.column_metadata.map((column) => column?.id).filter(Boolean)
    : [];

  if (!columnIds.length) {
    return [];
  }

  const dataByColumn = {};
  if (Array.isArray(table?.table_data)) {
    columnIds.forEach((columnId, index) => {
      dataByColumn[columnId] = Array.isArray(table.table_data[index]) ? table.table_data[index] : [];
    });
  } else if (table?.table_data && typeof table.table_data === 'object') {
    columnIds.forEach((columnId) => {
      dataByColumn[columnId] = Array.isArray(table.table_data[columnId]) ? table.table_data[columnId] : [];
    });
  }

  const rowCount = Math.max(0, ...Object.values(dataByColumn).map((values) => values.length));
  return Array.from({ length: rowCount }, (_, rowIndex) => {
    const row = {};
    columnIds.forEach((columnId) => {
      row[columnId] = dataByColumn[columnId]?.[rowIndex];
    });
    return row;
  });
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
  matrixShellEl.hidden = false;

  const firstRowKeys = Object.keys(normalizedRecords[0] || {});
  const hasRecognizedFields = itemRows.some((row) => row.id_material || row.descricao || row.total !== undefined);
  if (hasRecognizedFields) {
    setStatus(`Loaded ${itemRows.length} SKU rows directly from the selected Grist backing table. First row keys: ${firstRowKeys.join(', ') || '(none)'}.`);
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

async function fetchSelectedBackingTableRows() {
  const tableId = await window.grist.getSelectedTableId();
  if (!tableId) {
    throw new Error('No Grist backing table is currently selected.');
  }

  const table = await window.grist.docApi.fetchTable(tableId);
  return tableDataToRows(table);
}

async function refreshFromSelectedBackingTable() {
  try {
    setStatus('Loading summary rows from the selected Grist backing table...');
    const rows = await fetchSelectedBackingTableRows();
    renderTable(rows);
  } catch (error) {
    matrixShellEl.hidden = true;
    summaryPillEl.hidden = true;
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Failed to load the selected Grist backing table: ${message}`);
    console.error(error);
  }
}

if (window.grist) {
  window.grist.ready({ requiredAccess: 'read table' });

  let lastTableId = null;
  window.grist.on('message', (message) => {
    const nextTableId = message?.tableId || null;
    const tableChanged = Boolean(nextTableId && nextTableId !== lastTableId);
    if (nextTableId) {
      lastTableId = nextTableId;
    }

    if (tableChanged || message?.dataChange) {
      void refreshFromSelectedBackingTable();
    }
  });

  void refreshFromSelectedBackingTable();
} else {
  setStatus('This widget must be opened inside Grist.');
}
