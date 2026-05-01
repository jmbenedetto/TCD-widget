const SOURCE_TABLE_ID = 'SaidaDados_ProjecaoEstoque';
const MONTH_FIELDS = Array.from({ length: 13 }, (_, index) => `M${String(index).padStart(2, '0')}`);
const REQUESTED_COLUMNS = [
  { name: 'id_material', optional: true },
  { name: 'periodo', optional: true },
  { name: 'ano_mes', optional: true },
  { name: 'descricao', optional: true },
  { name: 'fornecedor', optional: true },
  { name: 'categoria', optional: true },
  { name: 'abc_index', optional: true },
  { name: 'custo', optional: true },
  { name: 'lote', optional: true },
  { name: 'flag_validate', optional: true },
  { name: 'proposta_pedido_qtd', optional: true },
  { name: 'lead_time_meses', optional: true },
  { name: 'ultimo_periodo_possivel_pedido', optional: true },
];
const SOURCE_FIELD_MAP = {
  sku: ['id_material'],
  period: ['periodo'],
  monthLabel: ['ano_mes'],
  description: ['descricao'],
  supplier: ['fornecedor'],
  category: ['categoria'],
  abc: ['abc_index', 'abc'],
  unitCost: ['custo', 'custo_unit'],
  lot: ['lote'],
  validate: ['flag_validate'],
  proposalQty: ['proposta_pedido_qtd'],
  leadTime: ['lead_time_meses'],
  latestOrderPeriod: ['ultimo_periodo_possivel_pedido'],
};
const MODE_CONFIG = {
  emissao: {
    kicker: 'Planejamento de suprimentos',
    title: 'Pedidos propostos',
    subtitle: 'Revise as recomendações por período de emissão.',
    toggleLabel: 'Emissão',
    periodTransform: (row) => normalizePeriod(readMappedField(row, 'period')),
    fieldsNote: 'id_material, periodo, ano_mes, descricao, fornecedor, categoria, abc_index, custo, lote, flag_validate, proposta_pedido_qtd.',
  },
  recebimento: {
    kicker: 'Planejamento de suprimentos',
    title: 'Pedidos propostos',
    subtitle: 'Revise as recomendações por período de recebimento.',
    toggleLabel: 'Recebimento',
    periodTransform: (row) => shiftPeriod(normalizePeriod(readMappedField(row, 'period')), toInteger(readMappedField(row, 'leadTime'))),
    fieldsNote: 'id_material, periodo, ano_mes, descricao, fornecedor, categoria, abc_index, custo, lote, flag_validate, proposta_pedido_qtd, lead_time_meses.',
  },
};
const params = new URLSearchParams(window.location.search);
const debugMode = params.get('debug') === '1';
const initialMode = params.get('mode') === 'recebimento' ? 'recebimento' : 'emissao';
let currentMode = initialMode;
const statusEl = document.getElementById('status');
const notesEl = document.getElementById('notes-card');
const tableShellEl = document.getElementById('table-shell');
const sourcePillEl = document.getElementById('source-pill');
const summaryPillEl = document.getElementById('summary-pill');
const titleEl = document.getElementById('page-title');
const subtitleEl = document.getElementById('page-subtitle');
const kickerEl = document.getElementById('mode-kicker');
const modeToggleButtons = Array.from(document.querySelectorAll('[data-mode-toggle]'));
const numberFormatter = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });
const decimalFormatter = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
let table = null;
let currentTableId = null;
let latestRows = [];
let latestSourceLabel = null;

setElementVisible(notesEl, debugMode);
setElementVisible(sourcePillEl, debugMode);
setElementVisible(statusEl, debugMode);
applyModeUi();

function getModeConfig() {
  return MODE_CONFIG[currentMode];
}

function applyModeUi() {
  const modeConfig = getModeConfig();
  kickerEl.textContent = modeConfig.kicker;
  titleEl.textContent = modeConfig.title;
  subtitleEl.textContent = modeConfig.subtitle;
  for (const button of modeToggleButtons) {
    const isActive = button.dataset.modeToggle === currentMode;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  }
}

function setElementVisible(element, visible) {
  if (!element) {
    return;
  }
  element.hidden = !visible;
  element.classList.toggle('is-hidden', !visible);
}

function setStatus(message, { visible = debugMode } = {}) {
  statusEl.textContent = message;
  setElementVisible(statusEl, visible);
}

function getDocApi() {
  return window.grist?.docApi || window.grist?.raw?.docApi || null;
}

function normalizeRow(record) {
  return record?.fields ?? record ?? {};
}

function rowsFromColumnarTable(tableData) {
  const columns = Object.keys(tableData || {});
  if (!columns.length) {
    return [];
  }
  const rowCount = Array.isArray(tableData[columns[0]]) ? tableData[columns[0]].length : 0;
  return Array.from({ length: rowCount }, (_, index) => {
    const row = {};
    for (const column of columns) {
      row[column] = tableData[column]?.[index];
    }
    return row;
  });
}

function rowsFromTablePayload(payload) {
  if (!payload) {
    return [];
  }
  if (Array.isArray(payload)) {
    return payload.map(normalizeRow);
  }
  if (Array.isArray(payload.records)) {
    return payload.records.map(normalizeRow);
  }
  if (payload.tableData && typeof payload.tableData === 'object') {
    return rowsFromColumnarTable(payload.tableData);
  }
  if (typeof payload === 'object') {
    const firstValue = Object.values(payload)[0];
    if (Array.isArray(firstValue)) {
      return rowsFromColumnarTable(payload);
    }
  }
  return [];
}

function readMappedField(row, logicalField) {
  const candidates = SOURCE_FIELD_MAP[logicalField] || [];
  for (const field of candidates) {
    if (row[field] !== undefined && row[field] !== null && row[field] !== '') {
      return row[field];
    }
  }
  return null;
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  const raw = String(value).trim();
  const normalized = raw.includes(',')
    ? (raw.includes('.') ? raw.replaceAll('.', '').replace(',', '.') : raw.replace(',', '.'))
    : raw;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function toInteger(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  const numeric = typeof value === 'number' ? value : Number(String(value).trim().replace(',', '.'));
  return Number.isFinite(numeric) ? Math.trunc(numeric) : 0;
}

function normalizePeriod(value) {
  if (!value) {
    return null;
  }
  const match = String(value).trim().toUpperCase().match(/^M(\d{1,2})$/);
  if (!match) {
    return null;
  }
  return `M${String(Number(match[1])).padStart(2, '0')}`;
}

function shiftPeriod(period, offset) {
  if (!period) {
    return null;
  }
  const baseIndex = Number(period.slice(1));
  if (!Number.isFinite(baseIndex)) {
    return null;
  }
  const targetIndex = baseIndex + (Number.isFinite(offset) ? offset : 0);
  if (targetIndex < 0 || targetIndex >= MONTH_FIELDS.length) {
    return null;
  }
  return `M${String(targetIndex).padStart(2, '0')}`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function formatNumeric(value) {
  if (value === null || value === undefined || value === '') {
    return '<span class="muted">—</span>';
  }
  const numeric = toNumber(value);
  if (Number.isInteger(numeric)) {
    return numberFormatter.format(numeric);
  }
  return decimalFormatter.format(numeric);
}

function formatFlag(value) {
  if (value === null || value === undefined || value === '') {
    return '<span class="muted">—</span>';
  }
  const normalized = toInteger(value);
  const kind = normalized === 1 ? 'ok' : 'warn';
  return `<span class="flag-pill ${kind}">${normalized === 1 ? '1' : '0'}</span>`;
}

function hasLatestOrderPeriod(row) {
  const value = readMappedField(row, 'latestOrderPeriod');
  return value !== null && value !== undefined && String(value).trim() !== '';
}

function buildDisplayRows(rows) {
  const modeConfig = getModeConfig();
  const groups = new Map();
  const monthLabels = Object.fromEntries(MONTH_FIELDS.map((fieldId) => [fieldId, fieldId]));
  let skippedRows = 0;
  let excludedProposalRows = 0;
  let excludedProposalUnits = 0;

  for (const rawRow of rows) {
    const row = normalizeRow(rawRow);
    const sku = readMappedField(row, 'sku');
    const period = modeConfig.periodTransform(row);
    if (!sku || !period || !MONTH_FIELDS.includes(period)) {
      skippedRows += 1;
      continue;
    }

    const rawProposalQty = toNumber(readMappedField(row, 'proposalQty'));
    const isActionableProposal = rawProposalQty <= 0 || hasLatestOrderPeriod(row);
    const proposalQty = isActionableProposal ? rawProposalQty : 0;
    if (!isActionableProposal) {
      excludedProposalRows += 1;
      excludedProposalUnits += rawProposalQty;
    }

    const monthLabel = readMappedField(row, 'monthLabel');
    if (monthLabel) {
      monthLabels[period] = String(monthLabel);
    }

    const rowKey = String(sku);
    if (!groups.has(rowKey)) {
      groups.set(rowKey, {
        _rowKey: rowKey,
        id_material: rowKey,
        descricao: readMappedField(row, 'description') || '',
        fornecedor: readMappedField(row, 'supplier') || '',
        categoria: readMappedField(row, 'category') || '',
        abc: readMappedField(row, 'abc') || '',
        custo_unit: toNumber(readMappedField(row, 'unitCost')) || 0,
        lote: toNumber(readMappedField(row, 'lot')) || 0,
        flag_validate: toInteger(readMappedField(row, 'validate')),
        latest_order_period: readMappedField(row, 'latestOrderPeriod') || '',
        total: 0,
        ...Object.fromEntries(MONTH_FIELDS.map((fieldId) => [fieldId, 0])),
      });
    }

    const grouped = groups.get(rowKey);
    grouped[period] += proposalQty;
    grouped.total += proposalQty;
    grouped.flag_validate = Math.max(grouped.flag_validate, toInteger(readMappedField(row, 'validate')));
    if (!grouped.descricao && readMappedField(row, 'description')) {
      grouped.descricao = readMappedField(row, 'description');
    }
    if (!grouped.fornecedor && readMappedField(row, 'supplier')) {
      grouped.fornecedor = readMappedField(row, 'supplier');
    }
    if (!grouped.categoria && readMappedField(row, 'category')) {
      grouped.categoria = readMappedField(row, 'category');
    }
    if (!grouped.abc && readMappedField(row, 'abc')) {
      grouped.abc = readMappedField(row, 'abc');
    }
    if (!grouped.latest_order_period && readMappedField(row, 'latestOrderPeriod')) {
      grouped.latest_order_period = readMappedField(row, 'latestOrderPeriod');
    }
  }

  const displayRows = Array.from(groups.values()).sort((left, right) => left.id_material.localeCompare(right.id_material));
  const totalRow = {
    _rowKey: '__total__',
    _rowType: 'total',
    id_material: 'TOTAL GERAL',
    descricao: `${displayRows.length} SKUs`,
    fornecedor: '',
    categoria: '',
    abc: '',
    custo_unit: '',
    lote: '',
    flag_validate: '',
    latest_order_period: '',
    total: 0,
    ...Object.fromEntries(MONTH_FIELDS.map((fieldId) => [fieldId, 0])),
  };

  for (const row of displayRows) {
    for (const fieldId of MONTH_FIELDS) {
      totalRow[fieldId] += toNumber(row[fieldId]);
    }
    totalRow.total += toNumber(row.total);
  }

  return {
    displayRows: [...displayRows, totalRow],
    skuCount: displayRows.length,
    skippedRows,
    excludedProposalRows,
    excludedProposalUnits,
    totalUnits: totalRow.total,
    monthLabels,
  };
}

function buildColumns(monthLabels) {
  const baseColumns = [
    {
      title: 'SKU',
      field: 'id_material',
      frozen: true,
      width: 128,
      resizable: true,
      cssClass: 'frozen-meta',
      formatter: (cell) => escapeHtml(cell.getValue() || '—'),
    },
    {
      title: 'Descrição',
      field: 'descricao',
      frozen: true,
      width: 340,
      minWidth: 220,
      resizable: true,
      cssClass: 'frozen-meta wrap-cell',
      formatter: (cell) => escapeHtml(cell.getValue() || '—'),
    },
    {
      title: 'Fornecedor',
      field: 'fornecedor',
      frozen: true,
      width: 200,
      minWidth: 160,
      resizable: true,
      cssClass: 'frozen-meta wrap-cell',
      formatter: (cell) => escapeHtml(cell.getValue() || '—'),
    },
    { title: 'Categoria', field: 'categoria', width: 170, minWidth: 140, resizable: true, formatter: (cell) => escapeHtml(cell.getValue() || '—') },
    { title: 'ABC', field: 'abc', width: 86, hozAlign: 'center', resizable: true, formatter: (cell) => escapeHtml(cell.getValue() || '—') },
    { title: 'Custo un.', field: 'custo_unit', width: 118, hozAlign: 'right', resizable: true, cssClass: 'numeric-cell', formatter: (cell) => formatNumeric(cell.getValue()) },
    { title: 'Lote', field: 'lote', width: 96, hozAlign: 'right', resizable: true, cssClass: 'numeric-cell', formatter: (cell) => formatNumeric(cell.getValue()) },
    { title: 'Validar', field: 'flag_validate', width: 96, hozAlign: 'center', resizable: true, formatter: (cell) => formatFlag(cell.getValue()) },
    { title: 'Últ. pedido', field: 'latest_order_period', width: 120, hozAlign: 'center', resizable: true, formatter: (cell) => escapeHtml(cell.getValue() || '—') },
    { title: 'Total', field: 'total', width: 120, hozAlign: 'right', resizable: true, cssClass: 'numeric-cell', formatter: (cell) => formatNumeric(cell.getValue()) },
  ];

  const monthColumns = MONTH_FIELDS.map((fieldId) => ({
    title: `${fieldId}<br><span class="muted">${escapeHtml(monthLabels[fieldId] || fieldId)}</span>`,
    field: fieldId,
    width: 108,
    minWidth: 90,
    resizable: true,
    hozAlign: 'right',
    cssClass: 'numeric-cell',
    headerSort: false,
    formatter: (cell) => formatNumeric(cell.getValue()),
  }));

  return [...baseColumns, ...monthColumns];
}

function ensureTable(columns) {
  if (!table) {
    table = new Tabulator('#summary-table', {
      data: [],
      columns,
      layout: 'fitDataFill',
      responsiveLayout: false,
      height: '68vh',
      movableColumns: false,
      resizableColumns: true,
      selectableRows: false,
      placeholder: 'Nenhuma linha encontrada na fonte de dados.',
      rowFormatter: (row) => {
        const element = row.getElement();
        if (row.getData()?._rowType === 'total') {
          element.classList.add('total-row');
        } else {
          element.classList.remove('total-row');
        }
      },
    });
    return;
  }
  table.setColumns(columns);
}

async function renderFromRows(rows, sourceLabel) {
  latestRows = rows;
  latestSourceLabel = sourceLabel;
  const modeConfig = getModeConfig();
  const { displayRows, skuCount, skippedRows, excludedProposalRows, excludedProposalUnits, totalUnits, monthLabels } = buildDisplayRows(rows);
  ensureTable(buildColumns(monthLabels));
  await table.replaceData(displayRows);
  setElementVisible(tableShellEl, true);
  setElementVisible(summaryPillEl, true);
  summaryPillEl.textContent = `${modeConfig.toggleLabel} · ${numberFormatter.format(skuCount)} SKUs · ${numberFormatter.format(totalUnits)} un.`;
  setStatus(
    `Fonte: ${currentTableId || SOURCE_TABLE_ID}. ${numberFormatter.format(rows.length)} linhas recebidas, ${numberFormatter.format(skuCount)} SKUs renderizados, ${numberFormatter.format(skippedRows)} linhas fora da janela M00-M12, ${numberFormatter.format(excludedProposalRows)} linhas de proposta excluídas por não terem último período possível de pedido, total excluído de ${numberFormatter.format(excludedProposalUnits)} un. Campos usados: ${modeConfig.fieldsNote}`,
    { visible: debugMode }
  );
}

async function fetchRowsFromDocApi(tableName) {
  const docApi = getDocApi();
  if (!docApi?.fetchTable || !tableName) {
    return [];
  }
  const payload = await docApi.fetchTable(tableName);
  return rowsFromTablePayload(payload);
}

async function fetchRowsFromSelectedTable() {
  if (!window.grist?.fetchSelectedTable) {
    return [];
  }
  const payload = await window.grist.fetchSelectedTable({ format: 'rows' });
  return rowsFromTablePayload(payload);
}

function persistModeInUrl() {
  const nextParams = new URLSearchParams(window.location.search);
  nextParams.set('mode', currentMode);
  const nextUrl = `${window.location.pathname}?${nextParams.toString()}${window.location.hash || ''}`;
  window.history.replaceState({}, '', nextUrl);
}

async function applyMode(nextMode) {
  if (!MODE_CONFIG[nextMode]) {
    return;
  }
  currentMode = nextMode;
  applyModeUi();
  persistModeInUrl();
  if (latestRows.length) {
    await renderFromRows(latestRows, latestSourceLabel || 'cached direct-source rows');
    return;
  }
  setStatus(`Carregando visão de ${MODE_CONFIG[nextMode].toggleLabel.toLowerCase()}...`, { visible: debugMode });
}

async function refreshRows(reason) {
  try {
    setStatus(`Carregando dados (${reason})...`, { visible: debugMode });

    if (currentTableId === SOURCE_TABLE_ID && latestRows.length) {
      await renderFromRows(latestRows, 'grist.onRecords');
      return;
    }

    const fromSourceTable = await fetchRowsFromDocApi(SOURCE_TABLE_ID);
    if (fromSourceTable.length) {
      currentTableId = SOURCE_TABLE_ID;
      await renderFromRows(fromSourceTable, 'grist.docApi.fetchTable(source table)');
      return;
    }

    const fromSelectedTable = await fetchRowsFromSelectedTable();
    if (fromSelectedTable.length) {
      await renderFromRows(fromSelectedTable, 'grist.fetchSelectedTable');
      return;
    }

    if (latestRows.length) {
      await renderFromRows(latestRows, 'cached onRecords payload');
      return;
    }

    setElementVisible(tableShellEl, false);
    setElementVisible(summaryPillEl, false);
    setStatus('Nenhuma linha foi retornada pela fonte de dados.', { visible: true });
  } catch (error) {
    setElementVisible(tableShellEl, false);
    setElementVisible(summaryPillEl, false);
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Não foi possível carregar os dados: ${message}`, { visible: true });
    console.error(error);
  }
}

for (const button of modeToggleButtons) {
  button.addEventListener('click', () => {
    void applyMode(button.dataset.modeToggle);
  });
}

if (window.grist) {
  window.grist.ready({ requiredAccess: 'read table', columns: REQUESTED_COLUMNS });

  window.grist.onRecords((records) => {
    latestRows = Array.isArray(records) ? records.map(normalizeRow) : [];
    void renderFromRows(latestRows, 'grist.onRecords');
  }, { format: 'rows' });

  window.grist.on('message', (message) => {
    if (message?.tableId) {
      currentTableId = message.tableId;
    }
    if (message?.tableId || message?.dataChange || message?.mappingsChange) {
      void refreshRows(message?.tableId ? 'table change' : 'data change');
    }
  });

  void refreshRows('initial load');
} else {
  setStatus('Este widget precisa ser aberto dentro do Grist.', { visible: true });
}
