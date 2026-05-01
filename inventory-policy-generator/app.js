const MONTH_BUCKETS = Array.from({ length: 13 }, (_, index) => `M${String(index).padStart(2, '0')}`);
const REQUESTED_COLUMNS = [
  { name: 'id_politica_estoque', optional: true },
  { name: 'policy_name', optional: true },
  { name: 'policy_type', optional: true },
  { name: 'mes_base', optional: true },
  { name: 'project_reference', optional: true },
  { name: 'generation_status', optional: true },
  { name: 'generated_rows_count', optional: true },
  { name: 'generated_at', optional: true },
  { name: 'generation_error', optional: true },
];

const el = {
  button: document.getElementById('generate-button'),
  statusText: document.getElementById('status-text'),
  statusBadge: document.getElementById('status-badge'),
  policyId: document.getElementById('policy-id'),
  policyName: document.getElementById('policy-name'),
  policyType: document.getElementById('policy-type'),
  policyMonth: document.getElementById('policy-month'),
  policyGenerationStatus: document.getElementById('policy-generation-status'),
  policyGeneratedRows: document.getElementById('policy-generated-rows'),
  previewProducts: document.getElementById('preview-products'),
  previewRows: document.getElementById('preview-rows'),
};

let selectedPolicy = null;
let generationRunning = false;

function getDocApi() {
  return window.grist?.docApi || window.grist?.raw?.docApi || null;
}

function setBadge(kind, label) {
  el.statusBadge.className = `badge badge-${kind}`;
  el.statusBadge.textContent = label;
}

function setStatus(message) {
  el.statusText.textContent = message;
}

function showPolicy(policy) {
  el.policyId.textContent = policy?.id_politica_estoque || '—';
  el.policyName.textContent = policy?.policy_name || '—';
  el.policyType.textContent = policy?.policy_type || '—';
  el.policyMonth.textContent = policy?.mes_base || '—';
  el.policyGenerationStatus.textContent = policy?.generation_status || '—';
  el.policyGeneratedRows.textContent = policy?.generated_rows_count ?? '—';
}

function updatePreview(productCount) {
  el.previewProducts.textContent = productCount ?? '—';
  el.previewRows.textContent = Number.isFinite(productCount) ? String(productCount * MONTH_BUCKETS.length) : '—';
}

function normalizeRow(record) {
  return record?.fields ?? record ?? {};
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

async function fetchTableRows(tableName) {
  const docApi = getDocApi();
  if (!docApi?.fetchTable) {
    throw new Error('Grist docApi.fetchTable is not available in this widget runtime.');
  }
  const payload = await docApi.fetchTable(tableName);
  return rowsFromTablePayload(payload);
}

function selectedRowId(policy) {
  return policy?.id ?? policy?.recordId ?? null;
}

function buildStartPatch(policy) {
  const patch = {};
  if ('generation_status' in policy) {
    patch.generation_status = 'generating';
  }
  if ('generation_error' in policy) {
    patch.generation_error = '';
  }
  return patch;
}

function buildSuccessPatch(policy, rowCount) {
  const patch = {};
  if ('generation_status' in policy) {
    patch.generation_status = 'generated';
  }
  if ('generated_rows_count' in policy) {
    patch.generated_rows_count = rowCount;
  }
  if ('generated_at' in policy) {
    patch.generated_at = new Date().toISOString();
  }
  if ('generation_error' in policy) {
    patch.generation_error = '';
  }
  return patch;
}

function buildErrorPatch(policy, message) {
  const patch = {};
  if ('generation_status' in policy) {
    patch.generation_status = 'error';
  }
  if ('generation_error' in policy) {
    patch.generation_error = message;
  }
  return patch;
}

function chunk(list, size) {
  const batches = [];
  for (let index = 0; index < list.length; index += size) {
    batches.push(list.slice(index, index + size));
  }
  return batches;
}

function buildAddActions(policy, products) {
  const policyRef = selectedRowId(policy);
  return products.flatMap((product) =>
    MONTH_BUCKETS.map((bucket) => [
      'AddRecord',
      'SaidaDados_Politica',
      null,
      {
        id_produto: product.id_produto,
        policy_ref: policyRef,
        bucket_planejamento: bucket,
      },
    ])
  );
}

function countExistingRows(rows, policy) {
  const policyRef = selectedRowId(policy);
  const policyId = String(policy.id_politica_estoque || '');
  return rows.filter((row) => {
    const samePolicyRef = row.policy_ref === policyRef;
    const samePolicyId = String(row.id_politica_estoque || '') === policyId;
    return samePolicyRef || samePolicyId;
  }).length;
}

async function applyActions(actions) {
  if (!actions.length) {
    return;
  }
  const docApi = getDocApi();
  if (!docApi?.applyUserActions) {
    throw new Error('Grist docApi.applyUserActions is not available in this widget runtime.');
  }
  await docApi.applyUserActions(actions);
}

async function updatePolicyRow(policy, patch) {
  if (!Object.keys(patch).length) {
    return;
  }
  const rowId = selectedRowId(policy);
  if (!rowId) {
    throw new Error('Selected policy row is missing its internal Grist record id.');
  }
  await applyActions([['UpdateRecord', 'Entrada_Politicas', rowId, patch]]);
}

async function generateDetailRows() {
  if (!selectedPolicy || generationRunning) {
    return;
  }
  const rowId = selectedRowId(selectedPolicy);
  if (!rowId) {
    throw new Error('Select a concrete row from Entrada_Politicas before generating.');
  }
  if (!selectedPolicy.id_politica_estoque) {
    throw new Error('The selected policy row is missing id_politica_estoque.');
  }

  generationRunning = true;
  el.button.disabled = true;
  setBadge('running', 'Running');
  setStatus('Reading source tables and validating the selected policy...');

  try {
    await updatePolicyRow(selectedPolicy, buildStartPatch(selectedPolicy));

    const [products, existingRows] = await Promise.all([
      fetchTableRows('Entrada_Produtos'),
      fetchTableRows('SaidaDados_Politica'),
    ]);

    const validProducts = products.filter((row) => row.id_produto);
    const existingCount = countExistingRows(existingRows, selectedPolicy);
    updatePreview(validProducts.length);

    if (!validProducts.length) {
      throw new Error('Entrada_Produtos returned no valid id_produto rows.');
    }
    if (existingCount > 0) {
      throw new Error(`This policy already has ${existingCount} detail rows in SaidaDados_Politica.`);
    }

    const addActions = buildAddActions(selectedPolicy, validProducts);
    setStatus(`Creating ${addActions.length} detailed rows in SaidaDados_Politica...`);

    for (const batch of chunk(addActions, 200)) {
      await applyActions(batch);
    }

    await updatePolicyRow(selectedPolicy, buildSuccessPatch(selectedPolicy, addActions.length));

    selectedPolicy = {
      ...selectedPolicy,
      generation_status: 'generated',
      generated_rows_count: addActions.length,
      generated_at: new Date().toISOString(),
      generation_error: '',
    };
    showPolicy(selectedPolicy);
    setBadge('success', 'Generated');
    setStatus(`Success. Created ${addActions.length} detailed rows for ${selectedPolicy.id_politica_estoque}.`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    try {
      await updatePolicyRow(selectedPolicy, buildErrorPatch(selectedPolicy, message));
    } catch (patchError) {
      console.error('Failed to persist generation error back to Grist.', patchError);
    }
    selectedPolicy = {
      ...selectedPolicy,
      generation_status: 'error',
      generation_error: message,
    };
    showPolicy(selectedPolicy);
    setBadge('error', 'Error');
    setStatus(`Generation failed: ${message}`);
    console.error(error);
  } finally {
    generationRunning = false;
    el.button.disabled = !selectedPolicy;
  }
}

function handleSelection(record) {
  selectedPolicy = record ? normalizeRow(record) : null;
  showPolicy(selectedPolicy);
  if (selectedPolicy) {
    setBadge('idle', 'Ready');
    setStatus('Selected policy ready. Click the button to generate detailed rows.');
    el.button.disabled = false;
  } else {
    setBadge('idle', 'Idle');
    setStatus('Waiting for a selected policy row.');
    el.button.disabled = true;
  }
}

el.button.addEventListener('click', () => {
  void generateDetailRows();
});

if (window.grist) {
  window.grist.ready({ requiredAccess: 'full', columns: REQUESTED_COLUMNS });
  window.grist.onRecord((record) => {
    handleSelection(record);
  });
} else {
  setStatus('This widget must run inside Grist.');
  setBadge('error', 'Unavailable');
}
