import './styles.css';
import * as XLSX from 'xlsx';

const SHEET_NAME = 'Dashboard';
const REQUIRED_COLUMNS = [
  'OU Level 0',
  'OU Level 1',
  'OU Level 2',
  'Application Key',
  'Submission Status'
];

const STATUS_PRIORITY = [
  'Draft',
  'Submitted',
  'Approved',
  'Rejected',
  'Returned',
  'In Review',
  'In-Review',
  'Resubmitted',
  'Cancelled'
];

function normalizeHeader(v) {
  return String(v ?? '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function normalizeCell(v) {
  const s = String(v ?? '').trim();
  return s.length ? s : '(blank)';
}

function setStatus(message, kind = 'info') {
  const el = document.getElementById('status');
  el.textContent = message;
  el.dataset.kind = kind;
}

function clearOutput() {
  const wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '';
}

function assertDashboardSheet(workbook) {
  const sheet = workbook.Sheets?.[SHEET_NAME];
  if (!sheet) {
    const available = (workbook.SheetNames || []).join(', ');
    throw new Error(`Sheet “${SHEET_NAME}” not found. Available sheets: ${available || '(none)'}`);
  }
  return sheet;
}

function readRowsFromSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: '',
    raw: false
  });

  if (!Array.isArray(rows) || rows.length === 0) return [];

  const headerMap = new Map();
  for (const key of Object.keys(rows[0])) {
    headerMap.set(normalizeHeader(key), key);
  }

  const resolved = {};
  const missing = [];
  for (const req of REQUIRED_COLUMNS) {
    const actual = headerMap.get(normalizeHeader(req));
    if (!actual) missing.push(req);
    else resolved[req] = actual;
  }

  if (missing.length) {
    throw new Error(`Missing required columns in “${SHEET_NAME}”: ${missing.join(', ')}`);
  }

  return rows.map((r) => ({
    ou0: normalizeCell(r[resolved['OU Level 0']]),
    ou1: normalizeCell(r[resolved['OU Level 1']]),
    ou2: normalizeCell(r[resolved['OU Level 2']]),
    applicationKey: String(r[resolved['Application Key']] ?? '').trim(),
    status: normalizeCell(r[resolved['Submission Status']])
  }));
}

function sortStatuses(statuses) {
  const unique = Array.from(new Set(statuses));

  const prioIndex = new Map();
  STATUS_PRIORITY.forEach((s, i) => prioIndex.set(normalizeHeader(s), i));

  unique.sort((a, b) => {
    const ai = prioIndex.has(normalizeHeader(a)) ? prioIndex.get(normalizeHeader(a)) : 9999;
    const bi = prioIndex.has(normalizeHeader(b)) ? prioIndex.get(normalizeHeader(b)) : 9999;
    if (ai !== bi) return ai - bi;
    return a.localeCompare(b);
  });

  return unique;
}

function newAgg(statusList) {
  const byStatus = new Map();
  for (const s of statusList) byStatus.set(s, 0);
  return {
    byStatus,
    total: 0
  };
}

function addToAgg(agg, status, inc = 1) {
  agg.total += inc;
  agg.byStatus.set(status, (agg.byStatus.get(status) ?? 0) + inc);
}

function buildPivot(rows) {
  const filtered = rows.filter((r) => r.applicationKey.length > 0);
  const statuses = sortStatuses(filtered.map((r) => r.status));

  const root = {
    children: new Map(),
    agg: newAgg(statuses)
  };

  for (const r of filtered) {
    addToAgg(root.agg, r.status, 1);

    if (!root.children.has(r.ou0)) {
      root.children.set(r.ou0, {
        key: r.ou0,
        children: new Map(),
        agg: newAgg(statuses)
      });
    }
    const n0 = root.children.get(r.ou0);
    addToAgg(n0.agg, r.status, 1);

    if (!n0.children.has(r.ou1)) {
      n0.children.set(r.ou1, {
        key: r.ou1,
        children: new Map(),
        agg: newAgg(statuses)
      });
    }
    const n1 = n0.children.get(r.ou1);
    addToAgg(n1.agg, r.status, 1);

    if (!n1.children.has(r.ou2)) {
      n1.children.set(r.ou2, {
        key: r.ou2,
        children: new Map(),
        agg: newAgg(statuses)
      });
    }
    const n2 = n1.children.get(r.ou2);
    addToAgg(n2.agg, r.status, 1);
  }

  return { root, statuses };
}

function formatNumber(n) {
  return new Intl.NumberFormat('en-US').format(n);
}

function renderPivot({ root, statuses }) {
  const wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '';

  if (!root.children.size) {
    wrap.innerHTML = '<div class="empty">No rows found with a non-empty Application Key.</div>';
    return;
  }

  const table = document.createElement('table');
  table.className = 'pivot';

  const thead = document.createElement('thead');
  const htr = document.createElement('tr');

  const h0 = document.createElement('th');
  h0.textContent = 'BG-Unit-Subunit';
  h0.className = 'row-header';
  htr.appendChild(h0);

  for (const s of statuses) {
    const th = document.createElement('th');
    th.textContent = s;
    htr.appendChild(th);
  }

  const thGT = document.createElement('th');
  thGT.textContent = 'Grand Total';
  thGT.className = 'grand-total';
  htr.appendChild(thGT);

  thead.appendChild(htr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  function appendRow(label, level, agg, rowKind = 'normal') {
    const tr = document.createElement('tr');
    tr.dataset.level = String(level);
    tr.dataset.kind = rowKind;

    const tdLabel = document.createElement('td');
    tdLabel.className = 'label';
    tdLabel.style.paddingLeft = `${8 + level * 18}px`;
    tdLabel.textContent = label;
    tr.appendChild(tdLabel);

    for (const s of statuses) {
      const td = document.createElement('td');
      td.className = 'num';
      td.textContent = formatNumber(agg.byStatus.get(s) ?? 0);
      tr.appendChild(td);
    }

    const tdTotal = document.createElement('td');
    tdTotal.className = 'num grand-total';
    tdTotal.textContent = formatNumber(agg.total ?? 0);
    tr.appendChild(tdTotal);

    tbody.appendChild(tr);
  }

  const ou0Keys = Array.from(root.children.keys()).sort((a, b) => a.localeCompare(b));
  for (const ou0 of ou0Keys) {
    const n0 = root.children.get(ou0);
    appendRow(n0.key, 0, n0.agg, 'group0');

    const ou1Keys = Array.from(n0.children.keys()).sort((a, b) => a.localeCompare(b));
    for (const ou1 of ou1Keys) {
      const n1 = n0.children.get(ou1);
      appendRow(n1.key, 1, n1.agg, 'group1');

      const ou2Keys = Array.from(n1.children.keys()).sort((a, b) => a.localeCompare(b));
      for (const ou2 of ou2Keys) {
        const n2 = n1.children.get(ou2);
        appendRow(n2.key, 2, n2.agg, 'leaf');
      }
    }
  }

  appendRow('Grand Total', 0, root.agg, 'grand');

  table.appendChild(tbody);
  wrap.appendChild(table);
}

async function onFileSelected(file) {
  clearOutput();

  if (!file) {
    setStatus('No file selected.', 'info');
    return;
  }

  setStatus(`Reading ${file.name}...`, 'info');

  const buf = await file.arrayBuffer();
  const workbook = XLSX.read(buf, { type: 'array' });

  const sheet = assertDashboardSheet(workbook);
  const rows = readRowsFromSheet(sheet);

  const { root, statuses } = buildPivot(rows);
  renderPivot({ root, statuses });

  setStatus(`Rendered ${rows.length} rows from “${SHEET_NAME}”.`, 'success');
}

function init() {
  const input = document.getElementById('fileInput');
  input.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    try {
      await onFileSelected(file);
    } catch (err) {
      clearOutput();
      setStatus(err?.message ? String(err.message) : 'Failed to process file.', 'error');
    }
  });

  setStatus('Select an .xlsx file to generate the pivot table.', 'info');
}

init();
