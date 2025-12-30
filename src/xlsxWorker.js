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

function newAggRecord(statusList) {
  const byStatus = {};
  for (const s of statusList) byStatus[s] = 0;
  return {
    byStatus,
    total: 0
  };
}

function addToAggRecord(agg, status, inc = 1) {
  agg.total += inc;
  agg.byStatus[status] = (agg.byStatus[status] ?? 0) + inc;
}

function computePrevYearTargets(prevRows) {
  const perOU0Counts = new Map();
  let grand = 0;
  for (const r of prevRows) {
    if (!r.applicationKey) continue;
    grand += 1;
    perOU0Counts.set(r.ou0, (perOU0Counts.get(r.ou0) ?? 0) + 1);
  }

  const perOU0Targets = {};
  perOU0Counts.forEach((count, ou0) => {
    perOU0Targets[ou0] = Math.ceil(count * 1.1);
  });

  return {
    perOU0Targets,
    grandTarget: Math.ceil(grand * 1.1),
    grandCount: grand
  };
}

function buildPivotFlat(rows) {
  const filtered = rows.filter((r) => r.applicationKey.length > 0);
  const statuses = sortStatuses(filtered.map((r) => r.status)).filter(
    (s) => normalizeHeader(s) !== normalizeHeader('Draft')
  );

  const grandAgg = newAggRecord(statuses);
  const ou0Map = new Map();

  for (const r of filtered) {
    addToAggRecord(grandAgg, r.status, 1);

    if (!ou0Map.has(r.ou0)) {
      ou0Map.set(r.ou0, {
        key: r.ou0,
        agg: newAggRecord(statuses)
      });
    }
    addToAggRecord(ou0Map.get(r.ou0).agg, r.status, 1);
  }

  const ou0Aggs = Array.from(ou0Map.values()).sort((a, b) => a.key.localeCompare(b.key));

  return {
    statuses,
    grandAgg,
    ou0Aggs,
    filteredCount: filtered.length,
    totalCount: rows.length
  };
}

self.onmessage = (e) => {
  const { id, currentBuf, prevBuf } = e.data || {};
  try {
    const currentWb = XLSX.read(currentBuf, { type: 'array' });
    const prevWb = XLSX.read(prevBuf, { type: 'array' });

    const currentSheet = assertDashboardSheet(currentWb);
    const prevSheet = assertDashboardSheet(prevWb);

    const currentRows = readRowsFromSheet(currentSheet);
    const prevRows = readRowsFromSheet(prevSheet);

    const prev = computePrevYearTargets(prevRows);
    const pivot = buildPivotFlat(currentRows);

    self.postMessage({
      id,
      ok: true,
      pivot,
      prev
    });
  } catch (err) {
    self.postMessage({
      id,
      ok: false,
      error: err?.message ? String(err.message) : String(err)
    });
  }
};
