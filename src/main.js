import './styles.css';
import * as XLSX from 'xlsx';
import { toJpeg, toPng } from 'html-to-image';

const pivotWorker = new Worker(new URL('./xlsxWorker.js', import.meta.url), { type: 'module' });
let activeProcessId = 0;

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

function setExportEnabled(enabled) {
  const btn = document.getElementById('exportBtn');
  if (btn) btn.disabled = !enabled;
}

function clearOutput() {
  const wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '';
  setExportEnabled(false);
}

let currentFile = null;
let prevYearFile = null;

const prevTargets = {
  perOU0: new Map(),
  grand: null,
  available: false
};

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
  const statuses = sortStatuses(filtered.map((r) => r.status)).filter(
    (s) =>
      normalizeHeader(s) !== normalizeHeader('Draft') &&
      normalizeHeader(s) !== normalizeHeader('Submitted')
  );

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

function computeTarget(total) {
  const t = Number(total ?? 0);
  return Math.ceil(t * 1.1);
}

function resetPrevTargets() {
  prevTargets.perOU0 = new Map();
  prevTargets.grand = null;
  prevTargets.available = false;
}

async function parseWorkbookFromFile(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: 'array' });
}

function computePrevYearTargets(prevRows) {
  const perOU0Counts = new Map();
  let grand = 0;
  for (const r of prevRows) {
    if (!r.applicationKey) continue;
    grand += 1;
    perOU0Counts.set(r.ou0, (perOU0Counts.get(r.ou0) ?? 0) + 1);
  }

  const perOU0Targets = new Map();
  perOU0Counts.forEach((count, ou0) => perOU0Targets.set(ou0, Math.ceil(count * 1.1)));

  return {
    perOU0Targets,
    grandTarget: Math.ceil(grand * 1.1),
    grandCount: grand
  };
}

async function processIfReady() {
  clearOutput();

  if (!currentFile || !prevYearFile) {
    setStatus('Upload both files (Current Year and Prev Year) to generate the pivot.', 'info');
    return;
  }

  setStatus('Processing files...', 'info');
  resetPrevTargets();

  const myId = ++activeProcessId;
  const [currentBuf, prevBuf] = await Promise.all([
    currentFile.arrayBuffer(),
    prevYearFile.arrayBuffer()
  ]);

  const result = await new Promise((resolve, reject) => {
    const onMessage = (ev) => {
      const data = ev.data;
      if (!data || data.id !== myId) return;
      pivotWorker.removeEventListener('message', onMessage);
      pivotWorker.removeEventListener('error', onError);
      resolve(data);
    };
    const onError = (err) => {
      pivotWorker.removeEventListener('message', onMessage);
      pivotWorker.removeEventListener('error', onError);
      reject(err);
    };
    pivotWorker.addEventListener('message', onMessage);
    pivotWorker.addEventListener('error', onError);
    pivotWorker.postMessage({ id: myId, currentBuf, prevBuf }, [currentBuf, prevBuf]);
  });

  if (myId !== activeProcessId) return;

  if (!result.ok) {
    throw new Error(result.error || 'Failed to process files.');
  }

  const { pivot, prev } = result;
  prevTargets.perOU0 = new Map(Object.entries(prev.perOU0Targets || {}));
  prevTargets.grand = prev.grandTarget;
  prevTargets.available = true;

  renderPivotFlat(pivot);

  setStatus(
    `Rendered current year (${pivot.totalCount} rows). Prev year base: ${formatNumber(prev.grandCount)} | Target (Grand): ${formatNumber(prev.grandTarget)}.`,
    'success'
  );
}

function renderPivotFlat(pivot) {
  const wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '';

  if (!pivot?.ou0Aggs?.length) {
    wrap.innerHTML = '<div class="empty">No rows found with a non-empty Application Key.</div>';
    setExportEnabled(false);
    return;
  }

  const statuses = Array.isArray(pivot.statuses) ? pivot.statuses : [];

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

  const thTargetValue = document.createElement('th');
  thTargetValue.textContent = 'Target Value';
  thTargetValue.className = 'target-value-header';
  htr.appendChild(thTargetValue);

  const thTargetIllustration = document.createElement('th');
  thTargetIllustration.textContent = 'Target Illustration';
  thTargetIllustration.className = 'target-illustration-header';
  htr.appendChild(thTargetIllustration);

  thead.appendChild(htr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  const getCount = (agg, status) => Number(agg?.byStatus?.[status] ?? 0);
  const getTotal = (agg) => Number(agg?.total ?? 0);

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
      td.textContent = formatNumber(getCount(agg, s));
      tr.appendChild(td);
    }

    const tdTotal = document.createElement('td');
    tdTotal.className = 'num grand-total';
    tdTotal.textContent = formatNumber(getTotal(agg));
    tr.appendChild(tdTotal);

    const current = getTotal(agg);
    let target;
    if (prevTargets.available && rowKind === 'grand' && typeof prevTargets.grand === 'number') {
      target = prevTargets.grand;
    } else if (prevTargets.available && rowKind === 'group0' && prevTargets.perOU0.has(label)) {
      target = prevTargets.perOU0.get(label);
    } else {
      target = computeTarget(current);
    }
    const rawRatio = target > 0 ? current / target : 1;
    const progress = Math.max(0, Math.min(rawRatio, 1));
    const progressPct = Math.round(progress * 100);
    const markerPos = Math.min(Math.max(progress * 100, 4), 92);
    const markerCssPos =
      progressPct >= 100 ? 'calc(100% - var(--target-pill-h, 13px))' : `${markerPos}%`;

    const tdTargetValue = document.createElement('td');
    tdTargetValue.className = 'num target-value';
    tdTargetValue.textContent = formatNumber(target);
    tr.appendChild(tdTargetValue);

    const tdTargetIllustration = document.createElement('td');
    tdTargetIllustration.className = 'target-illustration';
    tdTargetIllustration.title = `Current: ${formatNumber(current)} | Target: ${formatNumber(target)} | Progress: ${progressPct}%`;
    const track = document.createElement('div');
    track.className = 'target-track';
    track.style.setProperty('--p', markerCssPos);

    const pill = document.createElement('div');
    pill.className = 'target-pill';
    pill.style.setProperty('--fill', progressPct >= 100 ? '100%' : markerCssPos);
    pill.classList.add(
      progressPct <= 15
        ? 'fill-red'
        : progressPct <= 25
          ? 'fill-red-light'
          : progressPct <= 50
            ? 'fill-orange'
            : progressPct <= 80
              ? 'fill-green-light'
              : 'fill-green'
    );

    const marker = document.createElement('div');
    marker.className = 'target-golf';

    const pct = document.createElement('div');
    pct.className = 'target-golf-label';
    pct.textContent = `${progressPct}%`;
    marker.appendChild(pct);

    const pointer = document.createElement('div');
    pointer.className = 'target-golf-pointer';
    marker.appendChild(pointer);

    const ball = document.createElement('div');
    ball.className = 'target-golf-ball';
    marker.appendChild(ball);

    track.appendChild(pill);
    track.appendChild(marker);
    tdTargetIllustration.appendChild(track);

    tr.appendChild(tdTargetIllustration);
    tbody.appendChild(tr);
  }

  for (const row of pivot.ou0Aggs) {
    appendRow(row.key, 0, row.agg, 'group0');
  }

  appendRow('Grand Total', 0, pivot.grandAgg, 'grand');

  table.appendChild(tbody);
  wrap.appendChild(table);

  setExportEnabled(true);
}

function renderPivot({ root, statuses }) {
  const wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '';

  if (!root.children.size) {
    wrap.innerHTML = '<div class="empty">No rows found with a non-empty Application Key.</div>';
    setExportEnabled(false);
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

  const thTargetValue = document.createElement('th');
  thTargetValue.textContent = 'Target Value';
  thTargetValue.className = 'target-value-header';
  htr.appendChild(thTargetValue);

  const thTargetIllustration = document.createElement('th');
  thTargetIllustration.textContent = 'Target Illustration';
  thTargetIllustration.className = 'target-illustration-header';
  htr.appendChild(thTargetIllustration);

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

    const current = Number(agg.total ?? 0);
    let target;
    if (prevTargets.available && rowKind === 'grand' && typeof prevTargets.grand === 'number') {
      target = prevTargets.grand;
    } else if (prevTargets.available && rowKind === 'group0' && prevTargets.perOU0.has(label)) {
      target = prevTargets.perOU0.get(label);
    } else {
      target = computeTarget(current);
    }
    const rawRatio = target > 0 ? current / target : 1;
    const progress = Math.max(0, Math.min(rawRatio, 1));
    const progressPct = Math.round(progress * 100);
    const markerPos = Math.min(Math.max(progress * 100, 4), 92);
    const markerCssPos =
      progressPct >= 100 ? 'calc(100% - var(--target-pill-h, 13px))' : `${markerPos}%`;

    const tdTargetValue = document.createElement('td');
    tdTargetValue.className = 'num target-value';
    tdTargetValue.textContent = formatNumber(target);
    tr.appendChild(tdTargetValue);

    const tdTargetIllustration = document.createElement('td');
    tdTargetIllustration.className = 'target-illustration';
    tdTargetIllustration.title = `Current: ${formatNumber(current)} | Target: ${formatNumber(target)} | Progress: ${progressPct}%`;
    const track = document.createElement('div');
    track.className = 'target-track';
    track.style.setProperty('--p', markerCssPos);

    const pill = document.createElement('div');
    pill.className = 'target-pill';
    pill.style.setProperty('--fill', progressPct >= 100 ? '100%' : markerCssPos);
    pill.classList.add(
      progressPct <= 15
        ? 'fill-red'
        : progressPct <= 25
          ? 'fill-red-light'
          : progressPct <= 50
            ? 'fill-orange'
            : progressPct <= 80
              ? 'fill-green-light'
              : 'fill-green'
    );

    const marker = document.createElement('div');
    marker.className = 'target-golf';

    const pct = document.createElement('div');
    pct.className = 'target-golf-label';
    pct.textContent = `${progressPct}%`;
    marker.appendChild(pct);

    const pointer = document.createElement('div');
    pointer.className = 'target-golf-pointer';
    marker.appendChild(pointer);

    const ball = document.createElement('div');
    ball.className = 'target-golf-ball';
    marker.appendChild(ball);

    track.appendChild(pill);
    track.appendChild(marker);
    tdTargetIllustration.appendChild(track);

    tr.appendChild(tdTargetIllustration);

    tbody.appendChild(tr);
  }

  const ou0Keys = Array.from(root.children.keys()).sort((a, b) => {
    if (a === '(blank)' && b !== '(blank)') return 1;
    if (b === '(blank)' && a !== '(blank)') return -1;
    return a.localeCompare(b);
  });
  for (const ou0 of ou0Keys) {
    const n0 = root.children.get(ou0);
    appendRow(n0.key, 0, n0.agg, 'group0');
  }

  appendRow('Grand Total', 0, root.agg, 'grand');

  table.appendChild(tbody);
  wrap.appendChild(table);

  setExportEnabled(true);
}

function defaultExportFileName(ext) {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const stamp = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
  return `dashboard_pivot_${stamp}.${ext}`;
}

function downloadDataUrl(dataUrl, filename) {
  const a = document.createElement('a');
  a.href = dataUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function buildRowChunkedExportNode(sourceTable) {
  const exportRoot = document.createElement('div');
  exportRoot.className = 'export-row-chunks';
  exportRoot.style.cssText = 'position:absolute;left:0;top:0;background:#ffffff;padding:20px;z-index:-9999;pointer-events:none;';

  const grid = document.createElement('div');
  grid.style.cssText = 'display:flex;gap:20px;align-items:flex-start;';
  exportRoot.appendChild(grid);

  const thead = sourceTable.querySelector('thead');
  const bodyRows = Array.from(sourceTable.querySelectorAll('tbody tr'));

  const totalRows = bodyRows.length;
  const parts = totalRows <= 30 ? 1 : 1 + Math.ceil((totalRows - 30) / 15);
  const rowsPer = Math.ceil(totalRows / parts);

  for (let i = 0; i < parts; i += 1) {
    const start = i * rowsPer;
    if (start >= bodyRows.length) break;
    const end = Math.min(start + rowsPer, bodyRows.length);

    const panel = document.createElement('div');
    panel.style.cssText = 'background:#ffffff;border:2px solid #9aa4b2;border-radius:8px;padding:12px;';

    const t = document.createElement('table');
    t.className = sourceTable.className;
    t.style.cssText = 'width:max-content;min-width:0;';

    if (thead) t.appendChild(thead.cloneNode(true));

    const tbody = document.createElement('tbody');
    for (let r = start; r < end; r += 1) {
      tbody.appendChild(bodyRows[r].cloneNode(true));
    }
    t.appendChild(tbody);

    panel.appendChild(t);
    grid.appendChild(panel);
  }

  return exportRoot;
}

async function exportCurrentView() {
  const node = document.getElementById('tableWrap');
  if (!node) return;

  const sourceTable = node.querySelector('table');
  if (!sourceTable) {
    setStatus('Nothing to export yet. Upload an .xlsx file first.', 'info');
    return;
  }

  const exportNode = buildRowChunkedExportNode(sourceTable);
  document.body.appendChild(exportNode);

  if (document.fonts?.ready) {
    try { await document.fonts.ready; } catch { /* ignore */ }
  }
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));

  const formatEl = document.getElementById('exportFormat');
  const format = String(formatEl?.value ?? 'png').toLowerCase();

  const opts = {
    backgroundColor: '#ffffff',
    pixelRatio: 2,
    cacheBust: true
  };

  setStatus('Exporting image...', 'info');

  let dataUrl;
  try {
    dataUrl = format === 'jpeg'
      ? await toJpeg(exportNode, { ...opts, quality: 0.95 })
      : await toPng(exportNode, opts);
  } finally {
    exportNode.remove();
  }

  downloadDataUrl(dataUrl, defaultExportFileName(format === 'jpeg' ? 'jpg' : 'png'));
  setStatus('Export complete.', 'success');
}

async function onFileSelected(file) {
  currentFile = file ?? null;
  await processIfReady();
}

async function onPrevFileSelected(file) {
  prevYearFile = file ?? null;
  await processIfReady();
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

  const prevInput = document.getElementById('prevFileInput');
  prevInput?.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    try {
      await onPrevFileSelected(file);
    } catch (err) {
      clearOutput();
      setStatus(err?.message ? String(err.message) : 'Failed to process previous year file.', 'error');
    }
  });

  const exportBtn = document.getElementById('exportBtn');
  exportBtn?.addEventListener('click', async () => {
    try {
      await exportCurrentView();
    } catch (err) {
      setStatus(err?.message ? String(err.message) : 'Export failed.', 'error');
    }
  });

  setStatus('Upload both .xlsx files (Current and Prev Year) to generate the pivot table.', 'info');
  setExportEnabled(false);
}

init();
