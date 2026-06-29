/* ── DOM refs ───────────────────────────────────── */
const uploadZone     = document.getElementById('uploadZone');
const fileInput      = document.getElementById('fileInput');
const fileInfo       = document.getElementById('fileInfo');
const fileNameEl     = document.getElementById('fileName');
const rowCountEl     = document.getElementById('rowCount');
const clearBtn       = document.getElementById('clearBtn');
const errorBox       = document.getElementById('errorBox');
const tableBody      = document.getElementById('tableBody');
const resultsMeta    = document.getElementById('resultsMeta');
const tableFooter    = document.getElementById('tableFooter');
const searchInput    = document.getElementById('searchInput');
const ths            = document.querySelectorAll('th.sortable');
const tablePlaceholder = document.getElementById('tablePlaceholder');
const tableArea      = document.getElementById('tableArea');
const loadingOverlay = document.getElementById('loadingOverlay');
const loadingMsg     = document.getElementById('loadingMsg');

/* ── State ──────────────────────────────────────── */
let allRows     = [];
let filtered    = [];
let sortCol     = 'connected';
let sortDir     = 'desc';
let uploadedFiles = []; // { name, rowCount } for each accepted file

/* ── Loading overlay helpers ────────────────────── */
function showLoading(msg) {
  loadingMsg.textContent = msg || 'Processing…';
  loadingOverlay.classList.add('active');
}
function hideLoading() {
  loadingOverlay.classList.remove('active');
}

/* ── Drag-and-drop ──────────────────────────────── */
uploadZone.addEventListener('dragover', e => {
  e.preventDefault();
  uploadZone.classList.add('drag-over');
});
['dragleave', 'dragend'].forEach(ev =>
  uploadZone.addEventListener(ev, () => uploadZone.classList.remove('drag-over'))
);
uploadZone.addEventListener('drop', e => {
  e.preventDefault();
  uploadZone.classList.remove('drag-over');
  const files = [...e.dataTransfer.files];
  if (files.length) handleFiles(files);
});
uploadZone.addEventListener('click', e => {
  if (e.target.closest('label') || e.target === fileInput) return;
  if (e.target === uploadZone || e.target.closest('.upload-inner')) fileInput.click();
});
fileInput.addEventListener('change', () => {
  if (fileInput.files.length) handleFiles([...fileInput.files]);
});

/* ── Clear ──────────────────────────────────────── */
clearBtn.addEventListener('click', reset);

function reset() {
  allRows       = [];
  filtered      = [];
  uploadedFiles = [];
  fileInfo.style.display          = 'none';
  errorBox.style.display          = 'none';
  tablePlaceholder.style.display  = 'flex';
  tableArea.style.display         = 'none';
  resultsMeta.textContent         = '';
  fileInput.value                 = '';
  searchInput.value               = '';
  tableBody.innerHTML             = '';
  tableFooter.textContent         = '';
}

/* ── Read one file → parsed JSON rows ───────────── */
function readFileAsJSON(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb   = XLSX.read(data, { type: 'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(ws, { defval: '' }));
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('FileReader error'));
    reader.readAsArrayBuffer(file);
  });
}

/* ── Multi-file handler ─────────────────────────── */
async function handleFiles(files) {
  reset();

  // Validate extensions up front
  const invalid = files.filter(f => !['xlsx','xls'].includes(f.name.split('.').pop().toLowerCase()));
  if (invalid.length) {
    showError(`Invalid file type${invalid.length > 1 ? 's' : ''}: ${invalid.map(f => f.name).join(', ')}. Please upload .xlsx or .xls files only.`);
    return;
  }

  showLoading(files.length === 1 ? `Reading ${files[0].name}…` : `Merging ${files.length} files…`);

  try {
    // Parse all files in parallel
    const allParsed = await Promise.all(files.map(f => readFileAsJSON(f)));

    // Validate each file before merging
    const REQUIRED = ['Remark By', 'Call Duration', 'Account No.'];
    for (let i = 0; i < files.length; i++) {
      if (!allParsed[i].length) {
        hideLoading();
        showError(`"${files[i].name}" appears to be empty.`);
        return;
      }
      const keys    = Object.keys(allParsed[i][0]);
      const findCol = label => keys.find(k => k.trim().toLowerCase() === label.toLowerCase()) || null;
      const missing = REQUIRED.filter(r => !findCol(r)).map(r => `"${r}"`);
      if (missing.length) {
        hideLoading();
        showError(`"${files[i].name}" is missing required column(s): ${missing.join(', ')}.`);
        return;
      }
    }

    // Record per-file stats for info strip
    uploadedFiles = files.map((f, i) => ({ name: f.name, rowCount: allParsed[i].length }));

    // ── Merge all rows into one flat array ──────────
    const merged = allParsed.flat();

    // Column detection on merged dataset
    const keys    = Object.keys(merged[0]);
    const findCol = label => keys.find(k => k.trim().toLowerCase() === label.toLowerCase()) || null;

    const colRemark     = findCol('Remark By');
    const colDuration   = findCol('Call Duration');
    const colAccount    = findCol('Account No.');
    const colPtpAmt     = findCol('PTP Amount');
    const colClaimAmt   = findCol('Claim Paid Amount');
    const colRemarkType = findCol('Remark Type');

    // ── Filter: keep rows with a valid Call Duration ──
    const INVALID_DURATION = ['', '00:00:00'];
    const validRows = merged.filter(row => !INVALID_DURATION.includes(String(row[colDuration]).trim()));

    if (!validRows.length) {
      hideLoading();
      showError('No valid call duration rows found after filtering blanks and "00:00:00" entries.');
      return;
    }

    // ── touchMap: unique Account No. across ALL merged rows ──
    const touchMap = {};
    merged.forEach(row => {
      const name    = String(row[colRemark]).trim() || '(blank)';
      const account = String(row[colAccount]).trim();
      if (!touchMap[name]) touchMap[name] = new Set();
      if (account) touchMap[name].add(account);
    });

    // ── connectedMap: unique Account No. from valid-duration rows ──
    const connMap = {};
    validRows.forEach(row => {
      const name    = String(row[colRemark]).trim() || '(blank)';
      const account = String(row[colAccount]).trim();
      if (!connMap[name]) connMap[name] = new Set();
      if (account) connMap[name].add(account);
    });

    // ── outgoingMap: Remark Type === "Outgoing" ──
    const outgoingMap = {};
    if (colRemarkType) {
      merged.forEach(row => {
        if (String(row[colRemarkType]).trim().toLowerCase() !== 'outgoing') return;
        const name    = String(row[colRemark]).trim() || '(blank)';
        const account = String(row[colAccount]).trim();
        if (!outgoingMap[name]) outgoingMap[name] = new Set();
        if (account) outgoingMap[name].add(account);
      });
    }

    // ── predictiveMap: Remark Type === "Predictive" ──
    const predictiveMap = {};
    if (colRemarkType) {
      merged.forEach(row => {
        if (String(row[colRemarkType]).trim().toLowerCase() !== 'predictive') return;
        const name    = String(row[colRemark]).trim() || '(blank)';
        const account = String(row[colAccount]).trim();
        if (!predictiveMap[name]) predictiveMap[name] = new Set();
        if (account) predictiveMap[name].add(account);
      });
    }

    // ── ptpMap ──
    const ptpMap = {};
    if (colPtpAmt) {
      merged.forEach(row => {
        const name  = String(row[colRemark]).trim() || '(blank)';
        const amt   = parseFloat(String(row[colPtpAmt]).replace(/,/g, '').trim());
        const valid = !isNaN(amt) && amt !== 0;
        if (!ptpMap[name]) ptpMap[name] = { count: 0, amount: 0 };
        if (valid) { ptpMap[name].count++; ptpMap[name].amount += amt; }
      });
    }

    // ── keptMap ──
    const keptMap = {};
    if (colClaimAmt) {
      merged.forEach(row => {
        const name  = String(row[colRemark]).trim() || '(blank)';
        const amt   = parseFloat(String(row[colClaimAmt]).replace(/,/g, '').trim());
        const valid = !isNaN(amt) && amt !== 0;
        if (!keptMap[name]) keptMap[name] = { count: 0, amount: 0 };
        if (valid) { keptMap[name].count++; keptMap[name].amount += amt; }
      });
    }

    // ── Build final collector rows ──
    const allNames = new Set(
      [...Object.keys(touchMap), ...Object.keys(connMap)]
        .filter(n => n !== '(blank)' && n.toLowerCase() !== 'system')
    );

    allRows = [...allNames].map(name => {
      const ptp  = ptpMap[name]  || null;
      const kept = keptMap[name] || null;
      return {
        name,
        outgoingCall:   outgoingMap[name]   ? outgoingMap[name].size   : 0,
        predictiveCall: predictiveMap[name] ? predictiveMap[name].size : 0,
        touchAccounts:  touchMap[name]      ? touchMap[name].size      : 0,
        connected:      connMap[name]       ? connMap[name].size       : 0,
        ptpCount:       ptp  ? ptp.count   : 0,
        ptpAmount:      ptp  ? ptp.amount  : 0,
        keptCount:      kept ? kept.count  : 0,
        keptAmount:     kept ? kept.amount : 0,
      };
    });

    updateFileInfoStrip();
    hideLoading();
    applySort();
    renderResults(merged.length, validRows.length);

  } catch (err) {
    hideLoading();
    showError('Could not parse one or more files. Make sure they are valid Excel files.');
    console.error(err);
  }
}

/* ── File info strip ────────────────────────────── */
function updateFileInfoStrip() {
  if (!uploadedFiles.length) { fileInfo.style.display = 'none'; return; }
  const total = uploadedFiles.reduce((s, f) => s + f.rowCount, 0);
  if (uploadedFiles.length === 1) {
    fileNameEl.textContent = uploadedFiles[0].name;
    fileNameEl.title       = '';
    rowCountEl.textContent = `${total.toLocaleString()} rows`;
  } else {
    fileNameEl.textContent = `${uploadedFiles.length} files merged`;
    fileNameEl.title       = uploadedFiles.map(f => `${f.name} (${f.rowCount.toLocaleString()} rows)`).join('\n');
    rowCountEl.textContent = `${total.toLocaleString()} rows total`;
  }
  fileInfo.style.display = 'flex';
}

/* ── Sort ───────────────────────────────────────── */
ths.forEach(th => {
  th.addEventListener('click', () => {
    const col = th.dataset.col;
    sortDir = (sortCol === col) ? (sortDir === 'asc' ? 'desc' : 'asc')
                                : (col === 'name' ? 'asc' : 'desc');
    sortCol = col;
    applySort();
    renderTable();
  });
});

function applySort() {
  allRows.sort((a, b) => {
    let va = a[sortCol], vb = b[sortCol];
    if (typeof va === 'string') { va = va.toLowerCase(); vb = vb.toLowerCase(); }
    if (va < vb) return sortDir === 'asc' ? -1 :  1;
    if (va > vb) return sortDir === 'asc' ?  1 : -1;
    return 0;
  });
  ths.forEach(t => {
    t.classList.remove('sort-asc', 'sort-desc');
    if (t.dataset.col === sortCol)
      t.classList.add(sortDir === 'asc' ? 'sort-asc' : 'sort-desc');
  });
}

/* ── Search ─────────────────────────────────────── */
searchInput.addEventListener('input', renderTable);

/* ── Render ─────────────────────────────────────── */
function renderResults(totalRows, validRows) {
  const n = allRows.length;
  const fileLabel = uploadedFiles.length > 1 ? `${uploadedFiles.length} files merged · ` : '';
  resultsMeta.textContent =
    `${fileLabel}${n} collector${n !== 1 ? 's' : ''} · ` +
    `${validRows.toLocaleString()} valid-duration rows` +
    ` (${(totalRows - validRows).toLocaleString()} excluded)`;

  tablePlaceholder.style.display = 'none';
  tableArea.style.display        = 'flex';
  renderTable();
}

function renderTable() {
  const q  = searchInput.value.trim().toLowerCase();
  filtered = q ? allRows.filter(r => r.name.toLowerCase().includes(q)) : [...allRows];

  if (!filtered.length) {
    tableBody.innerHTML =
      `<tr><td colspan="9"><div class="empty-state">No collectors match your search.</div></td></tr>`;
    tableFooter.textContent = '';
    return;
  }

  tableBody.innerHTML = filtered.map((r, i) => `
    <tr>
      <td>${escHtml(r.name)}</td>
      <td class="col-num"><span class="num num-outgoing">${r.outgoingCall.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-predictive">${r.predictiveCall.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-touch">${r.touchAccounts.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-connected">${r.connected.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-ptp-count">${r.ptpCount.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-ptp-amt">${fmt(r.ptpAmount)}</span></td>
      <td class="col-num"><span class="num num-kept-count">${r.keptCount.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-kept-amt">${fmt(r.keptAmount)}</span></td>
    </tr>
  `).join('');

  const tOut   = filtered.reduce((s, r) => s + r.outgoingCall, 0);
  const tPred  = filtered.reduce((s, r) => s + r.predictiveCall, 0);
  const tTouch = filtered.reduce((s, r) => s + r.touchAccounts, 0);
  const tConn  = filtered.reduce((s, r) => s + r.connected, 0);
  const tPtpC  = filtered.reduce((s, r) => s + r.ptpCount, 0);
  const tPtpA  = filtered.reduce((s, r) => s + r.ptpAmount, 0);
  const tKeptC = filtered.reduce((s, r) => s + r.keptCount, 0);
  const tKeptA = filtered.reduce((s, r) => s + r.keptAmount, 0);

  tableFooter.textContent =
    `Showing ${filtered.length} of ${allRows.length} collectors  ·  ` +
    `${tOut.toLocaleString()} outgoing  ·  ${tPred.toLocaleString()} predictive  ·  ` +
    `${tTouch.toLocaleString()} touch  ·  ${tConn.toLocaleString()} connected  ·  ` +
    `${tPtpC.toLocaleString()} PTP · ${fmt(tPtpA)} PTP amt  ·  ` +
    `${tKeptC.toLocaleString()} kept · ${fmt(tKeptA)} kept amt`;
}

/* ── Helpers ────────────────────────────────────── */
function showError(msg) {
  errorBox.textContent = '⚠ ' + msg;
  errorBox.style.display = 'block';
}
function escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function fmt(v) {
  if (v === 0) return '0.00';
  return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}