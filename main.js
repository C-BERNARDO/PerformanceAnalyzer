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
let allRows  = [];
let filtered = [];
let sortCol  = 'connected';
let sortDir  = 'desc';

/* ── Loading overlay helpers ────────────────────── */
function showLoading(filename) {
  loadingMsg.textContent = filename || 'Processing…';
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
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});
uploadZone.addEventListener('click', e => {
  // Let the <label for="fileInput"> handle its own click natively.
  // Only programmatically open the picker when clicking the zone background
  // or upload-inner areas that are NOT the label or hidden input itself.
  if (e.target.closest('label') || e.target === fileInput) return;
  if (e.target === uploadZone || e.target.closest('.upload-inner')) {
    fileInput.click();
  }
});
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

/* ── Clear ──────────────────────────────────────── */
clearBtn.addEventListener('click', reset);

function reset() {
  allRows  = [];
  filtered = [];
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

/* ── File handler ───────────────────────────────── */
function handleFile(file) {
  reset();
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) {
    showError('Invalid file type. Please upload an .xlsx or .xls file.');
    return;
  }

  showLoading(file.name);

  const reader = new FileReader();
  reader.onload = e => {
    // Small yield so overlay renders before heavy JS runs
    setTimeout(() => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb   = XLSX.read(data, { type: 'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });

        if (!json.length) {
          hideLoading();
          showError('The file appears to be empty.');
          return;
        }

        /* ── Column detection (case-insensitive, trim) ── */
        const keys    = Object.keys(json[0]);
        const findCol = label =>
          keys.find(k => k.trim().toLowerCase() === label.toLowerCase()) || null;

        const colRemark   = findCol('Remark By');
        const colDuration = findCol('Call Duration');
        const colAccount  = findCol('Account No.');
        const colPtpAmt   = findCol('PTP Amount');
        const colClaimAmt = findCol('Claim Paid Amount');

        const missing = [];
        if (!colRemark)   missing.push('"Remark By"');
        if (!colDuration) missing.push('"Call Duration"');
        if (!colAccount)  missing.push('"Account No."');
        if (missing.length) {
          hideLoading();
          showError(`Missing required column(s): ${missing.join(', ')}. Please check your file.`);
          return;
        }

        /* ── Filter: keep rows with a valid Call Duration ── */
        const INVALID_DURATION = ['', '00:00:00'];
        const validRows = json.filter(row => {
          const dur = String(row[colDuration]).trim();
          return !INVALID_DURATION.includes(dur);
        });

        if (!validRows.length) {
          hideLoading();
          showError('No valid call duration rows found after filtering blanks and "00:00:00" entries.');
          return;
        }

        /* ── touchMap: unique Account No. across ALL rows ── */
        const touchMap = {};
        json.forEach(row => {
          const name    = String(row[colRemark]).trim() || '(blank)';
          const account = String(row[colAccount]).trim();
          if (!touchMap[name]) touchMap[name] = new Set();
          if (account) touchMap[name].add(account);
        });

        /* ── connectedMap: unique Account No. from valid-duration rows ── */
        const connMap = {};
        validRows.forEach(row => {
          const name    = String(row[colRemark]).trim() || '(blank)';
          const account = String(row[colAccount]).trim();
          if (!connMap[name]) connMap[name] = new Set();
          if (account) connMap[name].add(account);
        });

        /* ── PTP grouping by Remark By ── */
        const ptpMap = {};
        if (colPtpAmt) {
          json.forEach(row => {
            const name   = String(row[colRemark]).trim() || '(blank)';
            const amt    = parseFloat(String(row[colPtpAmt]).replace(/,/g, '').trim());
            const valid  = !isNaN(amt) && amt !== 0;
            if (!ptpMap[name]) ptpMap[name] = { count: 0, amount: 0 };
            if (valid) { ptpMap[name].count++; ptpMap[name].amount += amt; }
          });
        }

        /* ── Kept grouping by Remark By (Claim Paid Amount) ── */
        const keptMap = {};
        if (colClaimAmt) {
          json.forEach(row => {
            const name  = String(row[colRemark]).trim() || '(blank)';
            const amt   = parseFloat(String(row[colClaimAmt]).replace(/,/g, '').trim());
            const valid = !isNaN(amt) && amt !== 0;
            if (!keptMap[name]) keptMap[name] = { count: 0, amount: 0 };
            if (valid) { keptMap[name].count++; keptMap[name].amount += amt; }
          });
        }

        /* ── Merge & exclude system/blank names ── */
        const allNames = new Set(
          [...Object.keys(touchMap), ...Object.keys(connMap)]
            .filter(n => n !== '(blank)' && n.toLowerCase() !== 'system')
        );

        allRows = [...allNames].map(name => {
          const ptp  = ptpMap[name]  || null;
          const kept = keptMap[name] || null;
          return {
            name,
            connected:     connMap[name]   ? connMap[name].size   : 0,
            touchAccounts: touchMap[name]  ? touchMap[name].size  : 0,
            ptpCount:      ptp  ? ptp.count   : 0,
            ptpAmount:     ptp  ? ptp.amount  : 0,
            keptCount:     kept ? kept.count  : 0,
            keptAmount:    kept ? kept.amount : 0,
          };
        });

        /* ── Update file info strip ── */
        fileNameEl.textContent = file.name;
        rowCountEl.textContent = `${json.length.toLocaleString()} rows`;
        fileInfo.style.display = 'flex';

        hideLoading();

        /* ── Render ── */
        applySort();
        renderResults(json.length, validRows.length);

      } catch (err) {
        hideLoading();
        showError('Could not parse the file. Make sure it is a valid Excel file.');
        console.error(err);
      }
    }, 60);
  };
  reader.readAsArrayBuffer(file);
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
  resultsMeta.textContent =
    `${n} collector${n !== 1 ? 's' : ''} · ` +
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
      `<tr><td colspan="7"><div class="empty-state">No collectors match your search.</div></td></tr>`;
    tableFooter.textContent = '';
    return;
  }

  tableBody.innerHTML = filtered.map((r, i) => `
    <tr>
      <td><span class="row-idx">${i + 1}</span>${escHtml(r.name)}</td>
      <td class="col-num"><span class="num num-touch">${r.touchAccounts.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-connected">${r.connected.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-ptp-count">${r.ptpCount.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-ptp-amt">${fmt(r.ptpAmount)}</span></td>
      <td class="col-num"><span class="num num-kept-count">${r.keptCount.toLocaleString()}</span></td>
      <td class="col-num"><span class="num num-kept-amt">${fmt(r.keptAmount)}</span></td>
    </tr>
  `).join('');

  const tTouch = filtered.reduce((s, r) => s + r.touchAccounts, 0);
  const tConn  = filtered.reduce((s, r) => s + r.connected, 0);
  const tPtpC  = filtered.reduce((s, r) => s + r.ptpCount, 0);
  const tPtpA  = filtered.reduce((s, r) => s + r.ptpAmount, 0);
  const tKeptC = filtered.reduce((s, r) => s + r.keptCount, 0);
  const tKeptA = filtered.reduce((s, r) => s + r.keptAmount, 0);

  tableFooter.textContent =
    `Showing ${filtered.length} of ${allRows.length} collectors  ·  ` +
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
