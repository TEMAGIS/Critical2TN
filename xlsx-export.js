// ════════════════════════════════════════════════════════════════
//  xlsx-export.js — Excel (.xlsx) downloads for Critical2TN
//  --------------------------------------------------------------
//  Load it AFTER the main inline <script> in index.html:
//      <script src="xlsx-export.js"></script>
//  It adds:
//   • an "Excel" button in the list header that downloads the list
//     exactly as currently filtered (sector, county, search);
//   • a "Download Excel" button in the draw-an-area export sheet that
//     puts Critical2TN assets and Tier II into one workbook (two tabs).
//  Every workbook gets a "Notice" tab with the handling notice, who
//  exported it, when, and which filters/area were used.
//  The SheetJS library is only downloaded the first time someone
//  clicks an Excel button.
// ════════════════════════════════════════════════════════════════
(function () {
  const XLSX_SRC = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
  const LIST_EXPORT_CAP = 10000;   // max rows for the list download (statewide is ~5,300 today)

  // ── load SheetJS on demand
  let xlsxPromise = null;
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (xlsxPromise) return xlsxPromise;
    xlsxPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = XLSX_SRC; s.async = true;
      s.onload = () => window.XLSX ? resolve(window.XLSX) : reject(new Error('the Excel library did not load'));
      s.onerror = () => { xlsxPromise = null; reject(new Error('couldn’t load the Excel library — check your connection')); };
      document.head.appendChild(s);
    });
    return xlsxPromise;
  }

  // ── building blocks
  function cellValue(v, numeric) {
    if (v == null || v === '') return null;
    if (numeric) { const n = Number(v); return Number.isFinite(n) ? n : String(v); }
    return String(v);                        // written as text, never as a formula
  }
  // cols = CORE_COLS / T2_COLS from index.html: [header, getter, isText?]
  function dataSheet(XLSX, cols, rows) {
    const aoa = [cols.map(c => c[0])].concat(rows.map(r => cols.map(c => cellValue(c[1](r), c[2] === false))));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = cols.map((c, i) => {
      let w = c[0].length;
      for (let r = 1; r < Math.min(aoa.length, 300); r++) { const v = aoa[r][i]; if (v != null) w = Math.max(w, String(v).length); }
      return { wch: Math.min(w + 2, 50) };
    });
    if (rows.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: cols.length - 1 } }) };
    return ws;
  }
  function noticeSheet(XLSX, extraRows) {
    const ws = XLSX.utils.aoa_to_sheet([
      ['Handling notice', EXPORT_NOTICE],
      ['Exported by', (TOKEN && TOKEN.username) || ''],
      ['Exported (UTC)', new Date().toISOString()],
      ...extraRows,
    ]);
    ws['!cols'] = [{ wch: 18 }, { wch: 90 }];
    return ws;
  }
  function newWorkbook(XLSX) {
    const wb = XLSX.utils.book_new();
    wb.Props = { Title: 'Critical2TN export', Author: (TOKEN && TOKEN.username) || '' };
    return wb;
  }
  async function withBusy(btn, label, fn) {
    const html = btn.innerHTML;
    btn.disabled = true; btn.textContent = label;
    try { await fn(); }
    catch (err) { console.warn('[Critical2TN] Excel export failed', err); alert('Couldn’t create the Excel file: ' + err.message); }
    finally { btn.innerHTML = html; btn.disabled = false; }
  }

  // ── 1. the asset list, as currently filtered
  function listFilterSummary() {
    const parts = [];
    if (activeFilter !== 'all' && SECTORS[activeFilter]) parts.push('Sector: ' + SECTORS[activeFilter].label);
    if (activeCounty) parts.push('County: ' + activeCountyName);
    if (searchTerm) parts.push('Search: “' + searchTerm + '”');
    return parts.join('; ') || 'None — everything your account can see';
  }
  function exportList(btn) {
    const rows = getFiltered();
    if (!rows.length) { alert('There’s nothing in the list to download.'); return; }
    if (rows.length > LIST_EXPORT_CAP) {
      alert(`That’s ${rows.length.toLocaleString()} records — the limit is ${LIST_EXPORT_CAP.toLocaleString()}. Filter by sector or county first.`);
      return;
    }
    withBusy(btn, 'Preparing…', async () => {
      const XLSX = await loadXlsx(), wb = newWorkbook(XLSX);
      XLSX.utils.book_append_sheet(wb, dataSheet(XLSX, CORE_COLS, rows), 'Critical2TN');
      XLSX.utils.book_append_sheet(wb, noticeSheet(XLSX, [['Records', rows.length], ['Filters', listFilterSummary()]]), 'Notice');
      XLSX.writeFile(wb, `Critical2TN_${dataStamp()}.xlsx`, { compression: true });
    });
  }
  function addListButton() {
    const host = document.querySelector('.header-actions');
    if (!host || document.getElementById('xlsx-btn')) return;
    const b = document.createElement('button');
    b.id = 'xlsx-btn'; b.type = 'button'; b.className = 'filter-btn';
    b.title = 'Download the list as an Excel file';
    b.setAttribute('aria-label', 'Download the list as an Excel file');
    b.innerHTML = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M4 20h16"/></svg>Excel';
    b.addEventListener('click', () => exportList(b));
    host.insertBefore(b, document.getElementById('sort-label'));
  }

  // ── 2. the drawn area (Critical2TN + Tier II in one workbook)
  function exportArea(btn) {
    const sel = exportSelection();
    if (!sel.total || sel.core.length > EXPORT_CAP || sel.t2.length > EXPORT_CAP) return;
    withBusy(btn, 'Preparing…', async () => {
      const XLSX = await loadXlsx(), wb = newWorkbook(XLSX);
      if (sel.core.length) XLSX.utils.book_append_sheet(wb, dataSheet(XLSX, CORE_COLS, sel.core), 'Critical2TN');
      if (sel.t2.length) XLSX.utils.book_append_sheet(wb, dataSheet(XLSX, T2_COLS, sel.t2), 'Tier II Hazmat');
      const area = drawState.areaSqMi < 0.1 ? '< 0.1' : drawState.areaSqMi.toFixed(1);
      XLSX.utils.book_append_sheet(wb, noticeSheet(XLSX, [
        ['Area', `${drawState.note || 'Drawn area'} · ${area} sq mi`],
        ['Critical2TN rows', sel.core.length],
        ['Tier II rows', sel.t2.length],
      ]), 'Notice');
      XLSX.writeFile(wb, `Critical2TN_Area_${dataStamp()}.xlsx`, { compression: true });
    });
  }
  function addAreaButton() {
    const btns = document.querySelector('#export-body .ex-btns');
    if (!btns) return;
    const sel = exportSelection();
    const over = sel.core.length > EXPORT_CAP || sel.t2.length > EXPORT_CAP;
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'ex-sec2';
    b.disabled = !sel.total || over || drawState.t2Status === 'loading';
    b.innerHTML = `Download Excel<small>${drawState.t2Status === 'loading' ? 'waiting for Tier II…' : 'both in one workbook · ' + sel.total.toLocaleString() + ' rows'}</small>`;
    b.addEventListener('click', () => exportArea(b));
    btns.appendChild(b);
  }
  const originalRenderExportSheet = window.renderExportSheet;
  window.renderExportSheet = function () { originalRenderExportSheet(); addAreaButton(); };

  addListButton();
})();
