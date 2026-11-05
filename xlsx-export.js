      <div class="ex-note">${escapeHtml(EXPORT_NOTICE)}<br><span style="opacity:.7">Export v${VERSION}</span></div>`;
    body.querySelectorAll('[data-sec]').forEach(cb => cb.addEventListener('change', () => {
      cb.checked ? listX.off.delete(cb.dataset.sec) : listX.off.add(cb.dataset.sec); renderListSheet();
    }));
    const t2cb = body.querySelector('[data-t2]'); if (t2cb) t2cb.addEventListener('change', () => { listX.t2Off = !t2cb.checked; renderListSheet(); });
    const retry = body.querySelector('[data-retry]'); if (retry) retry.addEventListener('click', () => { listX.t2.key = null; loadListT2(); });
    const btns = body.querySelector('.ex-btns');
    btns.querySelector('[data-x="csv-core"]').addEventListener('click', () => saveCsv(`Critical2TN_${dataStamp()}.csv`, csvText(CORE_COLS, listRows())));
    btns.querySelector('[data-x="csv-t2"]').addEventListener('click', () => saveCsv(`TierII_Hazmat_${dataStamp()}.csv`, csvText(T2_COLS, listT2Rows())));
    const xb = btns.querySelector('[data-x="xlsx"]');
    xb.addEventListener('click', () => withBusy(xb, 'Preparing…', async () => {
      const rows = listRows(), t2rows = listT2Rows(), XLSX = await loadXlsx(), wb = newWorkbook(XLSX);
      if (rows.length) XLSX.utils.book_append_sheet(wb, dataSheet(XLSX, CORE_COLS, rows), 'Critical2TN');
      if (t2rows.length) XLSX.utils.book_append_sheet(wb, dataSheet(XLSX, T2_COLS, t2rows), 'Tier II Hazmat');
      XLSX.utils.book_append_sheet(wb, noticeSheet(XLSX, [
        ['Filters', listFilterSummary()],
        ['Sectors left out', [...listX.off].map(k => (SECTORS[k] || UNKNOWN_SECTOR).label).join(', ') || 'None'],
        ['Critical2TN rows', rows.length],
        ['Tier II rows', t2rows.length],
      ]), 'Notice');
      saveXlsx(XLSX, wb, `Critical2TN_${dataStamp()}.xlsx`);
    }));
  }

  function openListExport() {
    const sheet = document.getElementById('export-sheet'); if (!sheet) return;
    listX.open = true; listX.off = new Set(); listX.t2Off = false;
    renderListSheet();
    sheet.removeAttribute('hidden');
    requestAnimationFrame(() => sheet.classList.add('open'));
    document.addEventListener('keydown', _exportKey);
    setTimeout(() => { const b = sheet.querySelector('.sheet-close-btn'); if (b) b.focus(); }, 60);
    loadListT2();
  }
  function addListButton() {
    const host = document.querySelector('.header-actions');
    if (!host || document.getElementById('xlsx-btn')) return;
    const b = document.createElement('button');
    b.id = 'xlsx-btn'; b.type = 'button'; b.className = 'filter-btn';
    b.title = 'Export the list (CSV or Excel)';
    b.setAttribute('aria-label', b.title);
    b.setAttribute('aria-haspopup', 'dialog');
    b.innerHTML = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M4 20h16"/></svg>Export';
    b.addEventListener('click', openListExport);
    host.insertBefore(b, document.getElementById('sort-label'));
    updateExportButton();
  }
  // Show Export after a county is picked under Filter → County, or straight away
  // for an account that only sees one county (its list is already that county).
  function exportCountyChosen() { return !!activeCounty || !!singleCountyAccount(); }
  function updateExportButton() {
    const b = document.getElementById('xlsx-btn'); if (!b) return;
    b.style.display = exportCountyChosen() ? '' : 'none';
  }
  // renderList runs after a county is picked or cleared and after the data loads.
  const originalRenderList = window.renderList;
  window.renderList = function () { originalRenderList.apply(this, arguments); updateExportButton(); };

  // ════════════ 2. DRAWN AREA: add Excel to the existing panel ════════════
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
      saveXlsx(XLSX, wb, `Critical2TN_Area_${dataStamp()}.xlsx`);
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
    const note = document.querySelector('#export-body .ex-note');
    if (note && !note.querySelector('[data-ver]')) note.insertAdjacentHTML('beforeend', `<br><span data-ver style="opacity:.7">Export v${VERSION}</span>`);
  }

  // One panel (#export-sheet) serves both exports; listX.open says which one is showing.
  const originalRender = window.renderExportSheet, originalOpen = window.openExportSheet, originalClose = window.closeExportSheet;
  window.renderExportSheet = function () { if (listX.open) return renderListSheet(); originalRender(); addAreaButton(); };
  window.openExportSheet = function () { listX.open = false; originalOpen(); };
  window.closeExportSheet = function (silent) { originalClose(silent); listX.open = false; };

  addListButton();
})();