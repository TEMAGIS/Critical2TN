// ════════════════════════════════════════════════════════════════
//  xlsx-export.js — Export panel + Excel (.xlsx) downloads for Critical2TN
//  --------------------------------------------------------------
//  Load it AFTER the main inline <script> in index.html:
//      <script src="xlsx-export.js"></script>
//  It adds:
//   • an "Export" button in the list header, shown only once the list is
//     narrowed to one county (picked under Filter → County, or a county
//     account that only sees its own county). It opens the same kind of
//     panel as "Draw an area → Export", for the list exactly as it's
//     filtered (sector, county, search): untick sectors, then download
//     Critical2TN as CSV, Tier II as CSV (when a county is picked), or
//     both in one Excel workbook.
//   • a "Download Excel" button in the draw-an-area export panel.
//  Every workbook gets a "Notice" tab with the handling notice, who
//  exported it, when, and which filters/area were used.
//  Excel files are saved the same way as the CSV files (a link click on a
//  generated file), so they work wherever the CSV downloads work,
//  including inside Experience Builder.
// ════════════════════════════════════════════════════════════════
(function () {
  // Excel library: tried in order. cdnjs is the same host the map library already loads from.
  const XLSX_SRCS = [
    'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  ];
  const LIST_EXPORT_CAP = 10000;   // max Critical2TN rows from the list (statewide is ~5,300 today)

  // ── small on-screen message (alert() is blocked inside the Experience Builder frame)
  let toastEl = null, toastTimer = null;
  function toast(html) {
    if (!toastEl) {
      toastEl = document.createElement('div');
      toastEl.setAttribute('role', 'status');
      toastEl.style.cssText = 'position:fixed;left:50%;bottom:max(16px,env(safe-area-inset-bottom));transform:translateX(-50%);z-index:6000;' +
        'width:min(92vw,420px);background:#0f1117;color:#fff;border-radius:12px;padding:12px 14px;box-shadow:0 8px 28px rgba(0,0,0,.35);' +
        "font:500 13px/1.45 'DM Sans',sans-serif;display:none";
      document.body.appendChild(toastEl);
    }
    toastEl.innerHTML = `<div>${html}</div>
      <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:10px">
        <button type="button" data-t="close" style="font:700 13px 'DM Sans',sans-serif;padding:9px 14px;min-height:40px;border-radius:8px;border:1px solid rgba(255,255,255,.4);background:transparent;color:#fff;cursor:pointer">Close</button>
      </div>`;
    toastEl.style.display = 'block';
    toastEl.querySelector('[data-t="close"]').onclick = hideToast;
    clearTimeout(toastTimer); toastTimer = setTimeout(hideToast, 10000);
  }
  function hideToast() { if (toastEl) toastEl.style.display = 'none'; }

  // ── load SheetJS on demand
  let xlsxPromise = null;
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src; s.async = true;
      s.onload = () => window.XLSX ? resolve(window.XLSX) : reject(new Error('loaded but empty'));
      s.onerror = () => { s.remove(); reject(new Error('blocked or unreachable')); };
      document.head.appendChild(s);
    });
  }
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (xlsxPromise) return xlsxPromise;
    xlsxPromise = (async () => {
      for (const src of XLSX_SRCS) {
        try { return await loadScript(src); }
        catch (e) { console.warn('[Critical2TN] Excel library not loaded from', src, e.message); }
      }
      xlsxPromise = null;
      throw new Error('couldn’t load the Excel library — your network may block it. The CSV downloads still work.');
    })();
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
    catch (err) { console.warn('[Critical2TN] export failed', err); toast('Couldn’t create the file: ' + escapeHtml(err.message)); }
    finally { btn.innerHTML = html; btn.disabled = false; }
  }
  // Same columns and safety rules as the area CSV export in index.html.
  function csvText(cols, rows) {
    const now = new Date().toISOString(), who = (TOKEN && TOKEN.username) || '';
    const head = cols.map(c => c[0]).concat(['exported_by', 'exported_utc', 'handling_notice']).join(',');
    const lines = rows.map(r => cols.map(c => csvCell(c[1](r), c[2] !== false)).concat([csvCell(who), csvCell(now), csvCell(EXPORT_NOTICE)]).join(','));
    return '\uFEFF' + [head, ...lines].join('\r\n') + '\r\n';
  }
  function saveCsv(name, text) {
    const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' })), a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }
  function saveXlsx(XLSX, wb, name) {
    const data = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
    const url = URL.createObjectURL(new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })), a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }
  const plural = (n, one, many) => `${n.toLocaleString()} ${n === 1 ? one : many}`;

  // ════════════ 1. EXPORT PANEL FOR THE LIST ════════════
  const listX = { open: false, off: new Set(), t2: { key: null, status: 'idle', rows: [], ctl: null } };
  function listFilterSummary() {
    const parts = [];
    if (activeFilter !== 'all' && SECTORS[activeFilter]) parts.push('Sector: ' + SECTORS[activeFilter].label);
    if (activeCounty) parts.push('County: ' + activeCountyName);
    if (searchTerm) parts.push('Search: “' + searchTerm + '”');
    return parts.join('; ') || 'None — everything your account can see';
  }
  const listRows = () => getFiltered().filter(a => !listX.off.has(sectorOf(a).key));
  const listT2Rows = () => listX.t2.status === 'ok' && !listX.t2Off ? listX.t2.rows : [];

  // Tier II for the list: only when the list is narrowed to a county (picked in Filter,
  // or the user's own county). Statewide Tier II is far too big for one file.
  async function loadListT2() {
    const where = refCountyWhere(REF_BY_KEY.tier2);
    if (!where) { if (listX.t2.ctl) listX.t2.ctl.abort(); listX.t2 = { key: '', status: 'nocounty', rows: [], ctl: null }; if (listX.open) renderListSheet(); return; }
    if (listX.t2.key === where && (listX.t2.status === 'ok' || listX.t2.status === 'loading')) return;
    if (listX.t2.ctl) listX.t2.ctl.abort();
    const ctl = new AbortController();
    listX.t2 = { key: where, status: 'loading', rows: [], ctl };
    if (listX.open) renderListSheet();
    try {
      const url = REF_BY_KEY.tier2.url + '/query';
      const base = { where, returnGeometry: 'true', outSR: '4326', f: 'json', orderByFields: 'OBJECTID' };
      let fields = T2_FIELDS.join(','), out = [], offset = 0, more = true, guard = 12;
      while (more && guard-- > 0 && out.length <= EXPORT_CAP) {
        let json;
        try { json = await refPost(url, { ...base, outFields: fields, resultOffset: String(offset), resultRecordCount: '1000' }, ctl.signal); }
        catch (err) {                                        // a field name the service doesn't have → retry once with all fields
          if (fields !== '*' && err.code === 400) { fields = '*'; continue; }
          throw err;
        }
        const f = json.features || []; out.push(...f); offset += f.length;
        more = json.exceededTransferLimit === true && f.length > 0;
      }
      if (listX.t2.ctl !== ctl) return;
      listX.t2 = { key: where, status: 'ok', rows: out, ctl: null };
    } catch (err) {
      if (listX.t2.ctl !== ctl) return;
      const denied = err.code === 403 || err.code === 498 || err.code === 499 || /permission|not authorized|token/i.test(err.message || '');
      console.warn('[Critical2TN] Tier II list export query failed', err);
      listX.t2 = { key: where, status: denied ? 'unavailable' : 'error', rows: [], ctl: null };
    }
    if (listX.open) renderListSheet();
  }

  function renderListSheet() {
    const body = document.getElementById('export-body'); if (!body) return;
    const all = getFiltered(), core = listRows(), t2 = listT2Rows(), st = listX.t2.status;
    const overC = core.length > LIST_EXPORT_CAP, overT = t2.length > EXPORT_CAP;
    const bySector = {}; all.forEach(a => { const k = sectorOf(a).key; bySector[k] = (bySector[k] || 0) + 1; });
    const t2Fac = new Set(listX.t2.rows.map(f => { const p = f.attributes || {}; return p.FacilityId || p.State_Facility_ID || p.FacilityName; })).size;
    const t2Line = {
      nocounty: 'Pick a county (Filter → County) or draw an area on the map to include Tier II.',
      loading: 'Checking Tier II…', unavailable: 'Not available to your account', error: 'Couldn’t reach Tier II',
      ok: `${plural(t2Fac, 'facility', 'facilities')} · ${plural(listX.t2.rows.length, 'chemical record', 'chemical records')}`,
    }[st] || '';
    const secRows = Object.keys(bySector).sort((a, b) => bySector[b] - bySector[a]).map(k => {
      const s = SECTORS[k] || UNKNOWN_SECTOR;
      return `<label class="ex-sec"><input type="checkbox" data-sec="${k}" ${listX.off.has(k) ? '' : 'checked'}>${iconHTML(k, 20)}<span>${escapeHtml(s.label)}</span><b>${bySector[k]}</b></label>`;
    }).join('');
    const total = core.length + t2.length;
    body.innerHTML = `
      <h2 id="export-title" class="rl-title">Export list</h2>
      <div class="rl-sub">Filters: ${escapeHtml(listFilterSummary())} · only data your account can already see</div>
      <div class="ex-card">
        <div class="ex-head"><span>Critical2TN assets</span><b>${all.length.toLocaleString()}</b></div>
        ${secRows || '<div class="ex-none">Nothing in the list</div>'}
      </div>
      <div class="ex-card ${st === 'unavailable' || st === 'nocounty' ? 'dis' : ''}">
        <div class="ex-head"><span>${st === 'ok' ? `<input type="checkbox" data-t2 ${listX.t2Off ? '' : 'checked'}>` : ''}Tier II HAZMAT</span></div>
        <div class="ex-none" data-s="${st}">${escapeHtml(t2Line)}${st === 'error' ? ' <button class="ex-link" data-retry>Retry</button>' : ''}</div>
      </div>
      ${overC ? `<div class="ex-total over">That’s over the ${LIST_EXPORT_CAP.toLocaleString()}-row limit. Filter by sector or county, or untick sectors.</div>` : ''}
      ${overT ? `<div class="ex-total over">Tier II is over the ${EXPORT_CAP.toLocaleString()}-row limit. Draw a smaller area on the map instead.</div>` : ''}
      <div class="ex-btns ex-two">
        <button class="ex-pri" type="button" data-x="csv-core" ${!core.length || overC ? 'disabled' : ''}>Download Critical2TN (CSV)<small>${plural(core.length, 'row', 'rows')}</small></button>
        <button class="ex-pri" type="button" data-x="csv-t2" ${!t2.length || overT ? 'disabled' : ''}>Download Tier II Hazmat (CSV)<small>${st === 'loading' ? 'checking…' : plural(t2.length, 'row', 'rows')}</small></button>
        <button class="ex-sec2" type="button" data-x="xlsx" ${!total || overC || overT || st === 'loading' ? 'disabled' : ''}>Download Excel<small>${st === 'loading' ? 'waiting for Tier II…' : (t2.length ? 'both in one workbook · ' : '') + plural(total, 'row', 'rows')}</small></button>
      </div>
      <div class="ex-note">${escapeHtml(EXPORT_NOTICE)}</div>`;
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
  // Show Export only for a single county: one picked in the filter, or a county
  // account whose list is already just that county (they have no county picker).
  function exportCountyChosen() {
    if (activeCounty) return true;
    return typeof USER_SCOPE !== 'undefined' && USER_SCOPE.type === 'county' && USER_SCOPE.counties.length === 1;
  }
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
  }

  // One panel (#export-sheet) serves both exports; listX.open says which one is showing.
  const originalRender = window.renderExportSheet, originalOpen = window.openExportSheet, originalClose = window.closeExportSheet;
  window.renderExportSheet = function () { if (listX.open) return renderListSheet(); originalRender(); addAreaButton(); };
  window.openExportSheet = function () { listX.open = false; originalOpen(); };
  window.closeExportSheet = function (silent) { originalClose(silent); listX.open = false; };

  addListButton();
})();
