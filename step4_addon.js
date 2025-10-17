// 3seats Step-4 Addon – "Existing Rooms and Owners in Tripleseat Upload"
// - UI: inserts a new Step 4 card after Step 3 with a file picker
// - Reads an uploaded Excel (XLSX) in the browser (SheetJS)
// - Extracts Rooms and Owners/Users columns by header name heuristics
// - Injects into final Excel's "Lists" sheet: Rooms -> column B, Owners -> column E (from row 2)
// - Works automatically when user clicks your existing "Generate" (patches ExcelJS writeBuffer)

(function(){
  if (window.__TS_STEP4_WIRED__) return;
  window.__TS_STEP4_WIRED__ = true;

  // ---- minimal loader for SheetJS (XLSX) to read Excel client-side ----
  function loadScriptOnce(src){
    return new Promise((resolve,reject)=>{
      if ([...document.scripts].some(s => (s.src||"").includes(src))) return resolve();
      const el = document.createElement('script');
      el.src = src; el.async = true; el.defer = true;
      el.onload = resolve;
      el.onerror = () => reject(new Error("Failed to load: "+src));
      document.head.appendChild(el);
    });
  }
  async function ensureSheetJS(){
    // Pinned, small, stable build
    await loadScriptOnce("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
    if (!window.XLSX) throw new Error("SheetJS failed to load");
  }

  // ---- Step placement helpers ----
  function findStep3Card(){
    const hs = Array.from(document.querySelectorAll('h1,h2,h3,h4,strong,b,.title,.header'));
    for (const h of hs){
      const t = (h.textContent||'').trim().toLowerCase();
      // Find a heading that starts with "3)" (your existing Step 3)
      if (/^3[\).\s]/.test(t)) {
        return h.closest('section, .step, .card, .panel, .box, .container, .chunk, .ts-card, div') || h.parentElement;
      }
    }
    // fallback: top container
    return document.querySelector('.container, .step, .card, section, body > div') || document.body;
  }

  // Build Step 4 card by cloning Step 3's outer wrapper so styling matches
  function buildStep4From(step3){
    const step4 = step3 ? step3.cloneNode(false) : document.createElement('div');
    step4.id = 'ts-step4-card';

    const h3 = step3 && step3.querySelector('h1,h2,h3,h4,.title,.header,strong,b');
    const content3 =
      (step3 && Array.from(step3.children).find(el => el !== h3 && el.tagName && el.tagName.toLowerCase() !== 'script')) ||
      (step3 && step3.querySelector('.content, .body, .card-content, .section-body, .panel-body, .p-*, .px-*, .py-*'));

    const HeadTag = h3 ? (h3.tagName || 'div').toLowerCase() : 'h3';
    const head = document.createElement(HeadTag);
    head.textContent = '4) Existing Rooms and Owners in Tripleseat Upload';
    if (h3 && h3.classList.length) h3.classList.forEach(c => head.classList.add(c));
    else head.style.fontWeight = '600';

    const body = document.createElement(content3 ? content3.tagName.toLowerCase() : 'div');
    if (content3 && content3.classList.length) content3.classList.forEach(c => body.classList.add(c));

    const desc = document.createElement('div');
    const helper = step3 && step3.querySelector('.help,.hint,.description,small,.text-muted');
    if (helper && helper.classList.length) helper.classList.forEach(c => desc.classList.add(c));
    desc.innerHTML = 'Upload the Business Information Spreadsheet or other Excel document that lists the current <b>Rooms</b> and <b>Owners</b> loaded into Tripleseat. We\'ll pull those into your final Excel on the <b>Lists</b> tab.';

    // Controls row: try to mimic Step 3 row classes
    const row = document.createElement('div');
    const row3 = step3 && step3.querySelector('.controls, .row, .fields, .inline-controls, .grid, .flex');
    if (row3 && row3.classList.length) row3.classList.forEach(c => row.classList.add(c));
    else { row.style.display='flex'; row.style.flexWrap='wrap'; row.style.alignItems='center'; row.style.gap='10px'; }

    // File input styled like Step 3's inputs
    const file = document.createElement('input');
    file.type = 'file';
    file.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    file.id = 'ts-step4-file';
    const refInput = step3 && step3.querySelector('input[type="text"], .input, input, select, textarea');
    if (refInput){
      if (refInput.classList.length) refInput.classList.forEach(c => file.classList.add(c));
      const cs = getComputedStyle(refInput);
      file.style.padding = cs.padding;
      file.style.border = cs.border;
      file.style.borderRadius = cs.borderRadius;
      file.style.backgroundColor = cs.backgroundColor;
      file.style.color = cs.color;
    }

    const status = document.createElement('span');
    status.id = 'ts-step4-status';
    status.textContent = 'No file selected';
    if (helper && helper.classList.length) helper.classList.forEach(c => status.classList.add(c));

    row.appendChild(file);
    row.appendChild(status);
    body.appendChild(desc);
    body.appendChild(row);
    step4.appendChild(head);
    step4.appendChild(body);
    return step4;
  }

  // ---- Extractors ----
  function normalize(s){ return String(s||'').trim(); }
  function pickHeaderIndex(headers, candidates){
    // headers: array of strings; candidates: array of substrings to match (lowercased)
    const lc = headers.map(h => String(h||'').toLowerCase().trim());
    for (let i=0; i<lc.length; i++){
      const h = lc[i];
      for (const c of candidates){ if (h.includes(c)) return i; }
    }
    return -1;
  }
  function uniqueNonEmpty(arr){
    const seen = new Set(); const out = [];
    for (const v of arr.map(normalize)){
      if (!v) continue;
      const key = v.toLowerCase();
      if (!seen.has(key)){ seen.add(key); out.push(v); }
    }
    return out;
  }

  async function parseRoomsOwnersFromExcel(file){
    await ensureSheetJS();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type:'array' });

    // Heuristic: prefer the sheet named like "Business", else the first sheet
    const sheetNames = wb.SheetNames || [];
    const namePref = sheetNames.find(n => /business|info|information/i.test(n)) || sheetNames[0];
    const ws = wb.Sheets[namePref];
    if (!ws) throw new Error("No readable sheet found in uploaded file.");

    const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }); // 2D array
    if (!json.length) throw new Error("Uploaded sheet is empty.");

    const headers = (json[0] || []).map(cell => String(cell||'').trim());
    // Try several common header variants
    const roomIdx  = pickHeaderIndex(headers, ['room', 'location', 'space']);
    const ownerIdx = pickHeaderIndex(headers, ['owner', 'user']);

    const rows = json.slice(1);
    const rooms  = roomIdx  >= 0 ? rows.map(r => r[roomIdx])  : [];
    const owners = ownerIdx >= 0 ? rows.map(r => r[ownerIdx]) : [];

    return {
      rooms:  uniqueNonEmpty(rooms),
      owners: uniqueNonEmpty(owners),
      sheetUsed: namePref,
      foundRoomHeader: roomIdx >= 0 ? headers[roomIdx] : null,
      foundOwnerHeader: ownerIdx >= 0 ? headers[ownerIdx]: null
    };
  }

  // ---- Injection into final Excel via ExcelJS patch ----
  function applyListsToWorkbook(workbook, lists){
    if (!workbook || !lists) return;
    const wsName = 'Lists';
    let ws = workbook.getWorksheet(wsName);
    if (!ws) ws = workbook.addWorksheet(wsName);

    // Ensure headers exist; do not overwrite if you already set them elsewhere
    // B1 header commonly "unmatched_rooms" (we won't change)
    // E1 header commonly "unmatched_owners" (we won't change)

    // Clear existing columns below header
    const maxRows = Math.max(ws.rowCount || 1000, 1000);
    // Column B (2)
    for (let r = 2; r <= maxRows; r++){
      const cell = ws.getCell(r, 2); // B
      cell.value = null;
    }
    // Column E (5)
    for (let r = 2; r <= maxRows; r++){
      const cell = ws.getCell(r, 5); // E
      cell.value = null;
    }

    // Write rooms to column B starting B2
    let rIndex = 2;
    (lists.rooms || []).forEach(v => {
      ws.getCell(rIndex++, 2).value = v;
    });

    // Write owners to column E starting E2
    rIndex = 2;
    (lists.owners || []).forEach(v => {
      ws.getCell(rIndex++, 5).value = v;
    });
  }

  function patchExcelJSOnce(){
    if (window.__TS_STEP4_PATCHED__) return;
    window.__TS_STEP4_PATCHED__ = true;

    const ExcelJS = window.ExcelJS;
    if (!ExcelJS || !ExcelJS.Workbook || !ExcelJS.Workbook.prototype || !ExcelJS.Workbook.prototype.xlsx){
      // Retry shortly if ExcelJS not ready yet
      setTimeout(patchExcelJSOnce, 500);
      return;
    }

    const origWriteBuffer = ExcelJS.Workbook.prototype.xlsx.writeBuffer;
    if (!origWriteBuffer) return;

    ExcelJS.Workbook.prototype.xlsx.writeBuffer = function(){
      try{
        if (window.__TS_STEP4_LISTS && (window.__TS_STEP4_LISTS.rooms?.length || window.__TS_STEP4_LISTS.owners?.length)){
          applyListsToWorkbook(this, window.__TS_STEP4_LISTS);
        }
      }catch(e){
        console.warn('[Step4] Injection skipped:', e);
      }
      return origWriteBuffer.apply(this, arguments);
    };
    console.log('[Step4] ExcelJS writeBuffer patched for Rooms/Owners injection');
  }

  // ---- Wire UI + behavior ----
  function wireUI(){
    const step3 = findStep3Card();
    const step4 = buildStep4From(step3);
    // Insert *after* Step 3 so it reads as "4)"
    if (step3 && step3.parentElement) step3.parentElement.insertBefore(step4, step3.nextSibling);
    else document.body.appendChild(step4);

    // Patch Excel export so your existing "Generate" auto-injects lists
    patchExcelJSOnce();

    // Handle file selection
    const file = document.getElementById('ts-step4-file');
    const status = document.getElementById('ts-step4-status');
    file?.addEventListener('change', async ()=>{
      const f = file.files && file.files[0];
      if (!f){ status.textContent = 'No file selected'; return; }
      status.textContent = 'Reading…';
      try{
        const parsed = await parseRoomsOwnersFromExcel(f);
        window.__TS_STEP4_LISTS = { rooms: parsed.rooms, owners: parsed.owners };
        status.textContent = `Loaded ${parsed.rooms.length} rooms & ${parsed.owners.length} owners from “${parsed.sheetUsed}”`;
      }catch(e){
        console.error(e);
        status.textContent = 'Could not read that file. Please pick a valid Excel.';
        alert('Step 4 parse failed: ' + (e?.message || e));
      }
    });
  }

  document.addEventListener('DOMContentLoaded', wireUI);
})();
