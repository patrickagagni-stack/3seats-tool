// 3seats Step-4 Addon v5 – “Existing Rooms and Owners in Tripleseat Upload”
// - Finds the correct sheet (ignores EXAMPLE), extracts Rooms/Owners from col A headers
// - Injects into Lists!B2 (Rooms) and Lists!E2 (Owners) during export
// - ULTRA DEFENSIVE: never assumes shapes, avoids null/undefined paths that can trigger .slice

(function(){
  if (window.__TS_STEP4_WIRED_V5__) return;
  window.__TS_STEP4_WIRED_V5__ = true;

  // ---------- load SheetJS (for parsing uploaded spreadsheet) ----------
  function loadScriptOnce(src){
    return new Promise((resolve,reject)=>{
      if ([...document.scripts].some(s => (s.src||"").includes(src))) return resolve();
      const el = document.createElement('script');
      el.src = src; el.async = true; el.defer = true;
      el.onload = resolve; el.onerror = () => reject(new Error("Failed to load: "+src));
      document.head.appendChild(el);
    });
  }
  async function ensureSheetJS(){
    await loadScriptOnce("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
    if (!window.XLSX) throw new Error("SheetJS failed to load");
  }

  // ---------- place Step 4 card visually after current Step 3 ----------
  function findStep3Card(){
    const hs = Array.from(document.querySelectorAll('h1,h2,h3,h4,strong,b,.title,.header'));
    for (const h of hs){
      const t = (h.textContent||'').trim().toLowerCase();
      if (/^3[\).\s]/.test(t)) {
        return h.closest('section, .step, .card, .panel, .box, .container, .chunk, .ts-card, div') || h.parentElement;
      }
    }
    return document.querySelector('.container, .step, .card, section, body > div') || document.body;
  }

  function buildStep4From(step3){
    const step4 = step3 ? step3.cloneNode(false) : document.createElement('div');
    step4.id = 'ts-step4-card';

    const h3 = step3 && step3.querySelector('h1,h2,h3,h4,.title,.header,strong,b');
    const content3 =
      (step3 && Array.from(step3.children).find(el => el !== h3 && el.tagName && el.tagName.toLowerCase() !== 'script')) ||
      (step3 && step3.querySelector('.content, .body, .card-content, .section-body, .panel-body'));

    const HeadTag = h3 ? (h3.tagName || 'div').toLowerCase() : 'h3';
    const head = document.createElement(HeadTag);
    head.textContent = '4) Existing Rooms and Owners in Tripleseat Upload';
    if (h3 && h3.classList.length) h3.classList.forEach(c => head.classList.add(c));
    else head.style.fontWeight = '600';

    const body = document.createElement(content3 ? content3.tagName.toLowerCase() : 'div');
    if (content3 && content3.classList.length) content3.classList.forEach(c => body.classList.add(c));

    const desc = document.createElement('div');
    desc.innerHTML = 'Upload the Business Information Spreadsheet or other Excel document that lists the current <b>Rooms</b> and <b>Owners</b> loaded into Tripleseat. We\'ll pull those into your final Excel on the <b>Lists</b> tab.';

    const row = document.createElement('div');
    row.style.display='flex'; row.style.flexWrap='wrap'; row.style.alignItems='center'; row.style.gap='10px';

    const file = document.createElement('input');
    file.type = 'file';
    file.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    file.id = 'ts-step4-file';
    const refInput = step3 && step3.querySelector('input[type="text"], .input, input, select, textarea');
    if (refInput){
      if (refInput.classList.length) refInput.classList.forEach(c => file.classList.add(c));
      const cs = getComputedStyle(refInput);
      file.style.padding = cs.padding; file.style.border = cs.border;
      file.style.borderRadius = cs.borderRadius;
      file.style.backgroundColor = cs.backgroundColor; file.style.color = cs.color;
    }

    const status = document.createElement('span');
    status.id = 'ts-step4-status';
    status.textContent = 'No file selected';
    row.appendChild(file); row.appendChild(status);
    body.appendChild(desc); body.appendChild(row);
    step4.appendChild(head); step4.appendChild(body);
    return step4;
  }

  // ---------- parsing helpers ----------
  function safeRowCount(ws){ return (ws && typeof ws.rowCount === 'number' && ws.rowCount > 0) ? ws.rowCount : 5000; }
  function getCellValue(ws,r,c){
    try{
      const cell = ws && typeof ws.getCell === 'function' ? ws.getCell(r,c) : null;
      const v = cell ? (cell.value ?? '') : '';
      return (v === null || v === undefined) ? '' : String(v);
    }catch{ return ''; }
  }

  function find_header_in_col_a(ws, headers){
    const targets = headers.map(h => String(h).toLowerCase().trim());
    const max = safeRowCount(ws);
    for (let r=1; r<=max; r++){
      const s = getCellValue(ws,r,1).toLowerCase().trim();
      if (!s) continue;
      if (targets.includes(s)) return r;
    }
    return null;
  }

  function extract_list_below(ws, headerRow, skipRows=1){
    const out=[]; let blanks=0;
    const start = headerRow + 1 + skipRows;
    const end = Math.max(start + 50, start + 5000); // bounded sweep
    for (let r=start; r<=end; r++){
      const s = getCellValue(ws,r,1).trim();
      if (!s){
        blanks++; if (blanks>=3) break;
      } else {
        blanks=0;
        const low = s.toLowerCase();
        if (!/^enter\s|^type\s|^do not/i.test(low)) out.push(s);
      }
    }
    // de-dupe preserve order
    const seen=new Set(), dedup=[];
    for(const x of out){const k=x.toLowerCase(); if(!seen.has(k)){seen.add(k);dedup.push(x);} }
    return dedup;
  }

  function chooseBestSheet(wb){
    const badName=/(example|instruction|cover|template)/i;
    const goodName=/(your\s+business|business)/i;
    let best=null,bestScore=-Infinity;
    const sheets = Array.isArray(wb?.worksheets) ? wb.worksheets : [];
    for(const ws of sheets){
      const name=(ws.name||'').trim();
      let score=0;
      if(goodName.test(name))score+=5;
      if(badName.test(name))score-=5;
      const rHdr=find_header_in_col_a(ws,['event space']);
      const uHdr=find_header_in_col_a(ws,['full name']);
      if(rHdr)score+=3;if(uHdr)score+=3;
      if(rHdr)score+=extract_list_below(ws,rHdr,1).length;
      if(uHdr)score+=extract_list_below(ws,uHdr,1).length;
      if(score>bestScore){bestScore=score;best={ws,rHdr,uHdr,name};}
    }
    return best;
  }

  async function parseRoomsOwnersFromExcel(file){
    await ensureSheetJS();
    const buf=await file.arrayBuffer();
    const wbRaw=XLSX.read(buf,{type:'array'});
    const workbook={worksheets:wbRaw.SheetNames.map(n=>{
      const grid=XLSX.utils.sheet_to_json(wbRaw.Sheets[n],{header:1,defval:''});
      return {
        name:n,
        rowCount:grid.length,
        getCell:(r,c)=>({ value:(grid[r-1] && grid[r-1][c-1]) ?? '' })
      };
    })};
    const pick=chooseBestSheet(workbook);
    if(!pick)throw new Error("No suitable sheet found");
    let rooms=[],owners=[];
    if(pick.rHdr)rooms=extract_list_below(pick.ws,pick.rHdr,1);
    if(pick.uHdr)owners=extract_list_below(pick.ws,pick.uHdr,1);
    return {rooms,owners,sheetUsed:pick.name,foundRoomHeader:!!pick.rHdr,foundOwnerHeader:!!pick.uHdr};
  }

  // ---------- ExcelJS injection (ultra-safe) ----------
  function writeColumnValuesSafe(ws, colIndex, startRow, values){
    if (!ws || !Array.isArray(values) || values.length === 0) return;
    let r = startRow;
    for (const v of values){
      try {
        const cell = ws.getCell(r++, colIndex);
        cell.value = (v == null ? '' : v); // use '' not null
      } catch(e) {
        // keep going; don't abort export
        console.warn('[Step4] write cell failed', e);
      }
    }
  }
  function clearColumnRangeSafe(ws, colIndex, startRow, endRow){
    if (!ws) return;
    for (let r=startRow; r<=endRow; r++){
      try {
        ws.getCell(r, colIndex).value = '';
      } catch(e) { /* ignore */ }
    }
  }
  function isWorksheet(obj){
    return obj && typeof obj.getCell === 'function' && typeof obj.getRow === 'function';
  }
  function applyListsToWorkbook(workbook, lists){
    try{
      if (!workbook || !lists) return;
      const getWS = (name)=>{
        try {
          return (typeof workbook.getWorksheet === 'function') ? workbook.getWorksheet(name) : null;
        }catch{ return null; }
      };
      const addWS = (name)=>{
        try {
          return (typeof workbook.addWorksheet === 'function') ? workbook.addWorksheet(name) : null;
        }catch{ return null; }
      };

      let ws = getWS('Lists');
      if (!isWorksheet(ws)) ws = addWS('Lists');
      if (!isWorksheet(ws)) return; // silently bail if cannot obtain a worksheet

      const rooms = Array.isArray(lists.rooms) ? lists.rooms : [];
      const owners = Array.isArray(lists.owners) ? lists.owners : [];
      if (rooms.length === 0 && owners.length === 0) return;

      const clearRows = Math.max(rooms.length, owners.length, 100);
      clearColumnRangeSafe(ws, 2, 2, 1 + clearRows); // B2..B{clear}
      clearColumnRangeSafe(ws, 5, 2, 2 + clearRows); // E2..E{clear}

      writeColumnValuesSafe(ws, 2, 2, rooms);
      writeColumnValuesSafe(ws, 5, 2, owners);
    }catch(e){
      console.warn('[Step4] applyListsToWorkbook skipped:', e);
    }
  }

  function patchExcelJSOnce(){
    if (window.__TS_STEP4_PATCHED_V5__) return;
    window.__TS_STEP4_PATCHED_V5__ = true;

    // Wait until ExcelJS exists
    const tryPatch = ()=>{
      const ExcelJS = window.ExcelJS;
      if (!ExcelJS || !ExcelJS.Workbook || !ExcelJS.Workbook.prototype || !ExcelJS.Workbook.prototype.xlsx){
        setTimeout(tryPatch, 500);
        return;
      }
      const proto = ExcelJS.Workbook.prototype;
      const origWriteBuffer = proto.xlsx.writeBuffer;

      if (typeof origWriteBuffer !== 'function'){
        // nothing to patch
        return;
      }

      // Prevent double-patching
      if (proto.__ts_step4_wrapped__) return;
      proto.__ts_step4_wrapped__ = true;

      proto.xlsx.writeBuffer = function(){
        try{
          const lists = window.__TS_STEP4_LISTS;
          if (lists && (Array.isArray(lists.rooms) || Array.isArray(lists.owners))){
            applyListsToWorkbook(this, lists);
          }
        }catch(e){
          console.warn('[Step4] writeBuffer hook skipped:', e);
        }
        return origWriteBuffer.apply(this, arguments);
      };
      console.log('[Step4] Patched writeBuffer (v5)');
    };
    tryPatch();
  }

  // ---------- Wire UI ----------
  function wireUI(){
    const step3=findStep3Card();
    const step4=buildStep4From(step3);
    if(step3&&step3.parentElement)step3.parentElement.insertBefore(step4,step3.nextSibling);
    else document.body.appendChild(step4);

    // Patch export
    patchExcelJSOnce();

    const file=document.getElementById('ts-step4-file');
    const status=document.getElementById('ts-step4-status');
    file?.addEventListener('change',async()=>{
      const f=file.files&&file.files[0];
      if(!f){status.textContent='No file selected';return;}
      status.textContent='Reading…';
      try{
        const parsed=await parseRoomsOwnersFromExcel(f);
        window.__TS_STEP4_LISTS={rooms:parsed.rooms,owners:parsed.owners};
        status.textContent=`Loaded ${parsed.rooms.length} rooms & ${parsed.owners.length} owners from “${parsed.sheetUsed}”`;
      }catch(e){
        console.error(e);
        status.textContent='Could not read that file.';
        alert('Step 4 parse failed: '+(e?.message||e));
      }
    });
  }

  document.addEventListener('DOMContentLoaded', wireUI);
})();
