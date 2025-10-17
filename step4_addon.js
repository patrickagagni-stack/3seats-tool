// 3seats Step-4 Addon – “Existing Rooms and Owners in Tripleseat Upload”
// Reads an uploaded Excel, finds the proper sheet (ignoring EXAMPLE),
// extracts Rooms (under “Event Space”) and Owners (under “Full Name”),
// and injects them into Lists!B2 (Rooms) and Lists!E2 (Owners) in the final Excel.

(function(){
  if (window.__TS_STEP4_WIRED__) return;
  window.__TS_STEP4_WIRED__ = true;

  // ---------- load SheetJS ----------
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

  // ---------- UI placement ----------
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

  // ---------- helpers for parsing ----------
  function find_header_in_col_a(ws, headers){
    const targets = headers.map(h => String(h).toLowerCase().trim());
    const max = ws.rowCount || 5000;
    for (let r=1;r<=max;r++){
      const v = ws.getCell(r,1).value;
      const s = String(v||'').toLowerCase().trim();
      if (targets.includes(s)) return r;
    }
    return null;
  }

  function extract_list_below(ws, headerRow, skipRows=1){
    const out=[]; let blanks=0;
    for (let r=headerRow+1+skipRows; r<=ws.rowCount; r++){
      const v=ws.getCell(r,1).value;
      const s=String(v||'').trim();
      if (!s){ blanks++; if (blanks>=3) break; }
      else {
        blanks=0;
        if (!/^enter\s|^type\s|^do not/i.test(s.toLowerCase())) out.push(s);
      }
    }
    const seen=new Set(), dedup=[];
    for(const x of out){const k=x.toLowerCase(); if(!seen.has(k)){seen.add(k);dedup.push(x);} }
    return dedup;
  }

  function chooseBestSheet(wb){
    const badName=/(example|instruction|cover|template)/i;
    const goodName=/(your\s+business|business)/i;
    let best=null,bestScore=-Infinity;
    for(const ws of wb.worksheets){
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
      const ws=XLSX.utils.sheet_to_json(wbRaw.Sheets[n],{header:1,defval:''});
      return {
        name:n,
        rowCount:ws.length,
        getCell:(r,c)=>({value:(ws[r-1]&&ws[r-1][c-1])||''})
      };
    })};
    const pick=chooseBestSheet(workbook);
    if(!pick)throw new Error("No suitable sheet found");
    let rooms=[],owners=[];
    if(pick.rHdr)rooms=extract_list_below(pick.ws,pick.rHdr,1);
    if(pick.uHdr)owners=extract_list_below(pick.ws,pick.uHdr,1);
    return {rooms,owners,sheetUsed:pick.name,foundRoomHeader:!!pick.rHdr,foundOwnerHeader:!!pick.uHdr};
  }

  // ---------- ExcelJS patch ----------
  function applyListsToWorkbook(workbook, lists){
    if(!workbook||!lists)return;
    let ws=workbook.getWorksheet('Lists');
    if(!ws)ws=workbook.addWorksheet('Lists');
    const max=ws.rowCount||1000;
    for(let r=2;r<=max;r++){ws.getCell(r,2).value=null;ws.getCell(r,5).value=null;}
    let i=2; (lists.rooms||[]).forEach(v=>ws.getCell(i++,2).value=v);
    i=2; (lists.owners||[]).forEach(v=>ws.getCell(i++,5).value=v);
  }

  function patchExcelJSOnce(){
    if(window.__TS_STEP4_PATCHED__)return;
    window.__TS_STEP4_PATCHED__=true;
    const ExcelJS=window.ExcelJS;
    if(!ExcelJS||!ExcelJS.Workbook){setTimeout(patchExcelJSOnce,500);return;}
    const orig=ExcelJS.Workbook.prototype.xlsx.writeBuffer;
    if(!orig)return;
    ExcelJS.Workbook.prototype.xlsx.writeBuffer=function(){
      try{
        if(window.__TS_STEP4_LISTS&&(window.__TS_STEP4_LISTS.rooms?.length||window.__TS_STEP4_LISTS.owners?.length)){
          applyListsToWorkbook(this,window.__TS_STEP4_LISTS);
        }
      }catch(e){console.warn('[Step4] Injection skipped',e);}
      return orig.apply(this,arguments);
    };
    console.log('[Step4] ExcelJS patched for Rooms/Owners injection');
  }

  // ---------- Wire UI ----------
  function wireUI(){
    const step3=findStep3Card();
    const step4=buildStep4From(step3);
    if(step3&&step3.parentElement)step3.parentElement.insertBefore(step4,step3.nextSibling);
    else document.body.appendChild(step4);
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
