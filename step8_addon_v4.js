
// 3seats Step-8 Addon (v4) — external JS version
(function(){
  if (window.__TS_STEP8_WIRED_V4_EXT__) return;
  window.__TS_STEP8_WIRED_V4_EXT__ = true;

  const TEMPLATE_ID = "1ft0PuCB3EneQ8vW9lFv78c1KBC1giUOGTAmLsa8bETE";
  const CLIENT_ID   = "7010858919-jq4n8blq1b73o26pq3h4n0uk46roqfag.apps.googleusercontent.com";

  function loadScriptOnce(src){
    return new Promise((resolve, reject)=>{
      if ([...document.scripts].some(s => (s.src||"").includes(src))) return resolve();
      const el = document.createElement('script');
      el.src = src; el.async = true; el.defer = true;
      el.onload = resolve; el.onerror = () => reject(new Error("Failed to load: "+src));
      document.head.appendChild(el);
    });
  }
  async function ensureGoogleLibs(){
    await loadScriptOnce("https://apis.google.com/js/api.js");
    await loadScriptOnce("https://accounts.google.com/gsi/client");
  }

  const $ = (id)=>document.getElementById(id);
  const setStatus = (m)=>{ const s=$('ts-export-status'); if(s) s.textContent=m; };

// Replace the whole function with this:
function findStep7Card(){
  // Prefer a heading that mentions "export" (any step number)
  const hs = Array.from(document.querySelectorAll('h1,h2,h3,h4,strong,b,.title,.header'));
  for (const h of hs){
    const raw = (h.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
    if (raw.includes('export')) {
      const card = h.closest('section, .step, .card, .panel, .box, .container, .chunk, .ts-card, div');
      if (card) return card;
    }
  }
  // Fallback: find a container that has a "Generate" button or "Export Format"
  const genBtn = Array.from(document.querySelectorAll('button,.button,input[type="submit"],[role="button"]'))
    .find(b => ((b.textContent||b.value||'').toLowerCase().includes('generate')));
  if (genBtn) return genBtn.closest('section, .step, .card, .panel, .box, .container, .chunk, .ts-card, div');

  const label = Array.from(document.querySelectorAll('*')).find(el =>
    (el.textContent||'').toLowerCase().includes('export format')
  );
  if (label) return label.closest('section, .step, .card, .panel, .box, .container, .chunk, .ts-card, div');

  return null;
}



  function buildStep8From(step7){
    const step8 = step7 ? step7.cloneNode(false) : document.createElement('div');
    step8.id = 'ts-step8-card';

    const h7 = step7 && step7.querySelector('h1,h2,h3,h4,.title,.header,strong,b');
    const content7 =
      (step7 && Array.from(step7.children).find(el => el !== h7 && el.tagName && el.tagName.toLowerCase() !== 'script')) ||
      (step7 && step7.querySelector('.content, .body, .card-content, .section-body, .panel-body, .p-*, .px-*, .py-*'));

    const HeadTag = h7 ? (h7.tagName || 'div').toLowerCase() : 'h3';
    const h8 = document.createElement(HeadTag);
    h8.textContent = '8) Export to Google Sheets';
    if (h7 && h7.classList.length) h7.classList.forEach(c => h8.classList.add(c));
    else h8.style.fontWeight = '600';

    const body = document.createElement(content7 ? content7.tagName.toLowerCase() : 'div');
    if (content7 && content7.classList.length) content7.classList.forEach(c => body.classList.add(c));

    const row = document.createElement('div');
    const row7 = step7 && step7.querySelector('.controls, .row, .fields, .inline-controls, .grid, .flex');
    if (row7 && row7.classList.length) row7.classList.forEach(c => row.classList.add(c));
    else { row.style.display='flex'; row.style.flexWrap='wrap'; row.style.alignItems='center'; row.style.gap='10px'; }

    // --- Input field (matches Step 7 styling) ---
const name = document.createElement('input');
name.id = 'ts-export-base';
name.placeholder = 'master_output (in Google)';

// Copy Step 7 input's classes so it inherits the dark theme
const refInput = step7 && step7.querySelector('input[type="text"], .input, input, select, textarea');
if (refInput) {
  if (refInput.classList.length) refInput.classList.forEach(c => name.classList.add(c));
  const cs = getComputedStyle(refInput);
  name.style.minWidth = '280px';
  name.style.padding = cs.padding;
  name.style.border = cs.border;
  name.style.borderRadius = cs.borderRadius;
  // Match the dark input background + text color
  name.style.backgroundColor = cs.backgroundColor || '#1e1e2a';
  name.style.color = cs.color || '#e5e7eb';
}


    const btn = document.createElement('button');
    btn.id = 'ts-export-btn';
    btn.textContent = 'Export to Google Sheet';
    const refBtn = step7 && step7.querySelector('button, .button, input[type="button"], input[type="submit"]');
    if (refBtn){
      if (refBtn.classList.length) refBtn.classList.forEach(c => btn.classList.add(c));
      const cs = getComputedStyle(refBtn);
      btn.style.padding = cs.padding; btn.style.borderRadius = cs.borderRadius; btn.style.border = cs.border;
    }

    const status = document.createElement('span');
    status.id = 'ts-export-status';
    status.textContent = 'Idle';
    const helper = step7 && step7.querySelector('.help,.hint,.description,small,.text-muted');
    if (helper && helper.classList.length) helper.classList.forEach(c => status.classList.add(c));

    const desc = document.createElement('div');
    if (helper && helper.classList.length) helper.classList.forEach(c => desc.classList.add(c));
    desc.innerHTML = 'Pick the Excel you generated → Google converts it → we copy <b>Events</b> as values and <b>Lists</b> with formulas preserved into your Template copy.';

    row.appendChild(name); row.appendChild(btn); row.appendChild(status);
    body.appendChild(desc); body.appendChild(row);
    step8.appendChild(h8); step8.appendChild(body);
    return step8;
  }

  function hideOldStep7Button(step7){
    const scope = step7 || document;
    scope.querySelectorAll('button, a, input[type="button"], input[type="submit"]').forEach(el=>{
      const txt = (el.textContent || el.value || '').toLowerCase();
      if (txt.includes('export') && txt.includes('google') && !el.closest('#ts-step8-card')){
        el.style.display = 'none';
      }
    });
  }

  async function pickExcel(){ return await new Promise(res=>{ const i=document.createElement('input'); i.type='file'; i.accept='.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; i.style.position='fixed'; i.style.left='-9999px'; document.body.appendChild(i); i.onchange=()=>res(i.files?.[0]||null); i.click(); }); }
  async function ensureGapi(){ await new Promise((ok,err)=>{ gapi.load('client', async()=>{ try{ await gapi.client.init({}); ok(); }catch(e){ err(e); } }); }); await gapi.client.load('https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'); await gapi.client.load('https://sheets.googleapis.com/$discovery/rest?version=v4'); }
  async function ensureToken(){ return await new Promise((ok,err)=>{ const t=google.accounts.oauth2.initTokenClient({client_id:CLIENT_ID,scope:'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheets',use_fedcm_for_prompt:true,callback:()=>{}}); t.callback=(r)=>r?.access_token?ok(r.access_token):err(new Error(r?.error||'No token')); t.requestAccessToken({prompt:'consent'}); }); }
  async function readValues(id, sheet){ const r=await gapi.client.sheets.spreadsheets.values.get({spreadsheetId:id, range:`${sheet}!A1:ZZZ100000`}); return r.result.values||[]; }
  async function writeValues(id, sheet, vals){ if(!vals.length)return; await gapi.client.sheets.spreadsheets.values.update({spreadsheetId:id, range:`${sheet}!A1`, valueInputOption:'RAW', resource:{values:vals}}); }
  async function readFormulas(id, sheet){ const r=await gapi.client.sheets.spreadsheets.get({spreadsheetId:id, ranges:[`${sheet}!A1:ZZZ100000`], includeGridData:true, fields:'sheets(data(rowData(values(userEnteredValue))))'}); return r.result.sheets?.[0]?.data?.[0]?.rowData||[]; }
  async function writeFormulas(id, sheet, rows){ const meta=await gapi.client.sheets.spreadsheets.get({spreadsheetId:id, fields:'sheets(properties(sheetId,title))'}); const sid=meta.result.sheets.find(s=>s.properties.title===sheet)?.properties.sheetId; await gapi.client.sheets.spreadsheets.batchUpdate({spreadsheetId:id, resource:{requests:[{updateCells:{range:{sheetId:sid,startRowIndex:0,startColumnIndex:0}, rows:rows, fields:'userEnteredValue'}}]}}); }

  async function exportNow(){
    try{
      setStatus('Pick the Excel you just downloaded from Generate…');
      const picked = await pickExcel();
      if(!picked){ setStatus('Canceled.'); return; }

      await ensureGoogleLibs();
      await ensureGapi();
      const token = await ensureToken();
      gapi.client.setToken({access_token: token});

      const baseName = (($('ts-export-base')?.value || 'master_output').trim()||'master_output').replace(/\s+/g,'_');

      setStatus('Copying template…');
      const copy = await gapi.client.drive.files.copy({ fileId:TEMPLATE_ID, fields:'id', resource:{ name:`${baseName} (Google Sheet)` }});
      const destId = copy.result.id;

      setStatus('Uploading & converting…');
      const meta = { name: baseName + ' (Converted)', mimeType: 'application/vnd.google-apps.spreadsheet' };
      const form = new FormData();
      form.append('metadata', new Blob([JSON.stringify(meta)], {type:'application/json'}));
      form.append('file', picked, picked.name);
      const conv = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', { method:'POST', headers:{Authorization:'Bearer '+token}, body:form });
      const {id:convertedId} = await conv.json();

      const info = await gapi.client.sheets.spreadsheets.get({ spreadsheetId: convertedId, fields:'sheets(properties(title))' });
      const titles = (info.result.sheets||[]).map(s=>s.properties.title);
      const find = (c)=>{ const lc=titles.map(t=>t.toLowerCase()); for(const cand of c){ const i=lc.indexOf(cand.toLowerCase()); if(i>=0) return titles[i]; } return null; };
      const eTitle = find(['Events','Event','Sheet1']) || titles[0];
      const lTitle = find(['Lists','List']);

      async function ensure(title){
        const r = await gapi.client.sheets.spreadsheets.get({ spreadsheetId: destId, fields:'sheets(properties(title))' });
        const ok = (r.result.sheets||[]).some(s=>s.properties.title===title);
        if (!ok) await gapi.client.sheets.spreadsheets.batchUpdate({ spreadsheetId: destId, resource:{requests:[{addSheet:{properties:{title}}}] }});
      }
      await ensure('Events'); await ensure('Lists');

      setStatus('Copying Events…');
      const events = await readValues(convertedId, eTitle);
      await gapi.client.sheets.spreadsheets.values.clear({ spreadsheetId: destId, range: 'Events!A1:ZZZ100000' });
      await writeValues(destId, 'Events', events);

      if (lTitle){
        setStatus('Copying Lists…');
        const rows = await readFormulas(convertedId, lTitle);
        await gapi.client.sheets.spreadsheets.values.clear({ spreadsheetId: destId, range: 'Lists!A1:ZZZ100000' });
        await writeFormulas(destId, 'Lists', rows);
      }

      setStatus('Opening Google Sheet…');
      window.open(`https://docs.google.com/spreadsheets/d/${destId}/edit`, '_blank');
      setStatus('Done');
    }catch(e){
      console.error(e);
      alert('Export failed: ' + (e?.message || e));
    }
  }

  function wireUI(){
    const step7 = findStep7Card();
    if (!step7) return;
    const step8 = buildStep8From(step7);
    step7.parentElement.insertBefore(step8, step7.nextSibling);
    hideOldStep7Button(step7);
    document.getElementById('ts-export-btn')?.addEventListener('click', (e)=>{ e.preventDefault(); exportNow(); });
  }

  document.addEventListener('DOMContentLoaded', wireUI);
})();
