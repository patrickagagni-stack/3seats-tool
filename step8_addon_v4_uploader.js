// step8_addon_v4_uploader.js — Step "6) Export to Google Sheets" using GIS (v9)
// Copies formulas in Lists (e.g., columns I–L) from Excel into Google Sheets.
(function(){
  if (window.__TS_STEP8_V4U__) return; window.__TS_STEP8_V4U__ = true;

  const OAUTH_CLIENT_ID = "7010858919-jq4n8blq1b73o26pq3h4n0uk46roqfag.apps.googleusercontent.com";
  const TEMPLATE_ID     = "1ft0PuCB3EneQ8vW9lFv78c1KBC1giUOGTAmLsa8bETE";
  const SCOPES          = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets";

  // ---------- loaders ----------
  function loadScriptOnce(src){
    return new Promise((resolve,reject)=>{
      if ([...document.scripts].some(s=>(s.src||"").includes(src))) return resolve();
      const s=document.createElement("script"); s.src=src; s.async=true; s.defer=true;
      s.onload=resolve; s.onerror=()=>reject(new Error("Failed to load "+src));
      document.head.appendChild(s);
    });
  }
  async function ensureGapi(){
    await loadScriptOnce("https://apis.google.com/js/api.js");
    if (!window.gapi) throw new Error("gapi failed to load");
    if (!gapi.client) await new Promise(res=>gapi.load("client",res));
    if (!gapi.client.__ts_inited){
      await gapi.client.init({
        discoveryDocs:[
          "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest",
          "https://www.googleapis.com/discovery/v1/apis/sheets/v4/rest"
        ]
      });
      gapi.client.__ts_inited = true;
    }
  }
  async function ensureGIS(){
    await loadScriptOnce("https://accounts.google.com/gsi/client");
    if (!google?.accounts?.oauth2) throw new Error("Google Identity Services failed to load");
  }
  async function ensureSheetJS(){
    await loadScriptOnce("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
    if (!window.XLSX) throw new Error("SheetJS failed to load");
  }

  // ---------- helpers ----------
  function colLetter(n){ let s=""; while(n){ let r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26);} return s; }
  function aoaSize(aoa){
    const rows = aoa.length || 1;
    const cols = aoa.reduce((m,r)=>Math.max(m, (r?.length)||0), 0) || 1;
    return {rows, cols, endA1: `${colLetter(cols)}${rows}`};
  }

  // Build AOA from worksheet, preserving formulas: if cell.f exists, use "="+cell.f
  function sheetToAOAWithFormulas(ws){
    const r = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
    const aoa = [];
    for(let R=r.s.r; R<=r.e.r; ++R){
      const row=[];
      for(let C=r.s.c; C<=r.e.c; ++C){
        const addr = XLSX.utils.encode_cell({r:R,c:C});
        const cell = ws[addr];
        if (!cell) { row.push(""); continue; }
        if (cell.f != null && cell.f !== "") {
          row.push("=" + cell.f);
        } else {
          // prefer displayed text; fallback to raw value
          row.push(cell.w != null ? cell.w : (cell.v != null ? cell.v : ""));
        }
      }
      aoa.push(row);
    }
    return aoa;
  }

  // GIS token
  let tokenClient=null;
  function getAccessToken({forcePrompt=false}={}){
    return new Promise(async (resolve,reject)=>{
      try{
        await ensureGIS();
        tokenClient ||= google.accounts.oauth2.initTokenClient({
          client_id: OAUTH_CLIENT_ID,
          scope: SCOPES,
          callback: (resp)=> resp?.access_token ? resolve(resp.access_token) : reject(new Error("No access token"))
        });
        tokenClient.requestAccessToken({ prompt: forcePrompt ? "consent" : "" });
      }catch(e){ reject(e); }
    });
  }

  // ---------- UI ----------
  function findExportCard(){
    const hs=[...document.querySelectorAll("h1,h2,h3,h4,strong,b,.title,.header")];
    for(const h of hs){
      const t=(h.textContent||"").toLowerCase();
      if (t.includes("export")){
        return h.closest("section,.step,.card,.panel,.box,.container,.chunk,.ts-card,div") || h.parentElement;
      }
    }
    return null;
  }

  function buildStep6Card(){
    const host=document.createElement("section"); host.className="card"; host.id="ts-step6-card";
    const h2=document.createElement("h2"); h2.textContent="6) Export to Google Sheets"; host.appendChild(h2);

    const content=document.createElement("div"); content.className="content"; content.style.padding="8px 0 12px";
    const note=document.createElement("div");
    note.innerHTML="Pick the Excel you just generated (or another .xlsx). We'll copy your Google template and push all <b>Events</b> values and <b>Lists</b> values <i>including formulas</i> into it.";
    content.appendChild(note);

    const row=document.createElement("div");
    Object.assign(row.style,{display:"flex",flexWrap:"wrap",alignItems:"center",gap:"8px",marginTop:"8px"});

    const file=document.createElement("input"); file.type="file";
    file.accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; file.id="ts-step6-file"; row.appendChild(file);

    const name=document.createElement("input"); name.type="text"; name.placeholder="master_output (in Google)";
    name.id="ts-step6-name"; name.value="master_output"; name.style.minWidth="240px"; row.appendChild(name);

    const btn=document.createElement("button"); btn.type="button"; btn.textContent="Export to Google Sheet"; btn.className="btn"; row.appendChild(btn);

    const status=document.createElement("div"); status.id="ts-step6-status"; status.style.marginTop="8px"; status.textContent="Ready.";

    content.appendChild(row); content.appendChild(status); host.appendChild(content);

    btn.addEventListener("click", async ()=>{
      try{
        const f=file.files?.[0]; if(!f){ status.textContent="Please choose an .xlsx file first."; return; }

        status.textContent="Preparing libraries…"; await ensureGapi(); await ensureSheetJS();

        status.textContent="Authorizing with Google…";
        const access_token=await getAccessToken(); gapi.client.setToken({access_token});

        status.textContent="Copying template…";
        const copyRes=await gapi.client.drive.files.copy({
          fileId:TEMPLATE_ID, supportsAllDrives:true, fields:"id",
          resource:{ name:(name.value||"master_output")+" (in Google)" }
        });
        const destId=copyRes.result.id; if(!destId) throw new Error("Template copy failed.");

        status.textContent="Reading Excel…";
        const buf=await f.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array"});

        const eventsName = wb.SheetNames.find(n=>/^events$/i.test(n)) || wb.SheetNames[0];
        const listsName  = wb.SheetNames.find(n=>/^lists$/i.test(n))  || null;

        // Events: values only (fast)
        const eventsAOA = XLSX.utils.sheet_to_json(wb.Sheets[eventsName], { header:1, defval:"" });

        // Lists: formulas preserved
        let listsAOA = [];
        if (listsName) listsAOA = sheetToAOAWithFormulas(wb.Sheets[listsName]);

        status.textContent="Writing data to Google Sheet…";
        const data=[];

        if (eventsAOA.length){
          const {rows, cols, endA1} = aoaSize(eventsAOA);
          data.push({ range:`Events!A1:${endA1}`, majorDimension:"ROWS", values:eventsAOA });
        }
        if (listsAOA.length){
          const {endA1} = aoaSize(listsAOA);
          data.push({ range:`Lists!A1:${endA1}`, majorDimension:"ROWS", values:listsAOA });
        }

        if (data.length){
          await gapi.client.sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: destId,
            resource: { valueInputOption: "USER_ENTERED", data }
          });
        }

        status.innerHTML=`✅ Done. <a target="_blank" rel="noopener" href="https://docs.google.com/spreadsheets/d/${destId}">Open your Google Sheet</a>`;
      }catch(e){
        console.error(e);
        const msg = e?.result?.error?.message || e?.details || e?.message || JSON.stringify(e);
        status.textContent = "❌ Export failed: " + msg + (msg?.includes("idpiframe") ? " (Tip: make sure the Google sign-in pop-up wasn't blocked.)" : "");
      }
    });

    return host;
  }

  function inject(){
    const anchor=findExportCard(); if(!anchor) return;
    anchor.parentElement.insertBefore(buildStep6Card(), anchor.nextSibling);
  }

  document.addEventListener("DOMContentLoaded", inject);
})();
