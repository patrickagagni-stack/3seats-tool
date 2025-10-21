// step8_addon_v4_uploader.js — Step "6) Export to Google Sheets" (UI-aligned)
// Logic unchanged; only Step 6 UI (note + input row) styling is adjusted for alignment/spacing.
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
    if (!gapi.client) await new Promise(res=>gapi.load("client", res));
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

  // ---------- auth via GIS ----------
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

  // ---------- Drive helpers ----------
  async function driveCopyTemplate(name){
    const res = await gapi.client.drive.files.copy({
      fileId: TEMPLATE_ID,
      supportsAllDrives: true,
      fields: "id",
      resource: { name }
    });
    const id = res.result.id;
    if (!id) throw new Error("Template copy failed.");
    return id;
  }

  // Upload XLSX -> convert to Google Sheet using multipart/related upload
  async function driveUploadConvertXlsx(file, token, name){
    const boundary = "-------314159265358979323846";
    const metadata = {
      name: name + " (converted)",
      mimeType: "application/vnd.google-apps.spreadsheet"
    };
    const body = new Blob([
      `--${boundary}\r\n`+
      "Content-Type: application/json; charset=UTF-8\r\n\r\n"+
      JSON.stringify(metadata)+
      `\r\n--${boundary}\r\n`+
      "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n",
      file,
      `\r\n--${boundary}--`
    ], { type: `multipart/related; boundary=${boundary}` });

    const url = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id";
    const res = await fetch(url, {
      method: "POST",
      headers: { "Authorization": "Bearer " + token },
      body
    });
    const json = await res.json();
    if (!res.ok) throw new Error(json?.error?.message || JSON.stringify(json));
    if (!json.id) throw new Error("Conversion upload did not return an id.");
    return json.id;
  }

  async function driveDelete(fileId){
    try{ await gapi.client.drive.files.delete({ fileId }); }catch(_e){}
  }

  // ---------- Sheets helpers ----------
  async function getSheetsMeta(spreadsheetId){
    const meta = await gapi.client.sheets.spreadsheets.get({ spreadsheetId });
    return meta.result.sheets.map(s => s.properties); // {sheetId, title, index, ...}
  }
  function findSheetIdByTitle(props, title){
    const p = props.find(p => (p.title||"").toLowerCase() === title.toLowerCase());
    return p ? p.sheetId : null;
    }

  async function copyTabTo(targetSpreadsheetId, sourceSpreadsheetId, sourceSheetId){
    const res = await gapi.client.sheets.spreadsheets.sheets.copyTo({
      spreadsheetId: sourceSpreadsheetId,
      sheetId: sourceSheetId,
      resource: { destinationSpreadsheetId: targetSpreadsheetId }
    });
    return res.result.sheetId;
  }

  async function batchUpdate(spreadsheetId, requests){
    if (!requests.length) return;
    await gapi.client.sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: { requests }
    });
  }

  async function replaceSheetWithCopy({ destId, destTitle, srcId, srcTitle }){
    const destProps = await getSheetsMeta(destId);
    const srcProps  = await getSheetsMeta(srcId);

    const srcSheetId  = findSheetIdByTitle(srcProps , srcTitle);
    if (!srcSheetId) throw new Error(`Source sheet "${srcTitle}" not found in converted file.`);
    const destSheetId = findSheetIdByTitle(destProps, destTitle);

    const newDestSheetId = await copyTabTo(destId, srcId, srcSheetId);

    const requests = [];
    if (destSheetId) requests.push({ deleteSheet: { sheetId: destSheetId } });
    requests.push({
      updateSheetProperties: {
        properties: { sheetId: newDestSheetId, title: destTitle, index: 0 },
        fields: "title,index"
      }
    });
    await batchUpdate(destId, requests);
  }

  // ---------- locate Step 5 card to insert after ----------
  function findExportCard(){
    const hs=[...document.querySelectorAll("h1,h2,h3,h4,strong,b,.title,.header")];
    for(const h of hs){
      const t=(h.textContent||"").toLowerCase();
      if (t.includes("export") && !t.includes("google")){
        return h.closest("section,.step,.card,.panel,.box,.container,.chunk,.ts-card,div") || h.parentElement;
      }
    }
    return null;
  }

  // ---------- Step 6 UI ----------
  function buildStep6Card(){
    const host=document.createElement("section"); host.className="card"; host.id="ts-step6-card";
    const h2=document.createElement("h2"); h2.textContent="6) Export to Google Sheets"; host.appendChild(h2);

    const content=document.createElement("div"); content.className="content";

    // Note (keep your exact text)
    const note=document.createElement("div");
    note.innerHTML = "Pick the Excel you generated. We’ll <b>convert</b> it to a temporary Google Sheet (keeping formulas &amp; formatting), copy both tabs into your template copy, then clean up.";
    // --- style JUST this line, and align under header text ---
    note.style.margin    = "4px 0 12px 12px";  // indent matches header text start
    note.style.fontSize  = "0.9rem";
    note.style.lineHeight= "1.45";
    note.style.color     = "#9ca3af";         // subtle gray
    note.style.maxWidth  = "95%";
    content.appendChild(note);

    // Single-line input row (file + name + button)
    const row=document.createElement("div");
    // Use flex intentionally (one line), and indent to align with header text start
    row.style.display      = "flex";
    row.style.flexWrap     = "nowrap";
    row.style.alignItems   = "center";
    row.style.gap          = "10px";
    row.style.margin       = "0 0 8px 12px";

    // File picker (compact)
    const fileLabel = document.createElement("label");
    fileLabel.style.display = "inline-flex";
    fileLabel.style.alignItems = "center";
    fileLabel.style.gap = "6px";

    const fileSpan  = document.createElement("span");
    fileSpan.textContent = "Choose File";
    fileSpan.style.opacity = "0.8";

    const file = document.createElement("input");
    file.type = "file";
    file.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    file.id = "ts-step6-file";
    // narrow the control so row fits in one line on typical widths
    file.style.maxWidth = "190px";

    fileLabel.appendChild(fileSpan);
    fileLabel.appendChild(file);
    row.appendChild(fileLabel);

    // Name input (compact)
    const name = document.createElement("input");
    name.type = "text";
    name.placeholder = "master_output";
    name.id = "ts-step6-name";
    name.value = "master_output";
    name.style.width = "160px";
    name.style.minWidth = "140px";
    row.appendChild(name);

    // Primary action button
    const btn=document.createElement("button");
    btn.type="button";
    btn.textContent="Export to Google Sheet";
    btn.className="btn primary";
    row.appendChild(btn);

    const status=document.createElement("div");
    status.id="ts-step6-status";
    status.style.margin = "8px 0 0 12px";  // align with the row/note indent
    status.textContent="Ready.";

    content.appendChild(row);
    content.appendChild(status);
    host.appendChild(content);

    // ---------- click handler ----------
    btn.addEventListener("click", async ()=>{
      let tempId = null;
      try{
        const f=file.files?.[0]; if(!f){ status.textContent="Please choose an .xlsx file first."; return; }

        status.textContent="Preparing libraries…"; await ensureGapi();

        status.textContent="Authorizing with Google…";
        const access_token = await getAccessToken(); gapi.client.setToken({ access_token });

        const outName = (name.value || "master_output");
        status.textContent="Copying template…";
        const destId = await driveCopyTemplate(outName + " (Google)");

        status.textContent="Uploading & converting Excel → Google Sheet…";
        tempId = await driveUploadConvertXlsx(f, access_token, outName);

        // Copy both tabs wholesale (preserves formulas & formatting)
        status.textContent='Copying "Events" tab…';
        await replaceSheetWithCopy({ destId, destTitle: "Events", srcId: tempId, srcTitle: "Events" });

        status.textContent='Copying "Lists" tab…';
        await replaceSheetWithCopy({ destId, destTitle: "Lists", srcId: tempId, srcTitle: "Lists" });

        status.textContent="Cleaning up…";
        await driveDelete(tempId); tempId = null;

        const gUrl = `https://docs.google.com/spreadsheets/d/${destId}`;
        status.innerHTML = `✅ Done. <a target="_blank" rel="noopener" href="${gUrl}">Open your Google Sheet</a>`;
        try { setTimeout(() => window.open(gUrl, "_blank", "noopener"), 250); } catch(_) {}
      } catch (e) {
        console.error(e);
        const msg = e?.result?.error?.message || e?.details || e?.message || JSON.stringify(e);
        status.textContent =
          "❌ Export failed: " + msg +
          (msg?.includes("idpiframe") ? " (Tip: make sure the Google sign-in pop-up wasn't blocked.)" : "");
      }
    });

    return host;
  }

  function inject() {
    const anchor = findExportCard();
    if (!anchor) return;
    anchor.parentElement.insertBefore(buildStep6Card(), anchor.nextSibling);
  }

  document.addEventListener("DOMContentLoaded", inject);
})();
