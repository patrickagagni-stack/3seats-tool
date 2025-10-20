// Step 8 Addon (Uploader Variant) — v4u (renumbered to 6 in UI)
(function(){
  if (window.__TS_STEP8_V4U__) return;
  window.__TS_STEP8_V4U__ = true;

  const OAUTH_CLIENT_ID = "7010858919-jq4n8blq1b73o26pq3h4n0uk46roqfag.apps.googleusercontent.com";
  const TEMPLATE_ID     = "1ft0PuCB3EneQ8vW9lFv78c1KBC1giUOGTAmLsa8bETE";
  const SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets"
  ].join(" ");

  function loadScriptOnce(src){
    return new Promise((resolve,reject)=>{
      if ([...document.scripts].some(s=>(s.src||"").includes(src))) return resolve();
      const el = document.createElement("script");
      el.src = src; el.async = true; el.defer = true;
      el.onload = resolve; el.onerror = () => reject(new Error("Failed to load "+src));
      document.head.appendChild(el);
    });
  }
  async function ensureGapi(){
    await loadScriptOnce("https://apis.google.com/js/api.js");
    if (!window.gapi) throw new Error("gapi failed to load");
    if (!gapi.client) await new Promise(res=>gapi.load("client", res));
  }
  async function gapiInit(){
    await ensureGapi();
    if (!gapi.client.__ts_init){
      await gapi.client.init({
        clientId: OAUTH_CLIENT_ID,
        scope: SCOPES,
        discoveryDocs: [
          "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest",
          "https://www.googleapis.com/discovery/v1/apis/sheets/v4/rest"
        ]
      });
      gapi.client.__ts_init = true;
    }
  }
  async function ensureSheetJS(){
    await loadScriptOnce("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
    if (!window.XLSX) throw new Error("SheetJS failed to load");
  }

  function findExportCard(){
    const hs = Array.from(document.querySelectorAll("h1,h2,h3,h4,strong,b,.title,.header"));
    for (const h of hs){
      const t = (h.textContent||"").toLowerCase();
      if (t.includes("export")){
        const card = h.closest("section, .step, .card, .panel, .box, .container, .chunk, .ts-card, div") || h.parentElement;
        if (card) return card;
      }
    }
    return null;
  }

  function buildStep6Card(){
    const host = document.createElement("section");
    host.className = "card";
    host.id = "ts-step6-card";

    const h2 = document.createElement("h2");
    h2.textContent = "6) Export to Google Sheets";   // ← just the visible label changed
    host.appendChild(h2);

    const content = document.createElement("div");
    content.className = "content";
    content.style.paddingTop = "8px";
    content.style.paddingBottom = "12px";

    const note = document.createElement("div");
    note.innerHTML = "Pick the Excel you just generated (or another .xlsx). We'll copy your Google template and push all <b>Events</b> & <b>Lists</b> values into it.";
    content.appendChild(note);

    const row = document.createElement("div");
    row.style.display = "flex";
    row.style.flexWrap = "wrap";
    row.style.alignItems = "center";
    row.style.gap = "8px";
    row.style.marginTop = "8px";

    const file = document.createElement("input");
    file.type = "file";
    file.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    file.id = "ts-step6-file";
    row.appendChild(file);

    const name = document.createElement("input");
    name.type = "text";
    name.placeholder = "master_output (in Google)";
    name.id = "ts-step6-name";
    name.value = "master_output";
    name.style.minWidth = "240px";
    row.appendChild(name);

    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = "Export to Google Sheet";
    btn.className = "btn";
    row.appendChild(btn);

    const status = document.createElement("div");
    status.id = "ts-step6-status";
    status.style.marginTop = "8px";
    status.textContent = "Ready.";
    content.appendChild(row);
    content.appendChild(status);

    host.appendChild(content);

    btn.addEventListener("click", async ()=>{
      try{
        const f = file.files && file.files[0];
        if (!f){ status.textContent = "Please choose an .xlsx file first."; return; }
        status.textContent = "Authorizing with Google…";
        await gapiInit();

        const auth = gapi.auth2.getAuthInstance?.();
        if (auth){
          const user = auth.currentUser.get();
          const has = user && user.hasGrantedScopes && user.hasGrantedScopes(SCOPES);
          if (!has) await auth.signIn({ scope: SCOPES });
        }

        status.textContent = "Copying template…";
        const copyRes = await gapi.client.drive.files.copy({
          fileId: "1ft0PuCB3EneQ8vW9lFv78c1KBC1giUOGTAmLsa8bETE",
          resource: { name: (name.value || "master_output") + " (in Google)" },
          fields: "id"
        });
        const destId = copyRes.result.id;
        if (!destId) throw new Error("Template copy failed.");

        status.textContent = "Reading Excel…";
        await ensureSheetJS();
        const buf = await f.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });

        const sheetNames = wb.SheetNames;
        const eventsName = sheetNames.find(n=>/^events$/i.test(n)) || sheetNames[0];
        const listsName  = sheetNames.find(n=>/^lists$/i.test(n))  || null;

        const eventsAOA = XLSX.utils.sheet_to_json(wb.Sheets[eventsName], { header:1, defval:"" });
        const listsAOA  = listsName ? XLSX.utils.sheet_to_json(wb.Sheets[listsName],  { header:1, defval:"" }) : [];

        status.textContent = "Writing data to Google Sheet…";
        const data = [];
        if (eventsAOA && eventsAOA.length){
          data.push({ range: "Events!A1", majorDimension: "ROWS", values: eventsAOA });
        }
        if (listsAOA && listsAOA.length){
          data.push({ range: "Lists!A1", majorDimension: "ROWS", values: listsAOA });
        }
        if (data.length){
          await gapi.client.sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: destId,
            resource: { valueInputOption: "USER_ENTERED", data }
          });
        }

        status.innerHTML = `✅ Done. <a target="_blank" rel="noopener" href="https://docs.google.com/spreadsheets/d/${destId}">Open your Google Sheet</a>`;
      }catch(e){
        console.error(e);
        status.textContent = "❌ Export failed: " + (e.result?.error?.message || e.message || e);
      }
    });

    return host;
  }

  function inject(){
    const anchor = findExportCard();
    if (!anchor) return;
    anchor.parentElement.insertBefore(buildStep6Card(), anchor.nextSibling);
  }

  document.addEventListener("DOMContentLoaded", inject);
})();
