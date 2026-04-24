/* ═══════════════════════════════════════════════════════════════════════════
   zentex-sp.js  –  SharePoint REST API + Offline-Queue
   ═══════════════════════════════════════════════════════════════════════════
   Voraussetzungen:
     - PWA gehostet auf SharePoint (Session-Cookie vorhanden)
     - Monteur hat Mitwirken-Berechtigung auf der Site
     - Keine App Registration, kein MSAL, kein Premium

   Status-Logik:
     "Neu"           → Job angelegt (optional, für spätere Erweiterung)
     "InArbeit"      → Monteur hat Job abgeschlossen, Daten in SharePoint
     "Abgeschlossen" → Power Automate hat E-Mail gesendet
═══════════════════════════════════════════════════════════════════════════ */
(function(window) {
  "use strict";

  /* ── KONFIGURATION ──────────────────────────────────────────────────── */
  var SITE_URL = "https://zentexbrandschutz.sharepoint.com/sites/Sprinkler-ChecklisteEntwurf";

  // Listen-Titel (genau wie in SharePoint angezeigt)
  var LIST_JOBS      = "Sprinkler_Jobs";
  var LIST_JOBITEMS  = "Sprinkler_JobItems";
  var LIST_REGIE     = "Sprinkler_Regie";
  var LIST_CLOSERUNS = "JobCloseRuns";

  // localStorage-Key für Offline-Queue
  var QUEUE_KEY = "zentex_sp_queue_v1";

  /* ── INDEXEDDB FÜR PDF-OFFLINE-SPEICHERUNG ─────────────────────────── */
  // Nur für grosse PDF-Daten. Queue-Metadaten bleiben in localStorage.
  var IDB_NAME    = "zentex_pdf_store";
  var IDB_VERSION = 1;
  var IDB_STORE   = "pdfs";

  function _idbOpen(){
    return new Promise(function(resolve, reject){
      var req = indexedDB.open(IDB_NAME, IDB_VERSION);
      req.onupgradeneeded = function(e){
        var db = e.target.result;
        if(!db.objectStoreNames.contains(IDB_STORE)){
          db.createObjectStore(IDB_STORE);
        }
      };
      req.onsuccess = function(e){ resolve(e.target.result); };
      req.onerror   = function(e){ reject(e.target.error); };
    });
  }

  function _pdfKey(entry){
    return [
      entry.jobId || "unknown",
      entry.type || "job",
      entry.queuedAt || "0"
    ].join("_");
  }

  async function idbSavePdf(key, base64){
    var db;
    try{
      db = await _idbOpen();
      return await new Promise(function(resolve, reject){
        var tx = db.transaction(IDB_STORE, "readwrite");
        var req = tx.objectStore(IDB_STORE).put(base64, key);
        req.onsuccess = function(){ resolve(true); };
        req.onerror   = function(e){ reject(e.target.error); };
      });
    }catch(e){
      console.warn("idbSavePdf fehlgeschlagen:", e && e.message || e);
      return false;
    }finally{
      if(db) db.close();
    }
  }

  async function idbLoadPdf(key){
    var db;
    try{
      db = await _idbOpen();
      return await new Promise(function(resolve){
        var tx = db.transaction(IDB_STORE, "readonly");
        var req = tx.objectStore(IDB_STORE).get(key);
        req.onsuccess = function(e){ resolve(e.target.result || null); };
        req.onerror   = function(){ resolve(null); };
      });
    }catch(e){
      console.warn("idbLoadPdf fehlgeschlagen:", e && e.message || e);
      return null;
    }finally{
      if(db) db.close();
    }
  }

  async function idbDeletePdf(key){
    var db;
    try{
      db = await _idbOpen();
      return await new Promise(function(resolve){
        var tx = db.transaction(IDB_STORE, "readwrite");
        tx.objectStore(IDB_STORE).delete(key);
        tx.oncomplete = function(){ resolve(true); };
        tx.onerror    = function(){ resolve(false); };
      });
    }catch(e){
      console.warn("idbDeletePdf fehlgeschlagen:", e && e.message || e);
      return false;
    }finally{
      if(db) db.close();
    }
  }

  /* ── REQUEST DIGEST ─────────────────────────────────────────────────── */
  // SharePoint verlangt einen CSRF-Token (FormDigestValue) für POST-Requests.
  // Gültig ca. 30 Minuten, wird bei Bedarf automatisch erneuert.
  var _digest      = null;
  var _digestExpiry = 0;

  async function getDigest(){
    // Cached Digest verwenden, wenn noch gültig (5 Min Puffer)
    if(_digest && Date.now() < _digestExpiry - 300000){
      return _digest;
    }
    var resp = await fetch(SITE_URL + "/_api/contextinfo", {
      method: "POST",
      headers: { "Accept": "application/json;odata=noop" },
      credentials: "include"
    });
    if(!resp.ok){
      throw new Error("Digest-Fehler: HTTP " + resp.status + " – Bitte prüfen, ob SharePoint-Login aktiv ist.");
    }
    var data = await resp.json();
    _digest      = data.FormDigestValue;
    _digestExpiry = Date.now() + (data.FormDigestTimeoutSeconds || 1800) * 1000;
    return _digest;
  }

  /* ── SHAREPOINT LIST OPERATIONS ─────────────────────────────────────── */

  // Listeneintrag erstellen
  async function createItem(listTitle, fields){
    var digest = await getDigest();
    var resp = await fetch(
      SITE_URL + "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')/items",
      {
        method: "POST",
        headers: {
          "Accept":          "application/json;odata=noop",
          "Content-Type":    "application/json;odata=noop",
          "X-RequestDigest": digest
        },
        credentials: "include",
        body: JSON.stringify(fields)
      }
    );
    if(!resp.ok){
      var errText = "";
      try{ errText = await resp.text(); }catch(e){}
      throw new Error("SP Create " + listTitle + ": HTTP " + resp.status + " – " + errText.substring(0, 200));
    }
    return resp.json();
  }

  // Listeneintrag aktualisieren (braucht SP-interne ID)
  async function updateItem(listTitle, itemId, fields){
    var digest = await getDigest();
    var resp = await fetch(
      SITE_URL + "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')/items(" + itemId + ")",
      {
        method: "POST",
        headers: {
          "Accept":          "application/json;odata=noop",
          "Content-Type":    "application/json;odata=noop",
          "X-RequestDigest": digest,
          "IF-MATCH":        "*",
          "X-HTTP-Method":   "MERGE"
        },
        credentials: "include",
        body: JSON.stringify(fields)
      }
    );
    if(!resp.ok){
      var errText = "";
      try{ errText = await resp.text(); }catch(e){}
      throw new Error("SP Update " + listTitle + " #" + itemId + ": HTTP " + resp.status + " – " + errText.substring(0, 200));
    }
    return true;
  }

  // Prüfen ob ein Eintrag mit einem bestimmten Feldwert existiert
  async function findItemByField(listTitle, fieldName, fieldValue){
    var filter = encodeURIComponent(fieldName + " eq '" + fieldValue.replace(/'/g, "''") + "'");
    var resp = await fetch(
      SITE_URL + "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')/items?$filter=" + filter + "&$top=1&$select=Id," + fieldName,
      {
        method: "GET",
        headers: { "Accept": "application/json;odata=noop" },
        credentials: "include"
      }
    );
    if(!resp.ok) return null;
    var data = await resp.json();
    return (data.value && data.value.length > 0) ? data.value[0] : null;
  }

  // Upsert: Erstellen wenn nicht vorhanden, sonst aktualisieren
  async function upsertItem(listTitle, uniqueField, uniqueValue, fields){
    var existing = await findItemByField(listTitle, uniqueField, uniqueValue);
    if(existing){
      await updateItem(listTitle, existing.Id, fields);
      return { action: "updated", id: existing.Id };
    }else{
      var result = await createItem(listTitle, fields);
      return { action: "created", id: result.Id };
    }
  }

  /* ── PDF UPLOAD IN DOKUMENTBIBLIOTHEK ─────────────────────────── */

  var PDF_FOLDER = "/sites/Sprinkler-ChecklisteEntwurf/Freigegebene Dokumente/Serviceberichte";

  async function uploadPdfToLibrary(fileName, blob, folder){
    var digest = await getDigest();
    var folderPath = folder || PDF_FOLDER;
    var arrayBuffer = await blob.arrayBuffer();
    var resp = await fetch(
      SITE_URL + "/_api/web/GetFolderByServerRelativeUrl('" + encodeURIComponent(folderPath) + "')/Files/add(url='" + encodeURIComponent(fileName) + "',overwrite=true)",
      {
        method: "POST",
        headers: {
          "Accept":          "application/json;odata=noop",
          "X-RequestDigest": digest
        },
        credentials: "include",
        body: arrayBuffer
      }
    );
    if(!resp.ok){
      var errText = "";
      try{ errText = await resp.text(); }catch(e){}
      throw new Error("SP PDF Upload: HTTP " + resp.status + " – " + errText.substring(0, 200));
    }
    var data = await resp.json();
    if(!data || !data.ServerRelativeUrl){
      throw new Error("SP PDF Upload: Antwort enthält keine ServerRelativeUrl. Response: " + JSON.stringify(data).substring(0, 200));
    }
    return {
      serverRelativeUrl: data.ServerRelativeUrl,
      name:              data.Name || "",
      length:            data.Length || 0
    };
  }

  function blobToBase64(blob){
    return new Promise(function(resolve, reject){
      var reader = new FileReader();
      reader.onload = function(){ resolve(reader.result); };
      reader.onerror = function(){ reject(new Error("Blob→Base64 fehlgeschlagen")); };
      reader.readAsDataURL(blob);
    });
  }

  function base64ToBlob(dataUrl){
    var parts = dataUrl.split(",");
    var mime = parts[0].match(/:(.*?);/)[1];
    var raw = atob(parts[1]);
    var arr = new Uint8Array(raw.length);
    for(var i = 0; i < raw.length; i++) arr[i] = raw.charCodeAt(i);
    return new Blob([arr], { type: mime });
  }

  /* ── OFFLINE-QUEUE ──────────────────────────────────────────────────── */

  function getQueue(){
    try{
      return JSON.parse(localStorage.getItem(QUEUE_KEY) || "[]");
    }catch(e){ return []; }
  }

  function saveQueue(q){
    localStorage.setItem(QUEUE_KEY, JSON.stringify(q));
  }

  // Eintrag zur Queue hinzufügen (verhindert Duplikate pro jobId)
  function enqueue(entry){
    var q = getQueue();
    var idx = q.findIndex(function(e){
      return e.jobId === entry.jobId && e.type === entry.type;
    });

    if(idx >= 0){
      var oldEntry = q[idx];
      if(oldEntry && oldEntry.pdfIdb && _pdfKey(oldEntry) !== _pdfKey(entry)){
        idbDeletePdf(_pdfKey(oldEntry)).catch(function(err){
          console.warn("Altes IndexedDB-PDF konnte nicht gelöscht werden:", err && err.message || err);
        });
      }
      q[idx] = entry;
    }else{
      q.push(entry);
    }

    saveQueue(q);
  }

  // Einzelnen Queue-Eintrag verarbeiten
  async function processEntry(entry){
    // 0. PDF hochladen — aus IndexedDB (pdfIdb) oder Legacy-Base64 direkt im Entry
    var _hasPdfIdb = entry.pdfIdb && entry.pdfFileName;
    var _hasPdfBase64 = !entry.pdfIdb && entry.pdfBase64 && entry.pdfFileName;

    if(_hasPdfIdb || _hasPdfBase64){
      var _pdfBase64Raw = null;

      if(_hasPdfIdb){
        _pdfBase64Raw = await idbLoadPdf(_pdfKey(entry));
        if(!_pdfBase64Raw){
          throw new Error(
            "PDF-Daten für Job " + entry.jobId + " nicht in IndexedDB gefunden " +
            "(möglicherweise Browser-Daten gelöscht). Bitte Rapport erneut öffnen und abschliessen."
          );
        }
      }else{
        _pdfBase64Raw = entry.pdfBase64;
      }

      var _pdfBlob = base64ToBlob(_pdfBase64Raw);
      var pdfResult = await uploadPdfToLibrary(entry.pdfFileName, _pdfBlob);

      if(!pdfResult || !pdfResult.serverRelativeUrl){
        throw new Error("PDF-Upload: Keine gültige serverRelativeUrl erhalten für Job " + entry.jobId);
      }
      if(!pdfResult.name){
        throw new Error("PDF-Upload: Kein Dateiname in der Antwort für Job " + entry.jobId);
      }

      if(entry.closeRun){
        entry.closeRun.PdfDateiname         = pdfResult.name;
        entry.closeRun.PdfServerRelativeUrl = pdfResult.serverRelativeUrl;
        entry.closeRun.PdfLink              = SITE_URL + "/_layouts/15/download.aspx?SourceUrl=" + encodeURIComponent(pdfResult.serverRelativeUrl);
        entry.closeRun.ReportTyp            = entry.type === "regie" ? "Regie" : "Checkliste";
        entry.closeRun.SourceApp            = "zentex-pwa";
        entry.closeRun.AppVersion           = "1.0";
        entry.closeRun.FinalizeStatus       = "Ready";
      }

      // 🔧 FIX 2 – IDB-PDF wird NICHT mehr hier gelöscht.
      // Das Löschen erfolgt erst ganz am Ende, nachdem alle SP-Writes erfolgreich waren.
    }

    // 1. Haupteintrag (Sprinkler_Jobs oder Sprinkler_Regie)
    var mainList = entry.type === "regie" ? LIST_REGIE : LIST_JOBS;
    await upsertItem(mainList, "JobId", entry.jobId, entry.mainFields);

    // 2. Checklisten-Items (nur bei Checkliste)
    if(entry.type === "checkliste" && Array.isArray(entry.items)){
      for(var i = 0; i < entry.items.length; i++){
        // Upsert pro Item (Title = JobId_ItemNr)
        await upsertItem(LIST_JOBITEMS, "Title", entry.items[i].Title, entry.items[i]);
      }
    }

    // 3. JobCloseRuns (Audit-Log) – 🔧 FIX 4: upsert statt createItem
    if(entry.closeRun){
      await upsertItem(LIST_CLOSERUNS, "RunId", entry.closeRun.RunId, entry.closeRun);
    }

    // 🔧 FIX 2 – IDB-PDF erst jetzt löschen, wenn alle SP-Writes erfolgreich waren
    if(_hasPdfIdb){
      await idbDeletePdf(_pdfKey(entry)).catch(function(err){
        console.warn("IDB-PDF nach erfolgreichem Sync löschen fehlgeschlagen:", err && err.message || err);
      });
    }
  }

  // Alle wartenden Queue-Einträge abarbeiten
  async function syncQueue(){
    if(!navigator.onLine) return { synced: 0, failed: 0, remaining: getQueue().length, syncedJobIds: [] };

    var q = getQueue();
    if(!q.length) return { synced: 0, failed: 0, remaining: 0, syncedJobIds: [] };

    var synced  = 0;
    var failed  = 0;
    var remaining = [];
    var syncedJobIds = [];

    for(var i = 0; i < q.length; i++){
      try{
        await processEntry(q[i]);
        synced++;
        if(q[i].jobId) syncedJobIds.push(q[i].jobId);
      }catch(e){
        console.warn("Sync fehlgeschlagen für " + q[i].jobId + ":", e.message);
        // Retry-Counter erhöhen
        q[i].retries = (q[i].retries || 0) + 1;
        q[i].lastError = e.message;
        q[i].lastRetry = Date.now();
        remaining.push(q[i]);
        failed++;
      }
    }

    saveQueue(remaining);
    return { synced: synced, failed: failed, remaining: remaining.length, syncedJobIds: syncedJobIds };
  }

  // Anzahl wartende Einträge
  function queueCount(){
    return getQueue().length;
  }

  // 🔧 FIX 1 – Einzelnen Job aus der echten Upload-Queue entfernen + IndexedDB-PDF löschen
  async function removeJobFromQueue(jobId){
    var q = getQueue();
    var removed = [];
    var kept = [];
    q.forEach(function(e){
      if(e && e.jobId === jobId){ removed.push(e); }
      else { kept.push(e); }
    });
    saveQueue(kept);
    for(var i = 0; i < removed.length; i++){
      if(removed[i].pdfIdb){
        await idbDeletePdf(_pdfKey(removed[i])).catch(function(err){
          console.warn("IDB-PDF löschen fehlgeschlagen:", err && err.message || err);
        });
      }
    }
    return removed.length;
  }

  // 🔧 FIX 1 – Komplette Queue leeren + alle IndexedDB-PDFs löschen
  async function clearAllQueued(){
    var q = getQueue();
    saveQueue([]);
    for(var i = 0; i < q.length; i++){
      if(q[i] && q[i].pdfIdb){
        await idbDeletePdf(_pdfKey(q[i])).catch(function(err){
          console.warn("IDB-PDF löschen fehlgeschlagen:", err && err.message || err);
        });
      }
    }
  }

  /* ── ONLINE/OFFLINE EVENT ───────────────────────────────────────────── */
  window.addEventListener("online", function(){
    // Automatisch Queue abarbeiten bei Online
    setTimeout(function(){
      syncQueue().then(function(r){
        if(r.synced > 0){
          console.log("Auto-Sync: " + r.synced + " Jobs synchronisiert.");
        }
      }).catch(function(e){
        console.warn("Auto-Sync Fehler:", e.message);
      });
    }, 2000); // 2s Verzögerung, damit Netzwerk stabil ist
  });

  /* ── CONVENIENCE: PAYLOAD BUILDER ───────────────────────────────────── */

  // Checkliste → Sprinkler_Jobs Felder
  function buildJobFields(payload){
    var m = payload.meta || {};
    var ansprechmail       = String(m.ansprechmail || "").trim();
    var emailKunde         = String(m.emailKunde || payload.emailKunde || "").trim();
    var auftragsleiterMail = String(m.auftragsleiterMail || "").trim();

    // Externe Empfänger: dedupliziert, ;-getrennt
    var extSet = {};
    if(ansprechmail) extSet[ansprechmail] = true;
    if(emailKunde)   extSet[emailKunde]   = true;
    var externStr = Object.keys(extSet).join("; ");

    return {
      Title:              payload.jobId || "",
      JobId:              payload.jobId || "",
      AuftragNr:          m.saznr || "",
      Objekt:             m.firma || m.objekt || "",
      Ort:                m.ort || "",
      SAZNr:              m.saznr || "",
      Anlageart:          m.anlageart || "",
      Ansprechperson:     m.ansprechperson || "",
      KundeEmail:         ansprechmail || emailKunde || "",
      MonteurName:        m.servicemonteur || "",
      Datum:              m.serviceDatum || null,
      Bemerkungen:        payload.finalNote || "",
      StatusJob:          "InArbeit",
      MailExternAn:       externStr,
      MailInternAn:       auftragsleiterMail,
      MailGesendet:       false,
      Ansprechmail:       ansprechmail,
      AuftragsleiterMail: auftragsleiterMail,
      AuftragsleiterName: m.auftragsleiterName || "",
      ClosedAt:           payload.closedAt || new Date().toISOString()
    };
  }

  // Checklisten-Items → Sprinkler_JobItems Felder
  function buildJobItems(payload){
    if(!payload.checklist) return [];
    var items = [];
    var keys  = Object.keys(payload.checklist);
    for(var i = 0; i < keys.length; i++){
      var nr = keys[i];
      var it = payload.checklist[nr];
      if(!it) continue;
      items.push({
        Title:       (payload.jobId || "") + "_" + nr,
        JobID:       payload.jobId || "",
        Bereich:     it.bereich || "",
        Pruefpunkt:  it.text || "",
        Status:      it.status || "",
        Antwort:     it.status || "",
        Bemerkung:   it.note || "",
        WertNotiz:   it.note || "",
        ItemOrder:   parseInt(nr, 10) || 0
      });
    }
    return items;
  }

  // Regie → Sprinkler_Regie Felder
  function buildRegieFields(data){
    var emailKundeVal  = String(data.emailKunde || "").trim();
    var alMail         = String(data.auftragsleiterMail || "").trim();
    return {
      Title:              data.jobId || data.id || "",
      RegieTitel:         (data.gebaeude || "") + " – " + (data.kunde || ""),
      JobId:              data.jobId || data.id || "",
      AuftragNr:          data.auftragNr || "",
      Kunde:              data.kunde || "",
      KundeEmail:         emailKundeVal,
      Monteur:            data.nameMonteur || "",
      Stunden:            parseFloat(data.totalStunden) || 0,
      Material:           data.material || "",
      Bemerkung:          data.bemerkungen || "",
      Status:             "InArbeit",
      MailGesendet:       false,
      MailExternAn:       emailKundeVal,
      MailInternAn:       alMail,
      AuftragsleiterMail: alMail,
      AuftragsleiterName: data.auftragsleiterName || "",
      ClosedAt:           new Date().toISOString()
    };
  }

  // JobCloseRuns (Audit-Log)
  function buildCloseRun(jobId, type, monteur, firma, saz, kundeEmail, internEmail){
    // 🔧 FIX 4 – RunId deterministisch: jobId + type → idempotenter Upsert möglich
    var runId = (jobId || "unknown") + "_" + (type || "job");
    return {
      Title:           runId,
      RunId:           runId,
      JobId:           jobId || "",
      Status:          "InArbeit",
      StartedAt:       new Date().toISOString(),
      Monteur:         monteur || "",
      Anlagenbesitzer: firma || "",
      SAZ:             saz || "",
      KundeEmail:      kundeEmail || "",
      BueroEmail:      internEmail || ""
    };
  }

  /* ── HAUPTFUNKTION: JOB ABSCHLIESSEN ────────────────────────────────── */

  // Kompletter Abschluss: Daten sammeln → SharePoint oder Queue
  async function submitJob(payload, type){
    var jobId = payload.jobId || payload.id || "";
    var entry;
    var pdfBlob = payload.pdfBlob || null;
    var pdfFileName = payload.pdfFileName || "";

    if(type === "regie"){
      entry = {
        jobId:      jobId,
        type:       "regie",
        mainFields: buildRegieFields(payload),
        items:      null,
        closeRun:   buildCloseRun(
          jobId, "regie",
          payload.nameMonteur || "",
          payload.gebaeude || payload.kunde || "",
          payload.auftragNr || "",
          String(payload.emailKunde || "").trim(),
          String(payload.auftragsleiterMail || "").trim()
        ),
        queuedAt:   Date.now()
      };
    }else{
      var m = payload.meta || {};
      entry = {
        jobId:      jobId,
        type:       "checkliste",
        mainFields: buildJobFields(payload),
        items:      buildJobItems(payload),
        closeRun:   buildCloseRun(
          jobId, "checkliste",
          m.servicemonteur || "",
          m.firma || "",
          m.saznr || "",
          String(m.ansprechmail || "").trim(),
          String(m.auftragsleiterMail || "").trim()
        ),
        queuedAt:   Date.now()
      };
    }

    if(pdfBlob && pdfFileName){
      if(navigator.onLine){
        try{
          var pdfResult = await uploadPdfToLibrary(pdfFileName, pdfBlob);
          if(!pdfResult || !pdfResult.serverRelativeUrl){
            throw new Error("PDF-Upload: Keine gültige URL erhalten");
          }
          if(!pdfResult.name){
            throw new Error("PDF-Upload: Kein Dateiname in der Antwort erhalten");
          }
          entry.closeRun.PdfDateiname         = pdfResult.name;
          entry.closeRun.PdfServerRelativeUrl = pdfResult.serverRelativeUrl;
          entry.closeRun.PdfLink              = SITE_URL + "/_layouts/15/download.aspx?SourceUrl=" + encodeURIComponent(pdfResult.serverRelativeUrl);
          entry.closeRun.ReportTyp            = type === "regie" ? "Regie" : "Checkliste";
          entry.closeRun.SourceApp            = "zentex-pwa";
          entry.closeRun.AppVersion           = "1.0";
          entry.closeRun.FinalizeStatus       = "Ready";
        }catch(uploadErr){
          console.warn("PDF-Upload fehlgeschlagen, wird in Queue gespeichert:", uploadErr.message);
          try{
            var b64 = await blobToBase64(pdfBlob);
            var savedFb = await idbSavePdf(_pdfKey(entry), b64);
            if(savedFb){
              entry.pdfIdb      = true;
              entry.pdfFileName = pdfFileName;
            }else{
              entry.pdfBase64   = b64;
              entry.pdfFileName = pdfFileName;
            }
          }catch(b64Err){
            // 🔧 FIX 3 – PDF-Sicherung fehlgeschlagen: nicht als queued melden
            console.warn("PDF-Fallback-Speichern fehlgeschlagen:", b64Err.message);
            throw new Error(
              "PDF konnte weder hochgeladen noch lokal gespeichert werden. " +
              "Bitte Speicherplatz und Verbindung prüfen. (" + b64Err.message + ")"
            );
          }
          // Nur queuen, wenn PDF-Daten sicher vorhanden sind
          if(!entry.pdfIdb && !entry.pdfBase64){
            throw new Error(
              "PDF-Daten fehlen im Queue-Eintrag. Rapport nicht abgeschlossen."
            );
          }
          enqueue(entry);
          return { ok: true, mode: "queued", jobId: jobId, error: uploadErr.message };
        }
      }else{
        try{
          var b64offline = await blobToBase64(pdfBlob);
          var saved = await idbSavePdf(_pdfKey(entry), b64offline);
          if(saved){
            entry.pdfIdb      = true;
            entry.pdfFileName = pdfFileName;
          }else{
            entry.pdfBase64   = b64offline;
            entry.pdfFileName = pdfFileName;
          }
        }catch(b64Err2){
          // 🔧 FIX 3 – Offline-PDF-Sicherung fehlgeschlagen: nicht still weitermachen
          console.warn("PDF offline speichern fehlgeschlagen:", b64Err2.message);
          throw new Error(
            "PDF konnte nicht lokal gespeichert werden (offline). " +
            "Bitte Gerätespeicher prüfen. (" + b64Err2.message + ")"
          );
        }
        // Sicherheitscheck: PDF-Daten müssen im Entry vorhanden sein
        if(!entry.pdfIdb && !entry.pdfBase64){
          throw new Error("PDF-Daten fehlen im Queue-Eintrag (offline). Rapport nicht abgeschlossen.");
        }
      }
    }

    if(navigator.onLine){
      try{
        await processEntry(entry);
        return { ok: true, mode: "online", jobId: jobId };
      }catch(e){
        console.warn("Online-Upload fehlgeschlagen, in Queue gelegt:", e.message);
        if(pdfBlob && !entry.pdfBase64 && !entry.pdfIdb){
          try{
            var b64q = await blobToBase64(pdfBlob);
            var savedQ = await idbSavePdf(_pdfKey(entry), b64q);
            if(savedQ){
              entry.pdfIdb      = true;
              entry.pdfFileName = pdfFileName;
            }else{
              entry.pdfBase64   = b64q;
              entry.pdfFileName = pdfFileName;
            }
          }catch(e2){}
        }
        enqueue(entry);
        return { ok: true, mode: "queued", jobId: jobId, error: e.message };
      }
    }else{
      enqueue(entry);
      return { ok: true, mode: "queued", jobId: jobId };
    }
  }

  /* ── EXPORT ─────────────────────────────────────────────────────────── */
  window.ZentexSP = {
    // Konfiguration
    SITE_URL:  SITE_URL,

    // Hauptfunktionen
    submitJob:  submitJob,
    syncQueue:  syncQueue,
    queueCount: queueCount,

    // 🔧 FIX 1 – Queue-Cleanup-Methoden (für offline.html deleteJob / confirmClear)
    removeJobFromQueue: removeJobFromQueue,
    clearAllQueued:     clearAllQueued,

    // Low-Level (für spezielle Fälle)
    getDigest:   getDigest,
    createItem:  createItem,
    updateItem:  updateItem,
    findItem:    findItemByField,
    upsertItem:  upsertItem,
    uploadPdfToLibrary: uploadPdfToLibrary,

    // Payload Builder (für Tests/Debugging)
    buildJobFields:   buildJobFields,
    buildRegieFields: buildRegieFields,
    buildJobItems:    buildJobItems,
    buildCloseRun:    buildCloseRun,

    // Hilfsfunktionen
    blobToBase64:  blobToBase64,
    base64ToBlob:  base64ToBlob,
    idbSavePdf:    idbSavePdf,
    idbLoadPdf:    idbLoadPdf,
    idbDeletePdf:  idbDeletePdf
  };

})(window);