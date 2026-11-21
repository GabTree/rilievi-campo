/* Rilievi di Campo – importazione di piante da un elenco (Excel, CSV, testo incollato, KML/KMZ di Google My Maps, ZIP)
   Ogni riga diventa una pianta "da fare" (da completare in campo) oppure un rilievo già completo.
   Le colonne si riconoscono dal titolo (prima riga); nell'anteprima si possono correggere a mano. */
"use strict";

const Importer = (() => {
  // campi dell'app e parole con cui li riconosciamo nei titoli delle colonne
  const FIELDS = [
    ["", "— ignora —"],
    ["progetto", "Progetto", /progett|cliente|commessa|^lavoro|localit/],
    ["data", "Data valutazione", /^data|^date/],
    ["protocollo", "Protocollo (SIA / ARETÉ)", /protocoll/],
    ["cpc", "Nuova CPC", /\bcpc\b|propensione|classe/],
    ["ngs", "Numero gs", /\bgs\b|greenspace|green space/],
    ["ncart", "Numero cartellino", /cartell|\bcart\b|targhett|medagli/],
    ["nprog", "Progressivo", /progress|^n\.?$|^nr\.?$|^num|^numero|pianta|albero|^id$|^nome$|^name$/],
    ["specie", "Specie", /specie|essenza|species|nome scient|genere/],
    ["diam", "Diametro a 120 cm", /diam|dbh/],
    ["alt", "Altezza", /altez|height|^h$|^h \(m\)/],
    ["coord", "Coordinate / link Maps", /coord|posizion|link|maps|url|gps|^punto|wkt/],
    ["lat", "Latitudine", /^lat|latitud|^y$/],
    ["lon", "Longitudine", /^lon|^lng|longitud|^x$/],
    ["prob", "Note di campo", /note di campo|problemat|osservaz|difett|^note$|descri/],
    ["int", "Interventi", /interv/],
    ["strum", "Tipo di strumentale", /tipo di strument|^strument|prova strum|^prova$/],
    ["esito", "Esito della prova", /esito/],
    ["strumNote", "Note strumentale", /note strument/],
    ["pot", "Potatura (note)", /potatur/],
    ["scavo", "Scavo (note)", /scav|sorveglianza/],
    ["altro", "Altro (note)", /^altro/],
    ["foto", "Foto (nomi dei file)", /foto|immagin|photo|picture/],
  ];
  const IMG = /\.(jpe?g|png|heic|heif|webp)$/i;
  let rows = [], header = [], map = [], images = [], srcName = "", startProject = "";

  /* ------------------------------------------------------------ lettura dei file */
  const norm = (s) => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
  const xmlText = (el) => (el ? el.textContent : "");

  async function readXlsx(buf) {
    const zip = await JSZip.loadAsync(buf);
    const P = new DOMParser();
    const xml = async (path) => { const f = zip.file(path); return f ? P.parseFromString(await f.async("string"), "application/xml") : null; };
    const wb = await xml("xl/workbook.xml"), rels = await xml("xl/_rels/workbook.xml.rels");
    const first = wb.getElementsByTagName("sheet")[0];
    const rid = first.getAttribute("r:id") || first.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    let target = [...rels.getElementsByTagName("Relationship")].find((r) => r.getAttribute("Id") === rid).getAttribute("Target");
    target = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
    const sst = await xml("xl/sharedStrings.xml");
    const shared = sst ? [...sst.getElementsByTagName("si")].map((si) => [...si.getElementsByTagName("t")].map(xmlText).join("")) : [];
    // stili che sono date (per convertire i numeri seriali di Excel)
    const st = await xml("xl/styles.xml"), dateStyles = new Set();
    if (st) {
      const custom = {};
      [...st.getElementsByTagName("numFmt")].forEach((n) => (custom[n.getAttribute("numFmtId")] = n.getAttribute("formatCode")));
      const xfs = st.getElementsByTagName("cellXfs")[0];
      if (xfs) [...xfs.getElementsByTagName("xf")].forEach((x, i) => {
        const id = +x.getAttribute("numFmtId");
        if ((id >= 14 && id <= 22) || (custom[id] && /[dmy]/i.test(custom[id].replace(/\[[^\]]*\]|"[^"]*"/g, "")))) dateStyles.add(i);
      });
    }
    const sh = await xml(target), out = [];
    const col = (ref) => { let n = 0; for (const c of ref.replace(/\d+/g, "")) n = n * 26 + c.charCodeAt(0) - 64; return n - 1; };
    [...sh.getElementsByTagName("row")].forEach((row) => {
      const line = [];
      [...row.getElementsByTagName("c")].forEach((c) => {
        const t = c.getAttribute("t"), v = xmlText(c.getElementsByTagName("v")[0]);
        let val;
        if (t === "s") val = shared[+v] ?? "";
        else if (t === "inlineStr") val = [...c.getElementsByTagName("t")].map(xmlText).join("");
        else if (t === "b") val = v === "1" ? "sì" : "";
        else if (v !== "" && dateStyles.has(+c.getAttribute("s"))) val = serialToIso(+v);
        else val = v;
        line[col(c.getAttribute("r"))] = val;
      });
      out.push(Array.from(line, (x) => (x === undefined ? "" : String(x))));
    });
    return out;
  }
  function serialToIso(n) {
    const d = new Date(Math.round((n - 25569) * 86400000));
    return isNaN(d) ? String(n) : d.toISOString().slice(0, 10);
  }
  function readDelimited(text) {
    text = text.replace(/^﻿/, "");
    const firstLine = text.split(/\r?\n/)[0];
    const sep = firstLine.includes("\t") ? "\t" : (firstLine.split(";").length >= firstLine.split(",").length ? ";" : ",");
    const out = []; let row = [], cell = "", q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
      else if (ch === '"') q = true;
      else if (ch === sep) { row.push(cell); cell = ""; }
      else if (ch === "\n" || ch === "\r") { if (ch === "\r" && text[i + 1] === "\n") i++; row.push(cell); out.push(row); row = []; cell = ""; }
      else cell += ch;
    }
    if (cell !== "" || row.length) { row.push(cell); out.push(row); }
    return out.filter((r) => r.some((c) => c.trim() !== ""));
  }
  function readKml(text) {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    const out = [["Nome", "Note", "Coordinate"]];
    [...doc.getElementsByTagName("Placemark")].forEach((pm) => {
      const pt = pm.getElementsByTagName("Point")[0];
      if (!pt) return;
      const c = xmlText(pt.getElementsByTagName("coordinates")[0]).trim().split(/[\s]+/)[0].split(",");
      const desc = xmlText(pm.getElementsByTagName("description")[0]).replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
      out.push([xmlText(pm.getElementsByTagName("name")[0]).trim(), desc, `${c[1]}, ${c[0]}`]);
    });
    return out;
  }
  async function readAny(file) {
    const n = file.name.toLowerCase();
    if (n.endsWith(".xlsx") || n.endsWith(".xlsm")) return readXlsx(await file.arrayBuffer());
    if (n.endsWith(".kml")) return readKml(await file.text());
    if (n.endsWith(".kmz")) {
      const z = await JSZip.loadAsync(await file.arrayBuffer());
      const k = Object.values(z.files).find((f) => f.name.toLowerCase().endsWith(".kml"));
      if (!k) throw new Error("Nel KMZ non c'è un file KML");
      return readKml(await k.async("string"));
    }
    if (n.endsWith(".xls")) throw new Error("Il vecchio formato .xls non è leggibile: in Excel fai «Salva con nome» → .xlsx");
    return readDelimited(await file.text());
  }

  /* --------------------------------------------------------------- coordinate */
  // accetta: "41.9128, 12.4852" · "41,9128; 12,4852" · link di Google Maps · gradi-minuti-secondi 41°54'46"N 12°29'07"E
  function parseCoord(s) {
    s = String(s || "").trim();
    if (!s) return null;
    if (/maps\.app\.goo\.gl|goo\.gl\/maps/i.test(s)) return { err: "link breve" };
    let m = s.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || s.match(/[?&](?:q|query|ll|destination|daddr|center)=(-?\d+\.\d+)\s*(?:,|%2C)\s*(-?\d+\.\d+)/i) || s.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
    if (m) return check(+m[1], +m[2]);
    const dms = [...s.matchAll(/(\d+(?:[.,]\d+)?)\s*°\s*(\d+(?:[.,]\d+)?)?\s*['′]?\s*(\d+(?:[.,]\d+)?)?\s*(?:"|″|'')?\s*([NSEWO])/gi)];
    if (dms.length >= 2) {
      const v = dms.map((d) => { const x = +d[1].replace(",", ".") + (+(d[2] || "0").replace(",", ".")) / 60 + (+(d[3] || "0").replace(",", ".")) / 3600; return { x: /[SWO]/i.test(d[4]) ? -x : x, h: d[4].toUpperCase() }; });
      const lat = v.find((d) => /[NS]/.test(d.h)), lon = v.find((d) => /[EWO]/.test(d.h));
      if (lat && lon) return check(lat.x, lon.x);
    }
    const nums = s.match(/-?\d+(?:[.,]\d+)?/g);
    if (nums && nums.length >= 2) {
      // "41,9128, 12,4852" (virgole decimali e separatore) → prendo coppie con il punto, altrimenti le prime due
      const clean = s.match(/-?\d+\.\d+/g);
      const pair = clean && clean.length >= 2 ? clean : nums.map((x) => x.replace(",", "."));
      return check(+pair[0], +pair[1]);
    }
    return null;
  }
  function check(a, b) {
    if (!isFinite(a) || !isFinite(b)) return null;
    // in Italia la latitudine è 35–48 e la longitudine 6–19: se sono invertite le scambio
    if (a >= 6 && a <= 19 && b >= 35 && b <= 48) [a, b] = [b, a];
    if (Math.abs(a) > 90 || Math.abs(b) > 180) return null;
    return { lat: a, lon: b };
  }
  function parseDate(s) {
    s = String(s || "").trim();
    if (!s) return "";
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
    if (m) { const y = m[3].length === 2 ? "20" + m[3] : m[3]; return `${y}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`; }
    if (/^\d{5}(\.\d+)?$/.test(s)) return serialToIso(+s);
    return "";
  }
  const cleanNum = (s) => String(s || "").trim().replace(/^(gs|cart|rel|n)\.?\s*/i, "").replace(/\.0$/, "");

  /* ------------------------------------------------------------- dalle righe ai dati */
  function guessMap(h) {
    const used = new Set();
    return h.map((t) => {
      const n = norm(t);
      if (!n) return "";
      for (const [k, , re] of FIELDS) if (k && !used.has(k) && re.test(n)) { used.add(k); return k; }
      return "";
    });
  }
  function rowData(r) {
    const d = {};
    map.forEach((k, i) => { if (k && r[i] !== undefined && String(r[i]).trim() !== "") d[k] = String(r[i]).trim(); });
    let pos = null;
    if (d.coord) pos = parseCoord(d.coord);
    if (!pos && d.lat && d.lon) pos = check(+String(d.lat).replace(",", "."), +String(d.lon).replace(",", "."));
    return { d, pos };
  }
  // foto: per nome indicato nella colonna Foto, altrimenti per numero all'inizio del nome del file (es. 4521_1.jpg)
  function matchPhotos(d, used) {
    const base = (n) => norm(n).replace(/\.[a-z0-9]+$/, "");
    if (d.foto) {
      const want = d.foto.split(/[;\n,]+/).map((x) => base(x.trim())).filter(Boolean);
      return images.filter((f) => !used.has(f) && want.includes(base(f.name)));
    }
    const nums = [d.ngs, d.ncart, d.nprog].map(cleanNum).filter(Boolean);
    if (!nums.length) return [];
    return images.filter((f) => !used.has(f) && nums.some((n) => new RegExp(`^(gs|cart|rel)?\\.?${n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}([_\\-\\s.(]|$)`, "i").test(base(f.name))))
      .sort((a, b) => a.name.localeCompare(b.name, "it", { numeric: true }));
  }

  /* --------------------------------------------------------------------- interfaccia */
  function open(project) {
    startProject = project || "";
    rows = []; header = []; map = []; images = []; srcName = "";
    $("#impFile").value = ""; $("#impPaste").value = ""; $("#impPreview").innerHTML = ""; $("#impMsg").textContent = "";
    $("#impProject").value = startProject;
    $("#impStep2").classList.add("hidden");
    setSeg("segImpMode", "todo");
    show("imp");
  }
  async function loadFiles(files) {
    $("#impMsg").textContent = "Leggo…";
    try {
      let table = null; images = [];
      for (const f of files) {
        if (/\.zip$/i.test(f.name)) {
          const z = await JSZip.loadAsync(await f.arrayBuffer());
          for (const e of Object.values(z.files)) {
            if (e.dir || /(^|\/)(__MACOSX|\.)/.test(e.name)) continue;
            const name = e.name.split("/").pop();
            const blob = await e.async("blob");
            if (IMG.test(name)) images.push(new File([blob], name, { type: /png$/i.test(name) ? "image/png" : "image/jpeg" }));
            else if (!table && /\.(xlsx|xlsm|csv|txt|kml|kmz)$/i.test(name)) { table = await readAny(new File([blob], name)); srcName = name; }
          }
        } else if (IMG.test(f.name) || /^image\//.test(f.type)) images.push(f);
        else if (!table) { table = await readAny(f); srcName = f.name; }
      }
      if (!table) { $("#impMsg").textContent = images.length ? "Hai scelto solo foto: scegli anche il file Excel (insieme alle foto)." : "Nessun elenco trovato."; return; }
      setTable(table);
    } catch (e) { $("#impMsg").textContent = "Non riesco a leggere il file: " + (e.message || e); }
  }
  function setTable(table) {
    const hi = table.findIndex((r) => r.filter((c) => String(c).trim()).length >= 2);
    if (hi < 0) { $("#impMsg").textContent = "Il file sembra vuoto."; return; }
    header = table[hi].map((c) => String(c).trim());
    rows = table.slice(hi + 1).filter((r) => r.some((c) => String(c).trim() !== ""));
    map = guessMap(header);
    if (!$("#impProject").value.trim()) {
      const pi = map.indexOf("progetto"), fromCol = pi >= 0 && rows.find((r) => String(r[pi] || "").trim());
      $("#impProject").value = fromCol ? String(fromCol[pi]).trim() : srcName.replace(/\.[^.]+$/, "").replace(/[_]+/g, " ");
    }
    $("#impMsg").textContent = "";
    $("#impStep2").classList.remove("hidden");
    preview();
  }
  function preview() {
    const used = new Set();
    let ok = 0, noPos = 0, noNum = 0, nFoto = 0;
    const info = rows.map((r) => {
      const { d, pos } = rowData(r);
      const ph = matchPhotos(d, used); ph.forEach((f) => used.add(f)); nFoto += ph.length;
      const num = d.ngs || d.ncart || d.nprog;
      if (!num) noNum++; else ok++;
      if (!pos) noPos++;
      return { d, pos, ph };
    });
    const opts = (sel) => FIELDS.map(([k, l]) => `<option value="${k}"${k === sel ? " selected" : ""}>${l}</option>`).join("");
    const cols = header.map((h, i) => `<div class="imp-col"><span>${esc(h || "(colonna " + (i + 1) + ")")}</span>
      <select data-i="${i}">${opts(map[i])}</select></div>`).join("");
    const sample = info.slice(0, 5).map(({ d, pos, ph }) => `<tr><td>${esc([d.ngs && "gs." + cleanNum(d.ngs), d.ncart && "cart." + cleanNum(d.ncart), d.nprog && "rel." + cleanNum(d.nprog)].filter(Boolean).join(" ") || "—")}</td>
      <td>${esc(d.specie || "")}</td><td>${pos && !pos.err ? pos.lat.toFixed(5) + ", " + pos.lon.toFixed(5) : '<span class="warn">—</span>'}</td><td>${ph.length || ""}</td></tr>`).join("");
    const shortLinks = info.filter((x) => x.pos && x.pos.err).length;
    $("#impPreview").innerHTML = `
      <label class="f">Colonne riconosciute <span class="muted">(cambia se serve)</span></label><div class="imp-cols">${cols}</div>
      <label class="f">Anteprima</label>
      <table class="imp-tab"><tr><th>Numeri</th><th>Specie</th><th>Coordinate</th><th>Foto</th></tr>${sample}</table>
      <div class="muted">${rows.length} righe${rows.length > 5 ? " (mostrate le prime 5)" : ""} · ${images.length ? nFoto + " foto abbinate su " + images.length : "nessuna foto"}</div>
      ${noNum ? `<div class="warn">${noNum === 1 ? "1 riga" : noNum + " righe"} senza numero (gs, cartellino o progressivo): ${noNum === 1 ? "verrà saltata" : "verranno saltate"}.</div>` : ""}
      ${noPos - shortLinks > 0 ? `<div class="warn">${noPos - shortLinks === 1 ? "1 riga" : noPos - shortLinks + " righe"} senza coordinate leggibili: non comparirà sulla mappa.</div>`.replace(noPos - shortLinks === 1 ? "@" : "comparirà", "compariranno") : ""}
      ${shortLinks ? `<div class="warn">${shortLinks} link brevi (maps.app.goo.gl) non leggibili: aprili e copia le coordinate.</div>` : ""}`;
    $("#impPreview").querySelectorAll("select").forEach((s) => s.addEventListener("change", () => { map[+s.dataset.i] = s.value; preview(); }));
    $("#btnImpGo").textContent = `Importa ${ok} ${ok === 1 ? "pianta" : "piante"}`;
    $("#btnImpGo").disabled = !ok;
  }
  async function go() {
    const project = $("#impProject").value.trim();
    if (!project && map.indexOf("progetto") < 0) { $("#impMsg").textContent = "Scrivi il nome del progetto"; return; }
    if (!S.name || !S.init) { toast("Prima inserisci nome e iniziali del rilevatore (Impostazioni)"); return; }
    const mode = $("#segImpMode button.on")?.dataset.v || "todo";
    const btn = $("#btnImpGo"); btn.disabled = true;
    const used = new Set(), names = new Set((await DB.all()).map((x) => x.cartella).filter(Boolean));
    let n = 0, skip = 0;
    const projects = new Set();
    for (const r of rows) {
      const { d, pos } = rowData(r);
      if (!(d.ngs || d.ncart || d.nprog)) { skip++; continue; }
      const ph = matchPhotos(d, used); ph.forEach((f) => used.add(f));
      const rec = newDraft();
      const tipo = "Valutazione";
      const prot = /aret/i.test(d.protocollo || "") ? "ARETE" : "SIA";
      const cpcM = String(d.cpc || "").toUpperCase().replace(/\s+/g, "").match(/^(C\/D|A|B|C|D)$/);
      const esito = (d.esito || "").toLowerCase().startsWith("a") ? "Alto" : (d.esito || "").toLowerCase().startsWith("m") ? "Medio" : (d.esito || "").toLowerCase().startsWith("b") ? "Basso" : "";
      Object.assign(rec, {
        lavoro: (map.indexOf("progetto") >= 0 && d.progetto) || project, dataValutazione: parseDate(d.data) || (mode === "todo" ? "" : isoToday()),
        tipoValutazione: tipo, ngs: cleanNum(d.ngs), ncart: cleanNum(d.ncart), nprog: cleanNum(d.nprog), specie: d.specie || "",
        diametro120: (d.diam || "").replace(",", "."), altezza: (d.alt || "").replace(",", "."), problematiche: d.prob || "", note: "",
        interventi: d.int || "", strumento: d.strum || "", esitoProva: esito, strumNote: d.strumNote || "",
        protocollo: prot, cpc: prot === "SIA" && cpcM ? cpcM[1] : "",
        potOn: !!d.pot, potNote: d.pot && !/^(s[iì]|x|1)$/i.test(d.pot) ? d.pot : "",
        scavoOn: !!d.scavo, scavoNote: d.scavo && !/^(s[iì]|x|1)$/i.test(d.scavo) ? d.scavo : "",
        altroOn: !!d.altro, altroNote: d.altro && !/^(s[iì]|x|1)$/i.test(d.altro) ? d.altro : "",
        operatore: S.name, iniziali: (S.init || "XX").toUpperCase(), appVersion: APP_VERSION, importato: srcName || "testo incollato",
      });
      rec.strumOn = !!(rec.strumento || rec.esitoProva || rec.strumNote);
      rec.posizione = pos && !pos.err ? { lat: pos.lat, lon: pos.lon, acc: null, importata: true } : null;
      btn.textContent = `Importo ${n + 1}…`;
      for (const f of ph) {
        try {
          const bmp = await createImageBitmap(f, { imageOrientation: "from-image" });
          const meta = { time: f.lastModified || Date.now(), lat: rec.posizione?.lat, lon: rec.posizione?.lon, acc: null, heading: null, ref: "T", arrow: false, compassBad: false };
          const p = makePhotos(bmp, bmp.width, bmp.height, meta);
          rec.foto.push({ ...meta, clean: p.clean, arrow: null, thumb: p.thumb, origName: f.name });
        } catch (e) { console.warn("foto non leggibile", f.name, e); }
      }
      if (mode === "todo") { rec.status = "todo"; rec.cartella = ""; }
      else {
        rec.status = "pending";
        const base = nomeRilievo(rec); let name = base, k = 2;
        while (names.has(name)) name = `${base}-${k++}`;
        names.add(name); rec.cartella = name;
        rec.foto.forEach((p, i) => { const s = String(i + 1).padStart(2, "0"); p.name = `foto_${s}.jpg`; p.nameN = null; });
      }
      await DB.put(rec); n++; projects.add(rec.lavoro);
    }
    btn.disabled = false;
    toast(`${n} ${mode === "todo" ? "piante da fare" : "rilievi"} importati${skip ? ` · ${skip} righe saltate` : ""}`, 4000);
    refreshCounts(); MapView.plot();
    if (mode !== "todo") Sync.run();
    const first = [...projects][0];
    if (projects.size === 1) openProject(first); else show("list");
  }
  function bind() {
    $("#impFile").addEventListener("change", (e) => e.target.files.length && loadFiles([...e.target.files]));
    $("#btnImpPaste").addEventListener("click", () => {
      const t = $("#impPaste").value;
      if (!t.trim()) { $("#impMsg").textContent = "Incolla prima le righe copiate da Excel (con la riga dei titoli)"; return; }
      srcName = ""; images = []; setTable(readDelimited(t));
    });
    bindSeg("segImpMode", () => {});
    $("#btnImpGo").addEventListener("click", go);
    $("#btnImpBack").addEventListener("click", () => (startProject ? openProject(startProject) : show("list")));
  }
  return { open, bind, _parseCoord: parseCoord, _readDelimited: readDelimited };
})();
Importer.bind();
