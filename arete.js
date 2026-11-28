/* Scheda ARETÉ: sezione della scheda di rilievo e generazione del file Excel compilato.
 * Il modello (Scheda_ARETE_ORD_telefono.xlsx, fogli "ORD" e "A") lo carica ogni utente:
 * non è incluso nell'app perché la scheda ARETÉ è distribuita con licenza CC BY-NC-ND. */
"use strict";

const ARETE = (() => {
  let L = null;  // elenchi (arete_lists.json)
  const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

  const GEN = [
    { k: "dimora", cell: "C7", label: "Dimora", list: "dimora" },
    { k: "localiz", cell: "G7", label: "Localizzazione", list: "localiz" },
    { k: "psoc", cell: "J7", label: "Posizione sociale", list: "psoc" },
    { k: "stadio", cell: "N7", label: "Stadio di sviluppo", list: "stadio" },
    { k: "vincoli", cell: "S7", label: "Vincoli", list: "vincoli" },
    { k: "ubicaz", cell: "C8", label: "Ubicazione (provincia)", list: "ubicaz", remember: true },
    { k: "confl1", cell: "L8", label: "Conflitto 1", list: "conflitti" },
    { k: "confl2", cell: "N8", label: "Conflitto 2", list: "conflitti" },
    { k: "coerenza", cell: "S8", label: "Coerenza fito-climatica/geo", list: "coerenza" },
  ];
  const DIM = [
    { k: "Dch", cell: "J9", label: "D ch – diametro chioma (m)" },
    { k: "Dbr", cell: "M9", label: "D br – diametro branca (cm)" },
    { k: "Lbr", cell: "O9", label: "L br – lunghezza branca (m)" },
    { k: "Hbr", cell: "R9", label: "H br – altezza branca (m)" },
    { k: "Hbers", cell: "T9", label: "H bers – altezza bersaglio (m)" },
  ];
  const DIAG = [
    { k: "salute", cell: "E12", label: "Condizioni di salute e vigoria", list: "salute" },
    { k: "fitopatia", cell: "R12", label: "Fitopatia", list: "fitopatia", free: true },
    { k: "cariogeno", cell: "R13", label: "Agente cariogeno", list: "cariogeno", free: true },
  ];
  const range = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
  const ZONES = [
    { k: "zolla", label: "Zolla radicale", d: "B", g: "F", rows: range(15, 24) },
    { k: "colletto", label: "Colletto", d: "G", g: "K", rows: range(15, 24) },
    { k: "fusto", label: "Fusto", d: "L", g: "P", rows: range(15, 19) },
    { k: "castello", label: "Castello", d: "L", g: "P", rows: range(21, 24) },
    { k: "branche", label: "Branche e rami", d: "Q", g: "U", rows: range(15, 19) },
    { k: "chioma", label: "Chioma", d: "Q", g: "U", rows: range(21, 24) },
  ];
  const PER = ["Rottura branca / rami", "Rottura tronco / castello", "Rottura colletto", "Ribaltamento / scivolamento zolla"];
  const PRESC = [
    { k: "ic", col: "B", label: "Interventi colturali", list: "interventi" },
    { k: "icu", col: "J", label: "Urgenza", list: "urgenza" },
    { k: "mb", col: "L", label: "Mitigazione bersaglio", list: "mitigazione" },
    { k: "mbu", col: "O", label: "Urgenza", list: "urgenza" },
    { k: "pv", col: "Q", label: "Prescrizioni valutative", list: "valutative" },
    { k: "pvu", col: "T", label: "Urgenza / monitoraggio", list: "urgmonit" },
  ];
  const POSTDIM = [
    { k: "H", cell: "C47", label: "H – altezza (m)" }, { k: "Circ", cell: "H47", label: "Circ – circonferenza (cm)" },
    { k: "Dch", cell: "J47", label: "D ch (m)" }, { k: "Dbr", cell: "M47", label: "D br (cm)" },
    { k: "Lbr", cell: "O47", label: "L br (m)" }, { k: "Hbr", cell: "R47", label: "H br (m)" },
    { k: "Hbers", cell: "T47", label: "H bers (m)" },
  ];
  const BERS = {
    pre: { tA: "C32", fA: "B33", fArng: "$BQ$22:$BQ$28", m: "H32", tR: "N32", fR: "M33", fRrng: "$BQ$9:$BQ$15", la: "albero", lr: "rami" },
    post: { tA: "C51", fA: "B52", fArng: "$BQ$80:$BQ$86", tR: "N51", fR: "M52", fRrng: "$BQ$93:$BQ$99", la: "albero_post", lr: "rami_post" },
  };

  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const opts = (list, empty = true) => (empty ? '<option value=""></option>' : "") +
    (L[list] || []).map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");

  /* ------------------------------------------------------------ UI */
  function selectField(key, label, list, free) {
    if (free) {
      const id = "dl_" + list;
      if (!document.getElementById(id)) {
        const dl = document.createElement("datalist");
        dl.id = id;
        dl.innerHTML = (L[list] || []).map((v) => `<option value="${esc(v)}">`).join("");
        document.body.appendChild(dl);
      }
      return `<label class="f">${label}</label><input type="text" data-a="${key}" list="${id}" autocomplete="off">`;
    }
    return `<label class="f">${label}</label><select data-a="${key}">${opts(list)}</select>`;
  }
  const numField = (key, label, ph = "") =>
    `<label class="f">${label}</label><input type="number" inputmode="decimal" step="any" data-a="${key}" placeholder="${ph}">`;

  function freqOptions(tipo, which) {
    const f = (L.frequenze[tipo] || {})[which] || [];
    const variable = /pedoni|traffico/.test(tipo);
    return '<option value=""></option>' + f.map((v, i) =>
      `<option value="${i + 1}">${i + 1} · ${esc(v)}${variable ? " (indicativa)" : ""}</option>`).join("");
  }
  function bersHtml(tag) {
    const b = BERS[tag];
    return `<div class="sub">Bersaglio albero</div>
      ${selectField(`${tag}_tA`, "Tipo di bersaglio", "bersaglio")}
      <label class="f">Frequenza / esposizione (classe)</label><select data-a="${tag}_fA" data-freq="${b.la}" data-tipo="${tag}_tA"></select>
      ${tag === "pre" ? numField("pre_m", "Moltiplicatore (n. bersagli)", "1") : ""}
      <div class="sub">Bersaglio rami</div>
      ${selectField(`${tag}_tR`, "Tipo di bersaglio", "bersaglio")}
      <label class="f">Frequenza / esposizione (classe)</label><select data-a="${tag}_fR" data-freq="${b.lr}" data-tipo="${tag}_tR"></select>`;
  }
  function zoneHtml(z) {
    return `<div class="sub">${z.label} <span class="muted">(max ${z.rows.length})</span></div>
      <div class="zone" data-zone="${z.k}"></div>
      <button type="button" class="iconbtn addrow" data-zone="${z.k}">+ difetto</button>`;
  }
  function defectRow(z, d = "", g = "") {
    const dl = "dl_" + z;
    if (!document.getElementById(dl)) {
      const e = document.createElement("datalist");
      e.id = dl;
      e.innerHTML = (L[z] || []).map((v) => `<option value="${esc(v)}">`).join("");
      document.body.appendChild(e);
    }
    const div = document.createElement("div");
    div.className = "defrow";
    div.innerHTML = `<input type="text" list="${dl}" class="dd" placeholder="Scrivi o scegli" value="${esc(d)}">
      <select class="dg">${opts("gravita")}</select><button type="button" class="del2">×</button>`;
    div.querySelector(".dg").value = g;
    div.querySelector(".del2").addEventListener("click", () => div.remove());
    return div;
  }
  function prescRow(i) {
    return `<div class="sub">Prescrizione ${i + 1}</div>` +
      PRESC.map((p) => selectField(`pr${i}_${p.k}`, p.label, p.list)).join("");
  }

  function build(box) {
    box.innerHTML = `
      <details open><summary>Dati generali</summary>${GEN.map((f) => selectField(f.k, f.label, f.list)).join("")}</details>
      <details><summary>Dimensioni</summary><div class="muted" style="margin-top:6px">H e D tr sono quelli della scheda (altezza e diametro a 120 cm).</div>
        ${DIM.map((f) => numField(f.k, f.label)).join("")}</details>
      <details><summary>Diagnosi</summary>${DIAG.map((f) => selectField(f.k, f.label, f.list, f.free)).join("")}</details>
      <details><summary>Difetti e sintomi</summary>${ZONES.map(zoneHtml).join("")}</details>
      <details><summary>Pericolo (P) – pre intervento</summary>${PER.map((p, i) => selectField("P" + i, p, "pericolo")).join("")}</details>
      <details><summary>Bersaglio e impulso – pre intervento</summary>${bersHtml("pre")}
        <div class="muted" style="margin-top:6px">Per pedoni e traffico le fasce esatte dipendono dalle dimensioni: si sceglie la classe, il testo lo calcola la scheda.</div></details>
      <details><summary>Prescrizioni</summary><div id="prescBox"></div><button type="button" class="iconbtn" id="addPresc">+ prescrizione</button></details>
      <label class="postchk"><input type="checkbox" data-a="post_on"> Valutazione post intervento</label>
      <div id="postBox" class="hidden"><details open><summary>Post intervento</summary>
        <div class="muted">Lascia vuote le dimensioni che non cambiano.</div>
        ${POSTDIM.map((f) => numField("pd_" + f.k, f.label, "come pre")).join("")}
        ${bersHtml("post")}
        ${PER.map((p, i) => selectField("PP" + i, p + " – P post", "pericolo_post")).join("")}</details></div>`;
    // la parte post intervento compare solo con la spunta
    const postChk = box.querySelector('[data-a="post_on"]');
    box._updPost = () => box.querySelector("#postBox").classList.toggle("hidden", !postChk.checked);
    postChk.addEventListener("change", box._updPost);
    // frequenze dipendenti dal tipo
    box.querySelectorAll("select[data-freq]").forEach((s) => {
      const tipoSel = box.querySelector(`[data-a="${s.dataset.tipo}"]`);
      const upd = () => { const v = s.value; s.innerHTML = freqOptions(tipoSel.value, s.dataset.freq); s.value = v; };
      tipoSel.addEventListener("change", upd);
      upd();
    });
    box.querySelectorAll(".addrow").forEach((b) => b.addEventListener("click", () => {
      const z = ZONES.find((x) => x.k === b.dataset.zone), cont = box.querySelector(`.zone[data-zone="${z.k}"]`);
      if (cont.children.length >= z.rows.length) return;
      cont.appendChild(defectRow(z.k));
    }));
    let np = 0;
    box.querySelector("#addPresc").addEventListener("click", () => {
      if (np >= 4) return;
      box.querySelector("#prescBox").insertAdjacentHTML("beforeend", prescRow(np++));
    });
    box._resetPresc = () => { np = 0; box.querySelector("#prescBox").innerHTML = ""; };
  }

  function reset(box, remembered = {}) {
    box.querySelectorAll("[data-a]").forEach((e) => {
      if (e.type === "checkbox") e.checked = false; else e.value = "";
    });
    box.querySelectorAll(".zone").forEach((z) => (z.innerHTML = ""));
    box._resetPresc();
    if (remembered.ubicaz) box.querySelector('[data-a="ubicaz"]').value = remembered.ubicaz;
    box._updPost();
    box.querySelectorAll("select[data-freq]").forEach((s) => s.dispatchEvent(new Event("refresh")));
    box.querySelectorAll('select[data-a$="_tA"], select[data-a$="_tR"]').forEach((s) => s.dispatchEvent(new Event("change")));
  }

  // rimette nella scheda i valori salvati (per riaprire un rilievo e modificarlo)
  function fill(box, a) {
    if (!a) return;
    reset(box, {});
    const np = Object.keys(a).reduce((m, k) => { const x = k.match(/^pr(\d+)_/); return x ? Math.max(m, +x[1] + 1) : m; }, 0);
    for (let i = 0; i < np; i++) box.querySelector("#addPresc").click();
    const setv = (e, v) => { if (e.type === "checkbox") e.checked = !!v; else e.value = v ?? ""; };
    // prima i tipi di bersaglio (da cui dipendono le frequenze), poi tutto il resto
    box.querySelectorAll('[data-a$="_tA"], [data-a$="_tR"]').forEach((e) => { setv(e, a[e.dataset.a]); e.dispatchEvent(new Event("change")); });
    box.querySelectorAll("[data-a]").forEach((e) => { if (e.dataset.a in a) { setv(e, a[e.dataset.a]); e.dispatchEvent(new Event("change")); } });
    Object.entries(a.difetti || {}).forEach(([z, rows]) => {
      const cont = box.querySelector(`.zone[data-zone="${z}"]`);
      if (cont) rows.forEach((r) => cont.appendChild(defectRow(z, r.d, r.g)));
    });
  }

  function collect(box) {
    const a = {};
    box.querySelectorAll("[data-a]").forEach((e) => {
      a[e.dataset.a] = e.type === "checkbox" ? e.checked : e.value.trim();
    });
    if (!a.post_on) Object.keys(a).forEach((k) => { if (/^(pd_|post_|PP\d)/.test(k)) delete a[k]; });
    a.difetti = {};
    box.querySelectorAll(".zone").forEach((z) => {
      a.difetti[z.dataset.zone] = Array.from(z.querySelectorAll(".defrow"))
        .map((r) => ({ d: r.querySelector(".dd").value.trim(), g: r.querySelector(".dg").value }))
        .filter((x) => x.d || x.g);
    });
    return a;
  }

  /* -------------------------------------------------- File Excel */
  const colNum = (ref) => ref.match(/^[A-Z]+/)[0].split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const rowNum = (ref) => Number(ref.match(/\d+$/)[0]);

  function getCell(doc, ref) {
    const sd = doc.getElementsByTagNameNS(NS, "sheetData")[0];
    const rn = rowNum(ref);
    let row = null;
    for (const r of sd.getElementsByTagNameNS(NS, "row")) {
      const n = Number(r.getAttribute("r"));
      if (n === rn) { row = r; break; }
      if (n > rn) { row = doc.createElementNS(NS, "row"); row.setAttribute("r", rn); sd.insertBefore(row, r); break; }
    }
    if (!row) { row = doc.createElementNS(NS, "row"); row.setAttribute("r", rn); sd.appendChild(row); }
    const cn = colNum(ref);
    for (const c of row.getElementsByTagNameNS(NS, "c")) {
      const k = colNum(c.getAttribute("r"));
      if (k === cn) return c;
      if (k > cn) { const nc = doc.createElementNS(NS, "c"); nc.setAttribute("r", ref); row.insertBefore(nc, c); return nc; }
    }
    const nc = doc.createElementNS(NS, "c");
    nc.setAttribute("r", ref);
    row.appendChild(nc);
    return nc;
  }
  function setCell(doc, ref, value, kind) {
    const c = getCell(doc, ref);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute("t");
    if (value === null || value === undefined || value === "") return;
    if (kind === "f") {
      const f = doc.createElementNS(NS, "f");
      f.textContent = value;
      c.appendChild(f);
    } else if (kind === "n" || (kind !== "s" && typeof value === "number")) {
      const v = doc.createElementNS(NS, "v");
      v.textContent = String(value);
      c.appendChild(v);
    } else {
      c.setAttribute("t", "inlineStr");
      const is = doc.createElementNS(NS, "is"), t = doc.createElementNS(NS, "t");
      t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
      t.textContent = String(value);
      is.appendChild(t);
      c.appendChild(is);
    }
  }
  const asNum = (s) => {
    if (s === "" || s === null || s === undefined) return "";
    const n = Number(String(s).replace(",", "."));
    return isNaN(n) ? s : n;
  };
  function excelDate(t) {
    const d = new Date(t);
    return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
  }

  async function sheetPath(zip, name) {
    const wb = new DOMParser().parseFromString(await zip.file("xl/workbook.xml").async("string"), "application/xml");
    const rels = new DOMParser().parseFromString(await zip.file("xl/_rels/workbook.xml.rels").async("string"), "application/xml");
    const sh = Array.from(wb.getElementsByTagNameNS(NS, "sheet")).find((s) => s.getAttribute("name") === name);
    if (!sh) return null;
    const rid = sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const rel = Array.from(rels.getElementsByTagName("Relationship")).find((r) => r.getAttribute("Id") === rid);
    let t = rel.getAttribute("Target");
    return t.startsWith("/") ? t.slice(1) : "xl/" + t;
  }

  async function checkTemplate(blob) {
    const zip = await JSZip.loadAsync(blob);
    if (!(await sheetPath(zip, "ORD")) || !(await sheetPath(zip, "A")))
      throw new Error("Il file non contiene i fogli ORD e A della scheda ARETÉ");
    if (Object.keys(zip.files).some((n) => /vbaProject/i.test(n)))
      throw new Error("Usa la versione senza macro (Scheda_ARETE_ORD_telefono.xlsx)");
    return true;
  }

  // r = rilievo dell'app, a = dati ARETÉ raccolti
  async function makeXlsx(templateBlob, r) {
    const a = r.arete;
    const zip = await JSZip.loadAsync(templateBlob);
    const path = await sheetPath(zip, "ORD");
    const doc = new DOMParser().parseFromString(await zip.file(path).async("string"), "application/xml");
    const S = (ref, v, k) => setCell(doc, ref, v, k);

    S("H6", r.dataValutazione ? excelDate(r.dataValutazione + "T12:00:00") : excelDate(r.createdAt), "n");
    if (r.operatore) S("D6", r.operatore, "s");   // rilevatore = chi compila (altrimenti resta l'utente registrato del modello)
    S("J6", r.specie, "s");
    S("U6", asNum(r.ngs || r.nprog || r.ncart));
    S("C9", asNum(r.altezza));
    S("E9", asNum(r.diametro120));
    GEN.forEach((f) => S(f.cell, a[f.k], "s"));
    DIM.forEach((f) => S(f.cell, asNum(a[f.k])));
    DIAG.forEach((f) => S(f.cell, a[f.k], "s"));
    ZONES.forEach((z) => {
      const list = (a.difetti && a.difetti[z.k]) || [];
      z.rows.forEach((row, i) => {
        S(z.d + row, list[i] ? list[i].d : "", "s");
        S(z.g + row, list[i] ? list[i].g : "", "s");
      });
    });
    [26, 27, 28, 29].forEach((row, i) => S("E" + row, asNum(a["P" + i])));
    const bers = (tag) => {
      const b = BERS[tag];
      S(b.tA, a[`${tag}_tA`], "s");
      S(b.fA, a[`${tag}_fA`] ? `INDEX(${b.fArng},${Number(a[`${tag}_fA`])})` : "", "f");
      S(b.tR, a[`${tag}_tR`], "s");
      S(b.fR, a[`${tag}_fR`] ? `INDEX(${b.fRrng},${Number(a[`${tag}_fR`])})` : "", "f");
    };
    bers("pre");
    S("H32", asNum(a.pre_m) || 1);
    [42, 43, 44, 45].forEach((row, i) => PRESC.forEach((p) => S(p.col + row, a[`pr${i}_${p.k}`] || "", "s")));
    if (a.post_on) {
      POSTDIM.forEach((f) => { if (a["pd_" + f.k] !== "") S(f.cell, asNum(a["pd_" + f.k])); });
      bers("post");
      [54, 55, 56, 57].forEach((row, i) => S("E" + row, asNum(a["PP" + i])));
    } else {
      ["C51", "B52", "N51", "M52", "E54", "E55", "E56", "E57"].forEach((c) => S(c, ""));
    }
    const ip = typeof intTxt === "function" ? intTxt(r) : "";
    const note = [r.problematiche && "Note di campo: " + r.problematiche, ip && "Interventi proposti: " + ip,
      r.interventi && "Note sugli interventi: " + r.interventi].filter(Boolean).join(" · ");
    S("C59", note, "s");

    zip.file(path, new XMLSerializer().serializeToString(doc));
    // ricalcolo completo all'apertura in Excel
    let wbx = await zip.file("xl/workbook.xml").async("string");
    if (/<calcPr[^>]*fullCalcOnLoad/.test(wbx)) { /* già presente */ }
    else if (/<calcPr/.test(wbx)) wbx = wbx.replace(/<calcPr/, '<calcPr fullCalcOnLoad="1"');
    else wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
    // il foglio originale contiene un nome con una vecchia macro Excel 4.0 (xlm="1"): lo tolgo,
    // altrimenti Excel avvisa «file senza macro ma con contenuto da macro»
    wbx = wbx.replace(/<definedName\b[^>]*\bxlm="1"[^>]*\/>/g, "").replace(/<definedName\b[^>]*\bxlm="1"[^>]*>[\s\S]*?<\/definedName>/g, "")
      .replace(/<definedNames>\s*<\/definedNames>/g, "");
    zip.file("xl/workbook.xml", wbx);
    if (zip.file("xl/calcChain.xml")) zip.remove("xl/calcChain.xml");
    return zip.generateAsync({ type: "blob", compression: "DEFLATE",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }


  /* ---------------------------------------- Foglio dati ARETÉ: Sezione | Voce | Valore (solo voci compilate) */
  // file semplice da leggere anche dal telefono; i dati servono per compilare la scheda ufficiale
  function dataRows(r) {
    const a = r.arete || {}, rows = [], add = (sez, voce, val) => { if (val !== "" && val !== null && val !== undefined && val !== false) rows.push([sez, voce, val]); };
    const data = r.dataValutazione ? r.dataValutazione.split("-").reverse().join("/") : "";
    const num = (v) => { const n = Number(String(v).replace(",", ".")); return v === "" || isNaN(n) ? v : n; };
    const freq = (tipo, which, v) => {
      if (!v) return "";
      const t = ((L && L.frequenze[tipo]) || {})[which] || [];
      return `classe ${v}` + (t[v - 1] ? " · " + t[v - 1] : "");
    };
    add("Rilievo", "Nome", r.cartella); add("Rilievo", "Progetto", r.lavoro); add("Rilievo", "Data", data);
    add("Rilievo", "Rilevatore", r.operatore);
    add("Rilievo", "Numeri", [r.ngs && "gs. " + r.ngs, r.ncart && "cart. " + r.ncart, r.nprog && "rel. " + r.nprog].filter(Boolean).join("  "));
    const p = r.posizione || {};
    if (p.lat) add("Rilievo", "Posizione", `${p.lat.toFixed(6)}, ${p.lon.toFixed(6)}`);
    add("Rilievo", "Specie", r.specie);
    GEN.forEach((f) => add("Dati generali", f.label, a[f.k]));
    add("Dimensioni", "H – altezza (m)", num(r.altezza || "")); add("Dimensioni", "D tr – diametro a 120 cm (cm)", num(r.diametro120 || ""));
    DIM.forEach((f) => add("Dimensioni", f.label, num(a[f.k] || "")));
    DIAG.forEach((f) => add("Diagnosi", f.label, a[f.k]));
    ZONES.forEach((z) => ((a.difetti && a.difetti[z.k]) || []).forEach((d) => add("Difetti – " + z.label, d.d || "(difetto)", d.g ? "gravità: " + d.g : "")));
    const per = (key, sez) => PER.forEach((nm, i) => add(sez, nm, num(a[key + i] || "")));
    per("P", "Pericolo – pre intervento");
    const bers = (tag, sez) => {
      const b = BERS[tag];
      add(sez, "Bersaglio albero", a[`${tag}_tA`]); add(sez, "Frequenza albero", freq(a[`${tag}_tA`], b.la, Number(a[`${tag}_fA`]) || 0));
      if (tag === "pre") add(sez, "Moltiplicatore", num(a.pre_m || ""));
      add(sez, "Bersaglio rami", a[`${tag}_tR`]); add(sez, "Frequenza rami", freq(a[`${tag}_tR`], b.lr, Number(a[`${tag}_fR`]) || 0));
    };
    bers("pre", "Bersaglio – pre intervento");
    for (let i = 0; i < 4; i++) {
      const t = PRESC.map((q) => a[`pr${i}_${q.k}`] ? `${q.label}: ${a[`pr${i}_${q.k}`]}` : "").filter(Boolean).join(" · ");
      add("Prescrizioni", `Prescrizione ${i + 1}`, t);
    }
    if (a.post_on) {
      POSTDIM.forEach((f) => add("Post intervento", f.label, num(a["pd_" + f.k] || "")));
      bers("post", "Post intervento");
      per("PP", "Post intervento – pericolo");
    }
    add("Note", "Note di campo", r.problematiche);
    add("Note", "Interventi proposti", typeof intTxt === "function" ? intTxt(r) : "");
    add("Note", "Note sugli interventi", r.interventi);
    return rows;
  }
  async function dataXlsx(r) {
    const rows = [["Sezione", "Voce", "Valore"], ...dataRows(r)];
    const x = (v) => String(v).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
    let prev = "";
    const body = rows.map((row, k) => {
      const head = k === 0, nuova = !head && row[0] !== prev; if (!head) prev = row[0];
      return `<row r="${k + 1}">` + row.map((v, i) => {
        const ref = "ABC"[i] + (k + 1), st = head ? 1 : i === 0 ? (nuova ? 2 : 3) : 4;
        if (typeof v === "number") return `<c r="${ref}" s="${st}"><v>${v}</v></c>`;
        return `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${x(i === 0 && !head && !nuova ? "" : v)}</t></is></c>`;
      }).join("") + "</row>";
    }).join("");
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/><col min="3" max="3" width="44" customWidth="1"/></cols><sheetData>${body}</sheetData></worksheet>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FF2E5E3E"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2E5E3E"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top style="thin"><color rgb="FFB7C9BC"/></top><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const z = new JSZip(), R = "http://schemas.openxmlformats.org";
    z.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${R}/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
    z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${R}/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    z.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS}" xmlns:r="${R}/officeDocument/2006/relationships"><sheets><sheet name="Dati ARETE" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    z.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${R}/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    z.file("xl/worksheets/sheet1.xml", sheet);
    z.file("xl/styles.xml", styles);
    return z.generateAsync({ type: "blob", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  async function init() {
    L = await (await fetch("arete_lists.json")).json();
  }

  return { init, build, reset, fill, collect, makeXlsx, dataXlsx, dataRows, checkTemplate, specie: () => (L ? L.specie : []) };
})();
