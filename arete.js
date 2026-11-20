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
      <details><summary>Post intervento (facoltativo)</summary>
        <label class="f"><input type="checkbox" data-a="post_on"> Compila la valutazione post intervento</label>
        <div class="muted">Lascia vuote le dimensioni che non cambiano.</div>
        ${POSTDIM.map((f) => numField("pd_" + f.k, f.label, "come pre")).join("")}
        ${bersHtml("post")}
        ${PER.map((p, i) => selectField("PP" + i, p + " – P post", "pericolo_post")).join("")}</details>`;
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
    box.querySelectorAll("select[data-freq]").forEach((s) => s.dispatchEvent(new Event("refresh")));
    box.querySelectorAll('select[data-a$="_tA"], select[data-a$="_tR"]').forEach((s) => s.dispatchEvent(new Event("change")));
  }

  function collect(box) {
    const a = {};
    box.querySelectorAll("[data-a]").forEach((e) => {
      a[e.dataset.a] = e.type === "checkbox" ? e.checked : e.value.trim();
    });
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

    S("H6", excelDate(r.createdAt), "n");
    if (r.operatore) S("D6", r.operatore, "s");   // rilevatore = chi compila (altrimenti resta l'utente registrato del modello)
    S("J6", r.specie, "s");
    S("U6", asNum(r.tipoCantiere === "Comune" ? r.ngs || r.nprog : r.nprog || r.ncart));
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
    const note = [r.problematiche && "Note di campo: " + r.problematiche,
      r.interventi && "Interventi e nuova CPC: " + r.interventi].filter(Boolean).join(" · ");
    S("C59", note, "s");

    zip.file(path, new XMLSerializer().serializeToString(doc));
    // ricalcolo completo all'apertura in Excel
    let wbx = await zip.file("xl/workbook.xml").async("string");
    if (/<calcPr[^>]*fullCalcOnLoad/.test(wbx)) { /* già presente */ }
    else if (/<calcPr/.test(wbx)) wbx = wbx.replace(/<calcPr/, '<calcPr fullCalcOnLoad="1"');
    else wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
    zip.file("xl/workbook.xml", wbx);
    if (zip.file("xl/calcChain.xml")) zip.remove("xl/calcChain.xml");
    return zip.generateAsync({ type: "blob", compression: "DEFLATE",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  async function init() {
    L = await (await fetch("arete_lists.json")).json();
  }

  return { init, build, reset, collect, makeXlsx, checkTemplate, specie: () => (L ? L.specie : []) };
})();
