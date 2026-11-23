/* File in uscita per ogni progetto:
 *  Schede/   scheda di valutazione Word di ogni pianta (+ scheda ARETÉ se compilata)
 *  Foto/     foto di tutte le piante, nominate e in ordine
 *  Riepilogo_<progetto>.xlsx / .csv   tabella con tutte le piante (una riga per pianta)
 *  Mappa/    <progetto>.geojson + .qml (QGIS) e <progetto>.kml (Google Earth), pallini verdi con gli attributi
 */
"use strict";

const OUT = (() => {
  const VERDE = "2E5E3E", VERDE_CHIARO = "F1F6F2", GRIGIO = "6B6B6B", LINEA = "D0D5D1", ETICH = "F3F4F2", FONT = "Calibri", W = 9638;
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" }[c]));
  const dash = (v) => (v === undefined || v === null || String(v).trim() === "" ? "—" : String(v));
  const itDate = (iso) => { if (!iso) return "—"; const [y, m, d] = iso.split("-"); return `${d}/${m}/${y}`; };
  const mapsUrl = (p) => `https://www.google.com/maps/search/?api=1&query=${p.lat.toFixed(6)},${p.lon.toFixed(6)}`;
  const tipoTitolo = (r) => (r.tipoValutazione === "Sorveglianza scavo" ? "SORVEGLIANZA SCAVO" : "VALUTAZIONE FITOSTATICA");

  /* ---------------------------------------------------------- immagini */
  async function dims(blob) { const b = await createImageBitmap(blob); const d = { w: b.width, h: b.height }; b.close && b.close(); return d; }
  const buf = (blob) => blob.arrayBuffer();
  function canvasBlob(c, type = "image/png") { return new Promise((ok) => c.toBlob(ok, type, 0.9)); }

  async function logo() {
    if (!logo.b) logo.b = await (await fetch("brand/logo_orizzontale.png")).blob();
    return logo.b;
  }

  // estratto di mappa con il cerchio della pianta: satellite Esri (se c'è la chiave), altrimenti OpenStreetMap. Serve la rete.
  const satKey = () => (window.APP_CONFIG && APP_CONFIG.esriKey) || "";
  const SRC = {
    sat: { url: (z, x, y) => `https://ibasemaps-api.arcgis.com/arcgis/rest/services/World_Imagery/MapServer/tile/${z}/${y}/${x}?token=${encodeURIComponent(satKey())}`, attr: "Esri, Maxar, Earthstar Geographics" },
    osm: { url: (z, x, y) => `https://tile.openstreetmap.org/${z}/${x}/${y}.png`, attr: "© OpenStreetMap" },
  };
  // dove non ha immagini dettagliate Esri restituisce una tessera grigia uniforme: la riconosco dalla poca varietà di colori
  function isPlaceholder(bmp) {
    const c = document.createElement("canvas"); c.width = c.height = 16;
    const g = c.getContext("2d"); g.drawImage(bmp, 0, 0, 16, 16);
    const d = g.getImageData(0, 0, 16, 16).data; let m = 0, m2 = 0;
    for (let i = 0; i < d.length; i += 4) { const v = (d[i] + d[i + 1] + d[i + 2]) / 3; m += v; m2 += v * v; }
    const n = d.length / 4;
    return Math.sqrt(Math.max(0, m2 / n - (m / n) ** 2)) < 4;
  }
  async function drawTiles(src, lat, lon, z, W0, H0) {
    const n = 2 ** z, x = ((lon + 180) / 360) * n;
    const y = ((1 - Math.log(Math.tan((lat * Math.PI) / 180) + 1 / Math.cos((lat * Math.PI) / 180)) / Math.PI) / 2) * n;
    const c = document.createElement("canvas"); c.width = W0; c.height = H0;
    const g = c.getContext("2d");
    g.fillStyle = "#e9eee8"; g.fillRect(0, 0, W0, H0);
    const x0 = x * 256 - W0 / 2, y0 = y * 256 - H0 / 2, cx = Math.floor(x), cy = Math.floor(y);
    let ok = 0, centerBad = false;
    const jobs = [];
    for (let tx = Math.floor(x0 / 256); tx <= Math.floor((x0 + W0) / 256); tx++)
      for (let ty = Math.floor(y0 / 256); ty <= Math.floor((y0 + H0) / 256); ty++)
        jobs.push(fetch(src.url(z, tx, ty), { mode: "cors" })
          .then((r) => (r.ok ? r.blob() : null)).then((b) => (b ? createImageBitmap(b) : null))
          .then((bmp) => {
            if (!bmp) return;
            if (src === SRC.sat && tx === cx && ty === cy && isPlaceholder(bmp)) centerBad = true;
            g.drawImage(bmp, tx * 256 - x0, ty * 256 - y0); ok++;
          }).catch(() => {}));
    await Promise.all(jobs);
    return { c, g, ok, centerBad };
  }
  async function staticMap(lat, lon, z = 19, W0 = 840, H0 = 600) {
    let src = satKey() ? SRC.sat : SRC.osm, res = null;
    if (src === SRC.sat) {
      for (let zz = z; zz >= 16; zz--) { res = await drawTiles(src, lat, lon, zz, W0, H0); z = zz; if (!res.ok || !res.centerBad) break; }
      if (!res || !res.ok) src = SRC.osm;
    }
    if (src === SRC.osm) { z = Math.min(z, 18); res = await drawTiles(src, lat, lon, z, W0, H0); }
    const { c, g } = res;
    // cerchio giallo bordato di scuro: si vede sia sul satellite sia sulla mappa
    g.beginPath(); g.arc(W0 / 2, H0 / 2, 26, 0, Math.PI * 2); g.lineWidth = 9; g.strokeStyle = "rgba(0,0,0,.55)"; g.stroke();
    g.lineWidth = 5; g.strokeStyle = "#ffd60a"; g.stroke();
    g.beginPath(); g.arc(W0 / 2, H0 / 2, 4, 0, Math.PI * 2); g.fillStyle = "#ffd60a"; g.fill();
    // barra di scala
    const mpp = (156543.03392 * Math.cos((lat * Math.PI) / 180)) / 2 ** z;
    const len = [5, 10, 20, 50, 100, 200].find((m) => m / mpp >= 90) || 200, px = len / mpp;
    g.fillStyle = "rgba(255,255,255,.88)"; g.fillRect(12, H0 - 44, px + 78, 32);
    g.fillStyle = "#222"; g.fillRect(22, H0 - 24, px, 5); g.fillRect(22, H0 - 32, 3, 13); g.fillRect(22 + px - 3, H0 - 32, 3, 13);
    g.font = "bold 18px Arial"; g.fillText(`${len} m`, 32 + px, H0 - 19);
    // nord in alto
    g.fillStyle = "rgba(255,255,255,.88)"; g.beginPath(); g.arc(W0 - 34, 36, 24, 0, Math.PI * 2); g.fill();
    g.fillStyle = "#222"; g.beginPath(); g.moveTo(W0 - 34, 16); g.lineTo(W0 - 42, 38); g.lineTo(W0 - 26, 38); g.closePath(); g.fill();
    g.font = "bold 14px Arial"; g.fillText("N", W0 - 39, 55);
    const at = src.attr; g.font = "15px Arial"; const tw = g.measureText(at).width + 16;
    g.fillStyle = "rgba(255,255,255,.85)"; g.fillRect(W0 - tw, H0 - 24, tw, 24);
    g.fillStyle = "#333"; g.fillText(at, W0 - tw + 8, H0 - 7);
    return { blob: await canvasBlob(c, "image/jpeg"), attr: src === SRC.sat ? "immagine satellitare Esri" : "© OpenStreetMap" };
  }

  /* ---------------------------------------------- scheda di valutazione Word */
  async function schedaDocx(r, opts = {}) {
    const D = window.docx;
    const { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
      AlignmentType, VerticalAlign, Header, Footer, PageNumber, TabStopType, ExternalHyperlink, TableLayoutType } = D;
    const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const noB = { top: none, bottom: none, left: none, right: none };
    const thin = (c = LINEA) => ({ style: BorderStyle.SINGLE, size: 4, color: c });
    const tblNoB = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };
    const run = (t, o = {}) => new TextRun({ text: String(t), font: FONT, size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "1D1D1B" });
    const P = (t, o = {}) => new Paragraph({ children: Array.isArray(t) ? t : [run(t, o)], alignment: o.align,
      spacing: { before: o.before || 0, after: o.after ?? 60, line: o.line || 264 }, keepNext: o.keepNext, keepLines: true });
    const cell = (ch, w, o = {}) => new TableCell({ children: Array.isArray(ch) ? ch : [ch], width: { size: w, type: WidthType.DXA },
      shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
      borders: o.borders || { top: none, left: none, right: none, bottom: thin() },
      margins: { top: o.mt ?? 70, bottom: o.mb ?? 70, left: o.ml ?? 110, right: o.mr ?? 110 }, verticalAlign: o.v || VerticalAlign.CENTER });
    const heading = (n, t) => new Paragraph({ children: [new TextRun({ text: `${n}. ${t.toUpperCase()}`, font: FONT, size: 22, bold: true, color: VERDE, characterSpacing: 20 })],
      spacing: { before: 300, after: 120 }, keepNext: true, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: VERDE, space: 3 } } });
    const kv = (rows) => {
      const c1 = 3300, c2 = W - c1;
      return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [c1, c2], layout: TableLayoutType.FIXED,
        borders: { top: thin(), bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
        rows: rows.map(([k, v], i) => new TableRow({ cantSplit: true, children: [
          cell(P(k, { bold: true, after: 0, keepNext: i < rows.length - 1 }), c1, { fill: ETICH }),
          cell(Array.isArray(v) ? v : P(dash(v), { after: 0, keepNext: i < rows.length - 1 }), c2)] })) });
    };
    const img = async (blob, w, type) => { const d = await dims(blob); return new ImageRun({ data: await buf(blob), type, transformation: { width: w, height: Math.round((w * d.h) / d.w) } }); };
    const caption = (t) => P(t, { size: 18, color: GRIGIO, after: 0, before: 60 });
    const pos = r.posizione && r.posizione.lat ? r.posizione : null;
    const data = itDate(r.dataValutazione);
    const fotoBlob = (f) => f.arrow || f.clean;
    const dirTxt = (f) => (f.heading !== null && f.heading !== undefined ? `direzione dello scatto ${Math.round(f.heading)}° ${cardinal(f.heading)}` : "");
    // didascalia: «Figura N», con la direzione dello scatto solo se attiva in Impostazioni → Bussola
    // etichetta solo se il rilievo è stato fatto con le etichette attive (così ogni progetto resta coerente)
    const capTxt = (i, f) => [`Figura ${i}` + (f.etichetta && r.etichetteOn !== false ? " – " + f.etichetta : ""), typeof S !== "undefined" && S.gradiFoto ? dirTxt(f) : ""].filter(Boolean).join(" · ");
    const nums = numeri(r).join(" · ") || "—";
    const isArete = r.protocollo === "ARETE";
    const protTxt = isArete ? "ARETÉ" : r.protocollo === "SIA" ? "SIA" : "";
    const cpcTxt = isArete ? "vedi scheda ARETÉ" : r.cpc || "";
    const para = (t) => String(t || "").split(/\r?\n/).filter((x) => x.trim()).map((x, i, a) => P(x, { after: i < a.length - 1 ? 40 : 0 }));
    const noteCell = (t) => { const p = para(t); return p.length ? p : [P("—", { after: 0 })]; };

    // intestazione e piè di pagina
    const lg = await logo();
    const header = new Header({ children: [
      new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [4300, W - 4300], layout: TableLayoutType.FIXED, borders: tblNoB,
        rows: [new TableRow({ children: [
          cell(new Paragraph({ children: [await img(lg, 190, "png")] }), 4300, { borders: noB, ml: 0, mt: 0, mb: 0 }),
          cell([P([run(tipoTitolo(r), { size: 30, bold: true, color: VERDE })], { align: AlignmentType.RIGHT, after: 20 }),
            P("Scheda di rilievo arboreo", { size: 19, color: GRIGIO, align: AlignmentType.RIGHT, after: 20 }),
            P("Rif. " + r.cartella, { size: 15, color: GRIGIO, align: AlignmentType.RIGHT, after: 0 })], W - 4300, { borders: noB, mr: 0, mt: 0, mb: 0 })] })] }),
      new Paragraph({ children: [], spacing: { before: 60, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: VERDE, space: 1 } } }),
    ] });
    const studio = [S.studio, S.qualifica].filter(Boolean).join(" · ") || r.operatore || "";
    const footer = new Footer({ children: [
      new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINEA, space: 4 } }, tabStops: [{ type: TabStopType.RIGHT, position: W }], spacing: { after: 20 },
        children: [run(studio, { size: 16, color: GRIGIO }), new TextRun({ text: "\tPagina ", font: FONT, size: 16, color: GRIGIO }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GRIGIO }), new TextRun({ text: " di ", font: FONT, size: 16, color: GRIGIO }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: GRIGIO })] }),
      P(`Documento emesso il ${new Date().toLocaleDateString("it-IT")} · Rilievo eseguito con GabTree Rilievi di Campo`, { size: 15, color: GRIGIO, after: 0 }),
    ] });

    // riquadro dell'esito
    const q = W / 4;
    const bb = (p) => ({ top: p.t ? thin(VERDE) : none, bottom: p.b ? thin(VERDE) : none, left: p.l ? thin(VERDE) : none, right: p.r ? thin(VERDE) : none });
    const sumRow = (items, t, b) => new TableRow({ cantSplit: true, children: items.map(([l, v], i) => cell([
      P(l.toUpperCase(), { size: 15, bold: true, color: VERDE, after: 20 }), P(dash(v), { bold: true, after: 0 })], q,
      { fill: VERDE_CHIARO, borders: bb({ t, b, l: i === 0, r: i === items.length - 1 }), mt: t ? 120 : 60, mb: b ? 120 : 60, v: VerticalAlign.TOP })) });
    const esito = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [q, q, q, q], layout: TableLayoutType.FIXED, rows: [
      sumRow([["Data della valutazione", data], ["Pianta", nums], ["Specie", r.specie], ["Diametro / altezza", [r.diametro120 && r.diametro120 + " cm", r.altezza && r.altezza + " m"].filter(Boolean).join(" · ")]], true, false),
      sumRow([["Protocollo", protTxt], ["Nuova CPC", cpcTxt], ["Progetto", r.lavoro], ["Rilevatore", r.operatore]], false, true),
    ] });

    // foto generale + mappa + QR
    const children = [esito, P("", { after: 120 })];
    // foto generale a sinistra, estratto satellitare (un po' più grande) a destra
    const f1 = r.foto[0], CL = 4150, CR = W - CL;
    const left = f1 ? [new Paragraph({ children: [await img(fotoBlob(f1), 245, "jpg")] }), caption(capTxt(1, f1, "Vista generale"))] : [P("Nessuna foto", { color: GRIGIO })];
    let right = [P("Posizione non disponibile", { color: GRIGIO })];
    if (pos) {
      right = [P("Mappa non disponibile (documento creato senza rete)", { color: GRIGIO })];
      try {
        if (navigator.onLine) { const m = await staticMap(pos.lat, pos.lon); right = [new Paragraph({ children: [await img(m.blob, 330, "jpg")], alignment: AlignmentType.RIGHT }), caption(`Posizione della pianta (cerchio giallo) – ${m.attr}`)]; }
      } catch (e) { /* offline */ }
    }
    children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [CL, CR], layout: TableLayoutType.FIXED, borders: tblNoB,
      rows: [new TableRow({ cantSplit: true, children: [cell(left, CL, { borders: noB, ml: 0, v: VerticalAlign.TOP }), cell(right, CR, { borders: noB, mr: 0, v: VerticalAlign.TOP })] })] }));

    const coord = pos ? [new Paragraph({ spacing: { after: 0 }, children: [
      run(`${pos.lat.toFixed(6)}, ${pos.lon.toFixed(6)} (WGS84${pos.manuale ? ", posizione corretta a mano" : pos.acc ? ", ± " + Math.round(pos.acc) + " m" : ""}) · `),
      new ExternalHyperlink({ link: mapsUrl(pos), children: [new TextRun({ text: "apri in Google Maps", font: FONT, size: 21, color: "1A5FB4", underline: {} })] })] })] : "—";
    children.push(heading(1, "Dati generali"), kv([["Data della valutazione", data], ["Rilevatore", r.operatore], ["Progetto / committente", r.lavoro], ["Coordinate", coord]]));
    children.push(heading(2, "Esemplare"), kv([["N. Greenspaces (gs.)", r.ngs], ["N. cartellino (cart.)", r.ncart], ["N. progressivo (rel.)", r.nprog],
      ["Specie", r.specie], ["Diametro a 120 cm", r.diametro120 ? r.diametro120 + " cm" : ""], ["Altezza", r.altezza ? r.altezza + " m" : ""]]));
    // 3. Valutazione: protocollo e CPC, note di campo, interventi, poi le voci spuntate con le loro note
    const val = [["Protocollo di valutazione", protTxt || "—"]];
    if (isArete) val.push(["Classe di rischio", "vedi scheda ARETÉ allegata"]);
    else val.push(["Nuova CPC", r.cpc || "—"]);
    if (r.tipoValutazione === "Sorveglianza scavo") val.unshift(["Tipo di rilievo", "Sorveglianza scavo"]);
    const ip = r.interventiProposti || [];
    val.push(["Note di campo", noteCell(r.problematiche)],
      ["Interventi proposti", ip.length
        ? ip.map((x, i) => P([run(x.nome, { bold: true }), run(x.tempo ? " – " + x.tempo : "")], { after: i < ip.length - 1 ? 40 : 0 }))
        : [P("Nessun intervento proposto", { after: 0 })]]);
    if ((r.interventi || "").trim()) val.push(["Note sugli interventi", noteCell(r.interventi)]);
    if (r.strumOn) val.push(["Indagine strumentale", [
      P([run("Strumento: ", { bold: true }), run(dash(r.strumento))], { after: 20 }),
      P([run("Esito della prova: ", { bold: true }), run(r.esitoProva ? "rischio " + r.esitoProva.toLowerCase() : "—")], { after: r.strumNote ? 20 : 0 }),
      ...para(r.strumNote), ...(r.allegati && r.allegati.length ? [P(`Allegati: ${r.allegati.map((a) => a.name).join(", ")}`, { size: 18, color: GRIGIO, after: 0 })] : [])]]);
    if (r.potOn) val.push(["Potatura", noteCell(r.potNote)]);
    if (r.scavoOn) val.push(["Scavo", noteCell(r.scavoNote)]);
    if (r.altroOn) val.push(["Altro", noteCell(r.altroNote)]);
    children.push(heading(3, "Valutazione"), kv(val));

    // documentazione fotografica: 2 per riga
    if (r.foto.length > 1) {
      children.push(heading(4, "Documentazione fotografica"));
      const rest = r.foto.slice(1), rows = [];
      for (let i = 0; i < rest.length; i += 2) {
        const pair = rest.slice(i, i + 2);
        const cells = [];
        for (let k = 0; k < 2; k++) {
          const f = pair[k];
          cells.push(cell(f ? [new Paragraph({ children: [await img(fotoBlob(f), 290, "jpg")], alignment: AlignmentType.CENTER }), caption(capTxt(i + k + 2, f, ""))] : [P("")],
            4819, { borders: noB, v: VerticalAlign.TOP, mb: 200 }));
        }
        rows.push(new TableRow({ cantSplit: true, children: cells }));
      }
      children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [4819, 4819], layout: TableLayoutType.FIXED, borders: tblNoB, rows }));
    }
    children.push(heading(r.foto.length > 1 ? 5 : 4, "Note metodologiche"),
      P((r.tipoValutazione === "Sorveglianza scavo"
        ? "Sorveglianza eseguita durante i lavori di scavo in prossimità dell'esemplare. "
        : "La valutazione è stata eseguita con metodo visivo (VTA – Visual Tree Assessment)" + (r.strumOn ? " e con indagine strumentale" : "") +
          (isArete ? ", secondo il protocollo ARETÉ (scheda allegata)" : r.protocollo === "SIA" ? ", secondo il protocollo SIA con attribuzione della classe di propensione al cedimento (CPC)" : "") + ". ") +
        "I risultati si riferiscono alle condizioni dell'esemplare rilevate alla data del sopralluogo e possono variare per eventi meteorici, interventi o evoluzione di patologie.",
        { line: 276, after: 360 }));
    children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [4819, 4819], layout: TableLayoutType.FIXED, borders: tblNoB,
      rows: [new TableRow({ cantSplit: true, children: [
        cell([P("Luogo e data", { bold: true }), P(data)], 4819, { borders: noB, ml: 0, v: VerticalAlign.TOP }),
        cell([P("Il tecnico", { bold: true, align: AlignmentType.RIGHT }), P(S.qualifica || r.operatore || "", { align: AlignmentType.RIGHT }),
          new Paragraph({ children: [], spacing: { before: 700, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "8A8A8A", space: 1 } } }),
          P("(firma e timbro)", { size: 17, color: GRIGIO, align: AlignmentType.RIGHT })], 4819, { borders: noB, mr: 0, v: VerticalAlign.TOP })] })] }));

    const doc = new Document({ creator: "GabTree Rilievi di Campo", title: r.cartella, styles: { default: { document: { run: { font: FONT, size: 21 } } } },
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1900, right: 1134, bottom: 1300, left: 1134, header: 600, footer: 560 } } },
        headers: { default: header }, footers: { default: footer }, children }] });
    return Packer.toBlob(doc);
  }

  /* ------------------------------------------------------ riepilogo */
  // Lat e Lon: indici LAT/LON (formato numerico nel riepilogo, tolte dagli attributi del vettoriale)
  const COLS = [
    ["N.", 5], ["Nome rilievo", 44], ["Data valutazione", 12], ["N. gs", 9], ["N. cart", 9], ["N. progr.", 9],
    ["Specie", 28], ["Diametro 120 cm (cm)", 11], ["Altezza (m)", 9], ["Protocollo", 10], ["Nuova CPC", 9], ["Note di campo", 40],
    ["Interventi proposti", 44], ["Note sugli interventi", 34], ["Lat", 11], ["Lon", 11],
    ["Posizione corretta a mano", 10], ["N. foto", 7], ["Rilevatore", 18], ["Maps", 40],
  ];
  const LAT = 14, LON = 15;
  const intTxt = (r) => (r.interventiProposti || []).map((x) => x.nome + (x.tempo ? ` (${x.tempo})` : "")).join("; ");
  function rowOf(r, i) {
    const p = r.posizione || {};
    return [i + 1, r.cartella, itDate(r.dataValutazione),
      r.ngs || "", r.ncart || "", r.nprog ? Number(r.nprog) : "", r.specie || "", num(r.diametro120) ?? "", num(r.altezza) ?? "",
      r.protocollo === "ARETE" ? "ARETÉ" : r.protocollo || (r.tipoValutazione === "Sorveglianza scavo" ? "Sorveglianza scavo" : ""), r.cpc || "",
      r.problematiche || "", intTxt(r) || "Nessun intervento proposto", r.interventi || "",
      p.lat ? Number(p.lat.toFixed(6)) : "", p.lon ? Number(p.lon.toFixed(6)) : "",
      p.manuale ? "sì" : "", r.foto.length, r.operatore || "", p.lat ? mapsUrl(p) : ""];
  }
  function csv(list) {
    const lines = [COLS.map((c) => c[0]), ...list.map(rowOf)].map((row) =>
      row.map((x) => { const s = typeof x === "number" ? String(x).replace(".", ",") : String(x ?? ""); return `"${s.replace(/"/g, '""').replace(/\r?\n/g, " ")}"`; }).join(";"));
    return new Blob(["﻿" + lines.join("\r\n") + "\r\n"], { type: "text/csv;charset=utf-8" });
  }
  // file Excel minimo, scritto a mano: intestazione verde, filtri, riga bloccata, stampa su una pagina in orizzontale
  async function xlsx(list, title) {
    const colL = (i) => { let s = ""; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
    const cellX = (ref, v, st) => {
      if (v === "" || v === null || v === undefined) return st ? `<c r="${ref}" s="${st}"/>` : "";
      if (typeof v === "number") return `<c r="${ref}" s="${st || 0}"><v>${v}</v></c>`;
      return `<c r="${ref}" s="${st || 2}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
    };
    const rows = [`<row r="1">${COLS.map((c, i) => cellX(colL(i) + 1, c[0], 1)).join("")}</row>`];
    list.forEach((r, k) => {
      const vals = rowOf(r, k), rn = k + 2;
      rows.push(`<row r="${rn}">${vals.map((v, i) => cellX(colL(i) + rn, v, typeof v === "number" ? (i === LAT || i === LON ? 3 : 4) : 2)).join("")}</row>`);
    });
    const last = colL(COLS.length - 1) + (list.length + 1);
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${last}"/>
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>${COLS.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c[1]}" customWidth="1"/>`).join("")}</cols>
<sheetData>${rows.join("")}</sheetData>
<autoFilter ref="A1:${last}"/>
<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>
<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>
<headerFooter><oddHeader>&amp;L&amp;B${esc(title)}&amp;RGabTree Rilievi di Campo</oddHeader><oddFooter>&amp;RPagina &amp;P di &amp;N</oddFooter></headerFooter>
</worksheet>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000000"/></numFmts>
<fonts count="2"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2E5E3E"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FFD0D5D1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const z = new JSZip();
    z.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
    z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    z.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Piante" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Piante!$A$1:$${last.replace(/\d+$/, "")}$${list.length + 1}</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Piante!$1:$1</definedName></definedNames></workbook>`);
    z.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    z.file("xl/worksheets/sheet1.xml", sheet);
    z.file("xl/styles.xml", styles);
    return z.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  /* ------------------------------------------------------ vettoriale */
  function attrs(r, i) {
    const v = rowOf(r, i), o = {};
    COLS.forEach((c, k) => { if (k !== LAT && k !== LON) o[c[0]] = v[k] === "" ? null : v[k]; });
    return o;
  }
  function geojson(list) {
    const fc = { type: "FeatureCollection", name: "piante",
      crs: { type: "name", properties: { name: "urn:ogc:def:crs:OGC:1.3:CRS84" } },
      features: list.filter((r) => r.posizione && r.posizione.lat).map((r, i) => ({ type: "Feature",
        geometry: { type: "Point", coordinates: [Number(r.posizione.lon.toFixed(7)), Number(r.posizione.lat.toFixed(7))] }, properties: attrs(r, i) })) };
    return new Blob([JSON.stringify(fc, null, 1)], { type: "application/geo+json" });
  }
  function kml(list, title) {
    const pm = list.filter((r) => r.posizione && r.posizione.lat).map((r, i) => {
      const a = attrs(r, i);
      const rowsH = Object.entries(a).filter(([, v]) => v !== null).map(([k, v]) => `<tr><td><b>${esc(k)}</b></td><td>${esc(v)}</td></tr>`).join("");
      return `<Placemark><name>${esc(idLabel(r))}</name><styleUrl>#pianta</styleUrl>
<description><![CDATA[<table>${rowsH}</table>]]></description>
<ExtendedData>${Object.entries(a).map(([k, v]) => `<Data name="${esc(k)}"><value>${esc(v ?? "")}</value></Data>`).join("")}</ExtendedData>
<Point><coordinates>${r.posizione.lon.toFixed(7)},${r.posizione.lat.toFixed(7)},0</coordinates></Point></Placemark>`;
    }).join("\n");
    return new Blob([`<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>${esc(title)}</name>
<Style id="pianta"><IconStyle><color>ff327d2e</color><scale>0.9</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon></IconStyle><LabelStyle><scale>0.8</scale></LabelStyle></Style>
${pm}
</Document></kml>`], { type: "application/vnd.google-earth.kml+xml" });
  }
  // stile per QGIS (si carica da solo se ha lo stesso nome del file .geojson): pallino verde bordato di bianco + etichetta
  function qml() {
    return new Blob([`<!DOCTYPE qgis PUBLIC 'http://mrcc.com/qgis.dtd' 'SYSTEM'>
<qgis version="3.28" styleCategories="Symbology|Labeling">
 <renderer-v2 type="singleSymbol" enableorderby="0" symbollevels="0" forceraster="0">
  <symbols>
   <symbol type="marker" name="0" alpha="1" clip_to_extent="1" force_rhr="0">
    <layer class="SimpleMarker" enabled="1" pass="0" locked="0">
     <Option type="Map">
      <Option name="color" type="QString" value="46,125,50,255"/>
      <Option name="name" type="QString" value="circle"/>
      <Option name="outline_color" type="QString" value="255,255,255,255"/>
      <Option name="outline_width" type="QString" value="0.6"/>
      <Option name="outline_width_unit" type="QString" value="MM"/>
      <Option name="size" type="QString" value="4"/>
      <Option name="size_unit" type="QString" value="MM"/>
     </Option>
    </layer>
   </symbol>
  </symbols>
 </renderer-v2>
</qgis>`], { type: "text/xml" });
  }


  /* ------------------------------------------------------ tabella per albx_ricoltura (programma dei ricontrolli) */
  // Formato letto da albx (gabtree.py): primo foglio, riga 1 = intestazioni, una riga per pianta.
  // Obbligatorie: Località, Numero (intero o Gr.N), Data rilievo (gg/mm/aaaa, non futura). Una sola riga sbagliata
  // fa rifiutare tutta la tabella: le piante non valide vengono escluse qui e segnalate.
  const ALBX_COLS = ["Località", "Numero", "Data rilievo", "Specie", "Frequenza", "Interventi", "Note", "Priorità", "Latitudine", "Longitudine", "Relazione"];
  // «Abies alba L. - abete bianco» -> «Abies alba»; «Platanus × acerifolia (Aiton) Willd.» -> «Platanus × acerifolia»
  function albxSpecie(s) {
    const t = String(s || "").split(/\s+[-–—]\s+/)[0].replace(/\(.*$/, "").trim().split(/\s+/);
    if (!t[0]) return "";
    const out = [t[0][0].toUpperCase() + t[0].slice(1).toLowerCase()];
    for (let i = 1; i < t.length; i++) {
      const w = t[i];
      if (/^(x|×)$/i.test(w)) { out.push("×"); continue; }
      if (/^(subsp|ssp|var|f)\.$/i.test(w) && t[i + 1] && /^[a-zà-ü-]+$/.test(t[i + 1])) { out.push(w.toLowerCase(), t[++i]); continue; }
      if (/^[a-zà-ü-]+$/.test(w) && out.length < 4) { out.push(w); continue; }
      break;                         // autore (L., Mill., Spach…), cultivar o altro testo
    }
    if (out[out.length - 1] === "×") out.pop();
    return out.join(" ");
  }
  // numero per albx: il primo valido tra progressivo, cartellino e GreenSpaces
  function albxNumero(r) {
    for (const v of [r.nprog, r.ncart, r.ngs]) {
      const t = String(v || "").trim();
      if (/^\d+$/.test(t)) return Number(t);
      const g = t.match(/^gr\.?\s*(\d+)$/i);
      if (g) return "Gr." + Number(g[1]);
    }
    return null;
  }
  // anni tra i ricontrolli: X con abbattimento, altrimenti dal «Ricontrollo visivo» (sotto l'anno -> 1), altrimenti vuota
  function albxFrequenza(r) {
    const iv = r.interventiProposti || [];
    if (iv.some((x) => /^abbattimento/i.test(x.nome))) return "X";
    const anni = iv.filter((x) => /^ricontrollo visivo/i.test(x.nome)).map((x) => {
      const m = String(x.tempo || "").match(/tra\s+(\d+)\s+ann/i);
      return m ? Number(m[1]) : x.tempo ? 1 : null;
    }).filter((n) => n);
    return anni.length ? Math.max(1, Math.min(...anni)) : "";
  }
  function albxRighe(list, localita) {
    const avvisi = [], perChiave = new Map();
    const sito = String(localita || "").replace(/\s+/g, " ").trim();
    const oggi = new Date().toISOString().slice(0, 10);
    for (const r of list) {
      const nome = r.cartella || "rilievo senza nome";
      const num = albxNumero(r);
      if (num === null) { avvisi.push(`${nome}: manca un numero valido (intero, oppure Gr.1 per un gruppo)`); continue; }
      if (!r.dataValutazione || r.dataValutazione > oggi) { avvisi.push(`${nome}: data del rilievo mancante o futura`); continue; }
      const gruppo = typeof num === "string";
      const specie = albxSpecie(r.specie) || (gruppo ? "gruppo" : "");
      const p = r.posizione || {};
      const ll = p.lat && p.lon ? [Number(p.lat.toFixed(6)), Number(p.lon.toFixed(6))] : ["", ""];
      const interventi = (r.interventiProposti || []).filter((x) => !/^ricontrollo visivo/i.test(x.nome))
        .map((x) => x.nome + (x.tempo ? ", " + x.tempo : "")).join("; ");
      let note = String(r.problematiche || "").replace(/\s+/g, " ").trim();
      if (note.length > 300) note = note.slice(0, 297).replace(/\s+\S*$/, "") + "…";
      const riga = [sito, num, itDate(r.dataValutazione), specie, albxFrequenza(r), interventi, note, "", ll[0], ll[1], ""];
      const k = String(num).toLowerCase(), prima = perChiave.get(k);
      if (prima) {
        const tieni = (r.dataValutazione || "") >= (prima.r.dataValutazione || "") ? { r, riga } : prima;
        avvisi.push(`numero ${num} presente due volte: tengo il rilievo del ${itDate(tieni.r.dataValutazione)}`);
        perChiave.set(k, tieni);
      } else perChiave.set(k, { r, riga });
      if (!specie) avvisi.push(`${nome}: specie mancante (albx la chiede per le piante nuove)`);
    }
    const righe = [...perChiave.values()].map((x) => x.riga)
      .sort((a, b) => (typeof a[1] === "string") - (typeof b[1] === "string") || String(a[1]).localeCompare(String(b[1]), "it", { numeric: true }));
    return { righe, avvisi };
  }
  // Excel semplice: un foglio, niente celle unite, niente righe di titolo
  async function albxXlsx(list, localita) {
    const { righe, avvisi } = albxRighe(list, localita);
    const colL = (i) => String.fromCharCode(65 + i);
    const cellX = (ref, v, st) => {
      if (v === "" || v === null || v === undefined) return "";
      if (typeof v === "number") return `<c r="${ref}"${st ? ` s="${st}"` : ""}><v>${v}</v></c>`;
      return `<c r="${ref}"${st ? ` s="${st}"` : ""} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
    };
    const rows = [ALBX_COLS, ...righe].map((row, k) => `<row r="${k + 1}">${row.map((v, i) => cellX(colL(i) + (k + 1), v, k === 0 ? 1 : 0)).join("")}</row>`);
    const W = [28, 8, 12, 24, 10, 50, 50, 9, 11, 11, 30];
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>${W.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${rows.join("")}</sheetData></worksheet>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const z = new JSZip();
    z.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
    z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    z.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Rilievi" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    z.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    z.file("xl/worksheets/sheet1.xml", sheet);
    z.file("xl/styles.xml", styles);
    const blob = await z.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    return { blob, n: righe.length, avvisi };
  }

  return { schedaDocx, xlsx, csv, geojson, kml, qml, staticMap, albxXlsx, albxRighe, albxSpecie };
})();
