/* Rilievi in campo – prototipo PWA
 * Bussola, GPS, fotocamera con freccia N, scheda di rilievo,
 * salvataggio offline (IndexedDB) e sincronizzazione su OneDrive (Microsoft Graph). */
"use strict";

const APP_VERSION = "0.6.0";
const CFG = window.APP_CONFIG || {};
// errori imprevisti: mostrati in rosso in alto, per capire subito cosa non va
function showErr(msg) {
  const b = document.getElementById("errBanner");
  if (!b) return;
  b.textContent = "Errore: " + msg + "\n(tocca per chiudere)";
  b.classList.remove("hidden");
  b.onclick = () => b.classList.add("hidden");
}
window.addEventListener("error", (e) => showErr(`${e.message} (${(e.filename || "").split("/").pop()}:${e.lineno})`));
window.addEventListener("unhandledrejection", (e) => showErr(String((e.reason && e.reason.message) || e.reason)));
const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));
const toRad = (d) => (d * Math.PI) / 180;
const norm360 = (d) => ((d % 360) + 360) % 360;
const CARD = ["N", "NE", "E", "SE", "S", "SO", "O", "NO"];
const cardinal = (d) => CARD[Math.round(norm360(d) / 45) % 8];

function toast(msg, ms = 2600) {
  const t = $("#toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(t._h);
  t._h = setTimeout(() => t.classList.remove("show"), ms);
}

/* ---------------------------------------------------------------- DB */
const DB = {
  db: null,
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open("rilievi-campo", 1);
      r.onupgradeneeded = () => {
        const db = r.result;
        db.createObjectStore("rilievi", { keyPath: "id" });
        db.createObjectStore("kv");
      };
      r.onsuccess = () => { this.db = r.result; res(); };
      r.onerror = () => rej(r.error);
    });
  },
  tx(store, mode, fn) {
    return new Promise((res, rej) => {
      const t = this.db.transaction(store, mode);
      const s = t.objectStore(store);
      const out = fn(s);
      t.oncomplete = () => res(out ? out.result : undefined);
      t.onerror = () => rej(t.error);
    });
  },
  put: (r) => DB.tx("rilievi", "readwrite", (s) => s.put(r)),
  del: (id) => DB.tx("rilievi", "readwrite", (s) => s.delete(id)),
  all: () => DB.tx("rilievi", "readonly", (s) => s.getAll()),
  get: (id) => DB.tx("rilievi", "readonly", (s) => s.get(id)),
  setKV: (k, v) => DB.tx("kv", "readwrite", (s) => s.put(v, k)),
  getKV: (k) => DB.tx("kv", "readonly", (s) => s.get(k)),
};

/* ---------------------------------------------------------- Settings */
const S = { basemap: "osm", prjColors: {}, prjOff: [], folder: "Rilievi", name: "", init: "", north: "T", decl: 3.5, quality: 2048, arrow: true,
  provider: "", msClientId: "", gClientId: "", setupDone: false, areteOn: false, protocollo: "SIA", gradiFoto: false, etichette: true, clino: true, mapColor: "prj", mapFilter: "",
  studio: "", qualifica: "", out: { schede: true, arete: true, foto: true, riepilogo: true, vettoriale: false, albx: false } };
async function loadSettings() {
  const v = await DB.getKV("settings");
  if (v) Object.assign(S, v);
}
const saveSettings = () => DB.setKV("settings", { ...S });

/* ----------------------------------------------------------- Bussola */
const Compass = {
  heading: null,       // gradi, già riferiti al nord scelto (T o M)
  magnetic: null,
  accuracy: null,      // iOS: webkitCompassAccuracy (gradi)
  absolute: true,
  started: false,
  _vx: 0, _vy: 0,      // media vettoriale per smorzare il rumore
  listeners: [],
  onChange(fn) { this.listeners.push(fn); },

  async start() {
    if (this.started) return true;
    try {
      if (typeof DeviceOrientationEvent !== "undefined" && typeof DeviceOrientationEvent.requestPermission === "function") {
        const p = await DeviceOrientationEvent.requestPermission();  // iPhone: richiede un tocco dell'utente
        if (p !== "granted") { toast("Permesso bussola negato"); return false; }
      }
    } catch (e) { toast("Bussola non disponibile: " + e.message); return false; }
    const h = (e) => this._handle(e);
    if ("ondeviceorientationabsolute" in window) window.addEventListener("deviceorientationabsolute", h);
    else window.addEventListener("deviceorientation", h);
    this.started = true;
    return true;
  },

  // Direzione verso cui guarda la fotocamera posteriore (telefono in verticale)
  // oppure il bordo superiore (telefono in piano), calcolata da alpha/beta/gamma.
  _fromEuler(alpha, beta, gamma) {
    const a = toRad(alpha), b = toRad(beta), g = toRad(gamma);
    const cA = Math.cos(a), sA = Math.sin(a), cB = Math.cos(b), sB = Math.sin(b), cG = Math.cos(g), sG = Math.sin(g);
    // asse -Z del telefono (fotocamera posteriore) nel sistema Est-Nord-Su
    const zx = -(cA * sG + cG * sA * sB), zy = -(sA * sG - cA * cG * sB), zz = -(cB * cG);
    // asse +Y del telefono (bordo superiore)
    const yx = -cB * sA, yy = cA * cB;
    const upright = Math.abs(zz) < 0.7;  // telefono più verticale che orizzontale
    let hd = upright ? Math.atan2(zx, zy) : Math.atan2(yx, yy);
    hd = norm360((hd * 180) / Math.PI);
    if (!upright && screen.orientation && screen.orientation.angle) hd = norm360(hd + screen.orientation.angle);
    return hd;
  },

  _handle(e) {
    let mag = null;
    if (typeof e.webkitCompassHeading === "number" && !isNaN(e.webkitCompassHeading)) {
      mag = e.webkitCompassHeading;                       // iPhone: già riferita al nord magnetico
      this.accuracy = e.webkitCompassAccuracy;
      this.absolute = true;
    } else if (e.alpha !== null && e.alpha !== undefined) {
      mag = this._fromEuler(e.alpha, e.beta || 0, e.gamma || 0);
      this.absolute = e.absolute === true || e.type === "deviceorientationabsolute";
      this.accuracy = null;
    }
    if (mag === null) return;
    // smorzamento contro il tremolio: media esponenziale sul cerchio con costante di tempo TAU,
    // indipendente da quanti eventi al secondo manda il sensore
    const now = performance.now();
    const dt = this._t ? Math.min(1, (now - this._t) / 1000) : 1;
    this._t = now;
    const TAU = 0.6;  // secondi: più alto = più stabile ma più lento
    const k = 1 - Math.exp(-dt / TAU);
    this._vx += (Math.sin(toRad(mag)) - this._vx) * k;
    this._vy += (Math.cos(toRad(mag)) - this._vy) * k;
    this.magnetic = norm360((Math.atan2(this._vx, this._vy) * 180) / Math.PI);
    this.heading = S.north === "T" ? norm360(this.magnetic + Number(S.decl || 0)) : this.magnetic;
    // aggiornamento della grafica al massimo una volta per fotogramma
    if (!this._raf) this._raf = requestAnimationFrame(() => { this._raf = 0; this.listeners.forEach((fn) => fn(this)); });
  },

  status() {
    if (!this.started) return { msg: 'Tocca "Attiva bussola"', bad: false };
    if (this.heading === null) return { msg: "In attesa del sensore…", bad: false };
    if (!this.absolute) return { msg: "Il telefono non fornisce il nord assoluto", bad: true };
    if (this.accuracy !== null && (this.accuracy < 0 || this.accuracy > 25))
      return { msg: "Bussola da calibrare: muovi il telefono a 8", bad: true };
    return { msg: `Nord ${S.north === "T" ? "geografico" : "magnetico"}${this.accuracy ? " · ±" + Math.round(this.accuracy) + "°" : ""}`, bad: false };
  },
};

function drawRose() {
  let s = '<circle r="104" fill="#fff" stroke="#2e5e3e" stroke-width="3"/>';
  for (let d = 0; d < 360; d += 10) {
    const long = d % 30 === 0, r1 = long ? 86 : 94;
    const x1 = Math.sin(toRad(d)) * r1, y1 = -Math.cos(toRad(d)) * r1, x2 = Math.sin(toRad(d)) * 102, y2 = -Math.cos(toRad(d)) * 102;
    s += `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#2e5e3e" stroke-width="${long ? 3 : 1.5}"/>`;
  }
  ["N", "E", "S", "O"].forEach((l, i) => {
    const d = i * 90, x = Math.sin(toRad(d)) * 68, y = -Math.cos(toRad(d)) * 68 + 9;
    s += `<text x="${x}" y="${y}" text-anchor="middle" font-size="26" font-weight="800" fill="${l === "N" ? "#b3261e" : "#2e5e3e"}">${l}</text>`;
  });
  s += '<polygon points="0,-48 10,0 0,8 -10,0" fill="#b3261e"/><polygon points="0,48 10,0 0,-8 -10,0" fill="#9aa59c"/>';
  $("#roseSvg").innerHTML = s;
}
function arrowSvg(rot) {
  return `<g transform="rotate(${rot})"><circle r="46" fill="rgba(0,0,0,.45)" stroke="#fff" stroke-width="3"/>
    <polygon points="0,-34 13,10 0,2 -13,10" fill="#fff"/><text x="0" y="32" text-anchor="middle" font-size="22" font-weight="800" fill="#fff">N</text></g>`;
}
function updateCompassUI(c) {
  const st = c.status();
  $("#compassInfo").textContent = st.msg;
  $("#compassInfo").className = st.bad ? "warn" : "muted";
  if (c.heading === null) return;
  // rotazione "srotolata" (niente giro completo passando da 359° a 0°)
  if (updateCompassUI.rot === undefined) updateCompassUI.rot = -c.heading;
  let d = norm360(-c.heading - updateCompassUI.rot + 180) - 180;
  updateCompassUI.rot += d;
  $("#roseSvg").style.transform = `rotate(${updateCompassUI.rot}deg)`;
  // numero mostrato: cambia solo se la differenza supera 1,5° (isteresi)
  const prev = updateCompassUI.shown;
  if (prev === undefined || Math.abs(norm360(c.heading - prev + 180) - 180) >= 1.5) updateCompassUI.shown = c.heading;
  const h = Math.round(norm360(updateCompassUI.shown)) % 360;
  $("#degBig").textContent = `${h}° ${cardinal(h)}`;
  $("#camHeading").textContent = `${h}° ${cardinal(h)}${st.bad ? " ⚠" : ""}`;
  $("#camArrow").innerHTML = arrowSvg(-c.heading);
  $("#camArrow").classList.toggle("hidden", !S.arrow);
  MapView.heading(c.heading);
}


/* ------------------------------------------------- Sfondo mappa: stradale o satellite */
// Satellite = Esri World Imagery (foto aeree/satellitari, utili per vedere le chiome dall'alto).
// Serve una chiave gratuita ArcGIS Location Platform in config.js (esriKey); senza chiave il pulsante non compare.
// Le immagini satellitari NON vengono salvate per l'offline (lo vietano i termini Esri): senza rete si torna allo stradale.
const Basemap = {
  maps: [],
  hasSat() { return !!(window.APP_CONFIG && APP_CONFIG.esriKey); },
  current() { return this.hasSat() && S.basemap === "sat" && navigator.onLine ? "sat" : "osm"; },
  layer(kind) {
    if (kind === "sat") return L.tileLayer(
      "https://ibasemaps-api.arcgis.com/arcgis/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}?token=" + encodeURIComponent(APP_CONFIG.esriKey),
      { maxNativeZoom: 19, maxZoom: 21, attribution: 'Powered by <a href="https://www.esri.com">Esri</a> · Maxar, Earthstar Geographics' });
    return L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png",
      { maxNativeZoom: 19, maxZoom: 21, attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>' });
  },
  attach(map) {
    const entry = { map, kind: null, tl: null, btn: null };
    this.maps.push(entry);
    if (this.hasSat()) {
      const Ctl = L.Control.extend({ options: { position: "topleft" }, onAdd: () => {
        const b = L.DomUtil.create("button", "basebtn");
        b.type = "button";
        L.DomEvent.disableClickPropagation(b);
        L.DomEvent.on(b, "click", (ev) => { L.DomEvent.preventDefault(ev); this.toggle(); });
        entry.btn = b; return b;
      } });
      new Ctl().addTo(map);
    }
    this.apply(entry);
  },
  apply(e) {
    const k = this.current();
    if (e.btn) { e.btn.textContent = k === "sat" ? "Mappa" : "Satellite"; e.btn.classList.toggle("on", k === "sat"); }
    if (e.kind === k) return;
    if (e.tl) e.map.removeLayer(e.tl);
    e.tl = this.layer(k).addTo(e.map); e.tl.bringToBack(); e.kind = k;
  },
  applyAll() { this.maps.forEach((e) => this.apply(e)); },
  toggle() {
    if (!navigator.onLine) { toast("Satellite non disponibile senza rete"); return; }
    S.basemap = S.basemap === "sat" ? "osm" : "sat"; saveSettings(); this.applyAll();
  },
};
window.addEventListener("online", () => Basemap.applyAll());
window.addEventListener("offline", () => Basemap.applyAll());

/* --------------------------------------------------- Colori dei progetti sulla mappa */
// Ogni progetto ha un colore (scelto in automatico, modificabile con un tocco dalla Mappa o dalla pagina del progetto).
const PALETTE = ["#e53935", "#fb8c00", "#fdd835", "#43a047", "#00acc1", "#1e88e5", "#8e24aa", "#d81b60", "#6d4c41", "#ffffff"];
const prjKey = (r) => r.lavoro || "(senza progetto)";
// "da fare" = pianta importata da un elenco, ancora da rilevare in campo (non va nel cloud e non conta nei contatori)
const isTodo = (r) => r.status === "todo";
const hasPos = (r) => r.posizione && typeof r.posizione.lat === "number";
const navUrl = (r) => `https://www.google.com/maps/dir/?api=1&destination=${r.posizione.lat},${r.posizione.lon}`;
function popupHtml(r) {
  const head = `<b>${esc(idLabel(r))}</b>${r.specie ? "<br>" + esc(r.specie) : ""}<br><span class="muted">${esc(prjKey(r))}</span><br>`;
  return isTodo(r)
    ? head + `<span class="muted">da fare${r.strumento ? " · " + esc(r.strumento) : ""}</span><div class="pop-btns">
        <a class="pop-nav" href="${navUrl(r)}" target="_blank" rel="noopener">Portami qui</a>
        <button type="button" class="pop-open" data-id="${r.id}" data-todo="1">Completa rilievo</button></div>`
    : head + `<button type="button" class="pop-open" data-id="${r.id}">Apri scheda</button>`;
}
function bindPopups(map) {
  map.on("popupopen", (e) => {
    const b = e.popup.getElement().querySelector(".pop-open");
    if (b) b.addEventListener("click", () => { map.closePopup(); b.dataset.todo ? completeTodo(b.dataset.id) : openDetail(b.dataset.id); });
  });
}
// pallino pieno = rilevato, anello vuoto = da fare
function treeMarker(r, col, on = true) {
  const ll = [r.posizione.lat, r.posizione.lon];
  if (!on) return L.circleMarker(ll, isTodo(r) ? { radius: 4, color: "#888", weight: 1.5, fill: false } : { radius: 4, color: "#777", weight: 1, fillColor: "#bbb", fillOpacity: 0.55 });
  return isTodo(r)
    ? L.circleMarker(ll, { radius: 8, color: col, weight: 4, fillColor: "#fff", fillOpacity: 0.25 })
    : L.circleMarker(ll, { radius: 8, color: "#222", weight: 1.5, fillColor: col, fillOpacity: 1 });
}
function prjColor(name) {
  if (S.prjColors && S.prjColors[name]) return S.prjColors[name];
  let h = 0; for (const c of name) h = (h * 31 + c.charCodeAt(0)) >>> 0;
  return PALETTE[h % 8];
}
function setPrjColor(name, col) {
  S.prjColors = { ...(S.prjColors || {}), [name]: col }; saveSettings();
  MapView.plot(); ProjMap.plot();
}
// riga di pallini colorati: tocco = nuovo colore
function swatches(name, onPick) {
  const cur = prjColor(name);
  return `<div class="swatches">${PALETTE.map((c) => `<button type="button" class="sw${c === cur ? " on" : ""}" data-c="${c}" style="background:${c}" aria-label="colore ${c}"></button>`).join("")}</div>`;
}
function bindSwatches(box, name, after) {
  box.querySelectorAll(".sw").forEach((b) => b.addEventListener("click", (e) => {
    e.stopPropagation(); setPrjColor(name, b.dataset.c);
    box.querySelectorAll(".sw").forEach((x) => x.classList.toggle("on", x === b));
    after && after(b.dataset.c);
  }));
}

/* -------------------------------------------------------------- Mappa */
// Leaflet + OpenStreetMap: gratuito, senza chiave. Pallino blu = posizione, cono = direzione della bussola,
// pallini verdi = rilievi salvati.
const MapView = {
  map: null, me: null, acc: null, follow: true, layer: null, centered: false,
  init() {
    if (!window.L || this.map) return;
    this.map = L.map("map", { zoomControl: false, attributionControl: true }).setView([41.9, 12.5], 6);
    Basemap.attach(this.map);
    this.layer = L.layerGroup().addTo(this.map);
    bindPopups(this.map);
    this.map.on("dragstart", () => { this.follow = false; $("#btnCenter").classList.remove("follow"); });
    $("#btnCenter").classList.add("follow");
    $("#btnCenter").addEventListener("click", () => {
      this.follow = true; $("#btnCenter").classList.add("follow");
      const f = GPS.last;
      if (f) this.map.setView([f.lat, f.lon], Math.max(this.map.getZoom(), 18));
    });
    this.plot();
  },
  icon() {
    return L.divIcon({ className: "", iconSize: [60, 60], iconAnchor: [30, 30],
      html: '<div class="me-wrap"><div class="me-cone" id="meCone"></div><div class="me-dot"></div></div>' });
  },
  update(f) {
    if (!this.map) return;
    const ll = [f.lat, f.lon];
    if (!this.me) {
      this.acc = L.circle(ll, { radius: f.acc, color: "#1a73e8", weight: 1, fillColor: "#1a73e8", fillOpacity: 0.12, interactive: false }).addTo(this.map);
      this.me = L.marker(ll, { icon: this.icon(), interactive: false, keyboard: false }).addTo(this.map);
    } else { this.me.setLatLng(ll); this.acc.setLatLng(ll).setRadius(f.acc); }
    if (!this.centered) { this.map.setView(ll, 18); this.centered = true; }
    else if (this.follow) this.map.panTo(ll, { animate: true });
    const q = f.acc <= 5 ? "ok" : f.acc <= 15 ? "" : "warn";
    $("#mapInfo").innerHTML = `${f.lat.toFixed(6)}, ${f.lon.toFixed(6)} · <span class="${q}">±${Math.round(f.acc)} m</span>`;
    $("#btnGpsOn").classList.add("hidden");
  },
  heading(h) {
    const c = document.getElementById("meCone");
    if (c) { c.style.display = "block"; c.style.transform = `rotate(${updateCompassUI.rot !== undefined ? -updateCompassUI.rot : h}deg)`; }
  },
  async plot() {
    if (!this.map) return;
    this.layer.clearLayers();
    (await DB.all()).filter(hasPos).forEach((r) => {
      treeMarker(r, prjColor(prjKey(r))).bindPopup(popupHtml(r)).addTo(this.layer);
    });
  },
  refresh() { if (this.map) { setTimeout(() => this.map.invalidateSize(), 50); this.plot(); } },
};

// colori per classe CPC (mappa «Colora per: Classe CPC»)
const CPC_COLORI = [["A", "#2e7d32"], ["B", "#9ccc65"], ["C", "#fdd835"], ["C/D", "#fb8c00"], ["D", "#d32f2f"], ["ARETÉ", "#7e57c2"], ["senza classe", "#9e9e9e"]];
function cpcColor(r) {
  if (r.protocollo === "ARETE") return "#7e57c2";
  const m = CPC_COLORI.find(([k]) => k === r.cpc);
  return m ? m[1] : "#9e9e9e";
}
// "12" oppure "4/12 fatte" se nel progetto ci sono piante da fare
function avanz(list) { const t = list.filter(isTodo).length; return t ? `${list.length - t}/${list.length} fatte` : String(list.length); }
/* ------------------------------------------------- Scheda Mappa: progetti come layer */
// Elenco dei progetti con una spunta ciascuno: i progetti spuntati si "accendono" con il loro colore,
// gli altri restano pallini grigi piccoli. Tocco su un pallino = scheda della pianta.
const ProjMap = {
  map: null, layer: null, meMk: null, open: null,
  init() {
    if (this.map || !window.L) return;
    this.map = L.map("pmapMap", { zoomControl: false, attributionControl: true }).setView([41.9, 12.5], 6);
    Basemap.attach(this.map);
    this.layer = L.layerGroup().addTo(this.map);
    bindPopups(this.map);
    if (GPS.last) this.me(GPS.last);
  },
  me(f) {
    if (!this.map || !f) return;
    const ll = [f.lat, f.lon];
    if (!this.meMk) this.meMk = L.circleMarker(ll, { radius: 7, color: "#fff", weight: 3, fillColor: "#1a73e8", fillOpacity: 1, interactive: false }).addTo(this.map);
    else this.meMk.setLatLng(ll);
  },
  isOn(name) { return !(S.prjOff || []).includes(name); },
  setOn(name, on) {
    const off = new Set(S.prjOff || []); on ? off.delete(name) : off.add(name);
    S.prjOff = [...off]; saveSettings(); this.plot();
  },
  async groups() {
    const g = {};
    (await DB.all()).forEach((r) => (g[prjKey(r)] = g[prjKey(r)] || []).push(r));
    return g;
  },
  async plot() {
    if (!this.map) return;
    this.layer.clearLayers();
    const g = await this.groups();
    const names = Object.keys(g);
    // prima gli spenti (sotto), poi gli accesi (sopra)
    names.sort((a, b) => this.isOn(a) - this.isOn(b)).forEach((n) => {
      const on = this.isOn(n), col = prjColor(n);
      g[n].filter(hasPos).forEach((r) => {
        // filtro per intervento: le piante che non lo hanno restano pallini grigi piccoli
        const pass = !S.mapFilter || (S.mapFilter === "__none" ? !(r.interventiProposti || []).length && !isTodo(r)
          : (r.interventiProposti || []).some((x) => x.nome === S.mapFilter));
        const c = S.mapColor === "cpc" ? cpcColor(r) : col;
        treeMarker(r, c, on && pass).bindPopup(popupHtml(r)).addTo(this.layer);
      });
    });
  },
  renderControls() {
    setSeg("segColor", S.mapColor || "prj");
    $("#cpcLegend").classList.toggle("hidden", S.mapColor !== "cpc");
    $("#cpcLegend").innerHTML = CPC_COLORI.map(([k, c]) => `<span><i style="background:${c}"></i>${k}</span>`).join("");
    $("#pmFilter").innerHTML = `<option value="">Tutte le piante</option><option value="__none">Senza interventi proposti</option>` +
      INTERVENTI.map((n) => `<option value="${esc(n)}">Solo con: ${esc(n)}</option>`).join("");
    $("#pmFilter").value = S.mapFilter || "";
  },
  async renderList() {
    const g = await this.groups(), box = $("#pmapList");
    const names = Object.keys(g).sort((x, y) => Math.max(...g[y].map((r) => r.createdAt)) - Math.max(...g[x].map((r) => r.createdAt)));
    box.innerHTML = names.map((n, i) => `<div class="pl" data-i="${i}">
        <label class="pl-row"><input type="checkbox" ${this.isOn(n) ? "checked" : ""}>
          <span class="pl-name">${esc(n)}</span><span class="muted">${avanz(g[n])}</span></label>
        <button type="button" class="pl-col" style="background:${prjColor(n)}" aria-label="Cambia colore"></button>
        <div class="pl-pal hidden">${swatches(n)}</div></div>`).join("");
    box.querySelectorAll(".pl").forEach((el) => {
      const n = names[+el.dataset.i];
      el.querySelector("input").addEventListener("change", (e) => this.setOn(n, e.target.checked));
      el.querySelector(".pl-col").addEventListener("click", () => {
        const pal = el.querySelector(".pl-pal"), was = pal.classList.contains("hidden");
        box.querySelectorAll(".pl-pal").forEach((x) => x.classList.add("hidden"));
        if (was) pal.classList.remove("hidden");
      });
      bindSwatches(el.querySelector(".pl-pal"), n, (c) => { el.querySelector(".pl-col").style.background = c; el.querySelector(".pl-pal").classList.add("hidden"); });
    });
  },
  locate() {
    const go = (f) => this.map && f && this.map.setView([f.lat, f.lon], Math.max(this.map.getZoom(), 18));
    const f = GPS.current ? GPS.current() : GPS.last;
    if (f) { go(f); return; }
    GPS.start(); toast("Cerco la tua posizione…");
    let n = 0; const t = setInterval(() => { const g = GPS.last; if (g || ++n > 30) { clearInterval(t); if (g) go(g); else toast("Posizione non disponibile"); } }, 500);
  },
  async fit() {
    const g = await this.groups(), pts = [];
    Object.keys(g).filter((n) => this.isOn(n)).forEach((n) => g[n].forEach((r) => hasPos(r) && pts.push([r.posizione.lat, r.posizione.lon])));
    if (pts.length) this.map.fitBounds(pts, { padding: [30, 30], maxZoom: 19 });
    else if (GPS.last) this.map.setView([GPS.last.lat, GPS.last.lon], 17);
  },
  async all(on) {
    S.prjOff = on ? [] : Object.keys(await this.groups()); saveSettings();
    this.renderList(); this.plot();
  },
  async show() {
    this.init(); if (!this.map) return;
    setTimeout(() => this.map.invalidateSize(), 60);
    this.renderControls(); await this.renderList(); await this.plot();
    if (!this.fitted) { this.fitted = true; this.fit(); }
  },
};

/* --------------------------------------------------------------- GPS */
const GPS = {
  watch: null, last: null, best: null, fixed: null,
  start() {
    if (this.watch !== null || !navigator.geolocation) return;
    this.watch = navigator.geolocation.watchPosition(
      (p) => {
        const fix = { lat: p.coords.latitude, lon: p.coords.longitude, acc: p.coords.accuracy, alt: p.coords.altitude, time: p.timestamp };
        this.last = fix;
        // la lettura migliore degli ultimi 60 s
        if (!this.best || fix.acc <= this.best.acc || fix.time - this.best.time > 60000) this.best = fix;
        updateGpsUI();
        MapView.update(fix); ProjMap.me(fix);
        formMapFollowGps(fix);
      },
      (err) => gpsError(err),
      { enableHighAccuracy: true, maximumAge: 0, timeout: 30000 }
    );
  },
  stop() { if (this.watch !== null) navigator.geolocation.clearWatch(this.watch); this.watch = null; },
  current() { return this.fixed || this.best || this.last; },
};
function gpsError(err) {
  const msg = err.code === 1
    ? "Posizione non consentita. iPhone: Impostazioni → Privacy e sicurezza → Localizzazione → Siti web di Safari → «Mentre usi l'app». Poi tocca «Attiva posizione»."
    : err.code === 3 ? "Il GPS non risponde ancora (sei al chiuso?). Riprovo…" : "GPS non disponibile: " + err.message;
  $("#gpsInfo").textContent = msg;
  $("#mapInfo").textContent = msg;
  $("#btnGpsOn").classList.toggle("hidden", err.code === 3);
  if (err.code !== 3) GPS.stop();
}
function updateGpsUI() {
  const f = GPS.current();
  if (!f) return;
  const q = f.acc <= 5 ? "ok" : f.acc <= 15 ? "" : "warn";
  $("#gpsInfo").innerHTML = `${f.lat.toFixed(6)}, ${f.lon.toFixed(6)} · <span class="${q}">±${Math.round(f.acc)} m</span>` +
    (GPS.fixed ? " · <b>fissata</b>" : " · in aggiornamento");
  const a = $("#lnkMaps");
  a.href = `https://www.google.com/maps/search/?api=1&query=${f.lat},${f.lon}`;
  a.classList.remove("hidden");
}

/* -------------------------------------------------------- Fotocamera */
const Cam = {
  stream: null,
  async open() {
    try {
      this.stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: "environment" }, width: { ideal: 4032 }, height: { ideal: 3024 } }, audio: false,
      });
    } catch (e) { toast("Fotocamera non disponibile: " + e.message); return false; }
    const v = $("#video");
    v.srcObject = this.stream;
    await v.play().catch(() => {});
    return true;
  },
  close() {
    if (this.stream) this.stream.getTracks().forEach((t) => t.stop());
    this.stream = null;
    $("#video").srcObject = null;
  },
};

function canvasFrom(source, w, h, maxSide) {
  let W = w, H = h;
  if (maxSide && Math.max(w, h) > maxSide) { const k = maxSide / Math.max(w, h); W = Math.round(w * k); H = Math.round(h * k); }
  const c = document.createElement("canvas");
  c.width = W; c.height = H;
  c.getContext("2d").drawImage(source, 0, 0, W, H);
  return c;
}
function drawNorthArrow(canvas, heading) {
  const ctx = canvas.getContext("2d");
  const r = Math.max(28, Math.round(Math.min(canvas.width, canvas.height) * 0.05));
  const cx = canvas.width - r * 1.5, cy = r * 1.5;
  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(toRad(-heading));
  ctx.beginPath(); ctx.arc(0, 0, r, 0, Math.PI * 2); ctx.fillStyle = "rgba(0,0,0,0.45)"; ctx.fill();
  ctx.lineWidth = Math.max(2, r * 0.07); ctx.strokeStyle = "#fff"; ctx.stroke();
  ctx.beginPath(); ctx.moveTo(0, -r * 0.74); ctx.lineTo(r * 0.28, r * 0.2); ctx.lineTo(0, r * 0.04); ctx.lineTo(-r * 0.28, r * 0.2); ctx.closePath();
  ctx.fillStyle = "#fff"; ctx.fill();
  ctx.font = `800 ${Math.round(r * 0.48)}px Arial`; ctx.textAlign = "center"; ctx.textBaseline = "middle";
  ctx.fillText("N", 0, r * 0.62);
  ctx.restore();
}
const canvasToDataURL = (c) => c.toDataURL("image/jpeg", 0.88);
function dataURLtoBlob(d) {
  const [h, b] = d.split(","), bin = atob(b), u = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return new Blob([u], { type: h.match(/:(.*?);/)[1] });
}
function toDMS(v) {
  const a = Math.abs(v), d = Math.floor(a), mF = (a - d) * 60, m = Math.floor(mF), s = Math.round((mF - m) * 60 * 100);
  return [[d, 1], [m, 1], [s, 100]];
}
function exifDate(t) {
  const d = new Date(t), p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}:${p(d.getMonth() + 1)}:${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}
function withExif(dataURL, meta) {
  try {
    const zeroth = {}, exif = {}, gps = {};
    zeroth[piexif.ImageIFD.Software] = "Rilievi in campo " + (CFG.appVersion || "");
    zeroth[piexif.ImageIFD.Artist] = S.name || "";
    exif[piexif.ExifIFD.DateTimeOriginal] = exifDate(meta.time);
    if (meta.lat !== undefined && meta.lat !== null) {
      gps[piexif.GPSIFD.GPSLatitudeRef] = meta.lat >= 0 ? "N" : "S";
      gps[piexif.GPSIFD.GPSLatitude] = toDMS(meta.lat);
      gps[piexif.GPSIFD.GPSLongitudeRef] = meta.lon >= 0 ? "E" : "W";
      gps[piexif.GPSIFD.GPSLongitude] = toDMS(meta.lon);
      if (meta.acc) gps[piexif.GPSIFD.GPSHPositioningError] = [Math.round(meta.acc * 10), 10];
    }
    if (meta.heading !== null && meta.heading !== undefined) {
      gps[piexif.GPSIFD.GPSImgDirectionRef] = meta.ref;            // "T" geografico, "M" magnetico
      gps[piexif.GPSIFD.GPSImgDirection] = [Math.round(meta.heading * 100), 100];
    }
    const bytes = piexif.dump({ "0th": zeroth, Exif: exif, GPS: gps });
    return piexif.insert(bytes, dataURL);
  } catch (e) { console.warn("EXIF non scritto", e); return dataURL; }
}

// Crea la foto pulita (+ versione con freccia, se attiva) a partire da un'immagine/video
function makePhotos(source, w, h, meta) {
  const maxSide = Number(S.quality) || 0;
  const clean = canvasFrom(source, w, h, maxSide);
  const out = { clean: dataURLtoBlob(withExif(canvasToDataURL(clean), meta)), arrow: null };
  if (meta.arrow && meta.heading !== null) {
    const ar = canvasFrom(clean, clean.width, clean.height, 0);
    drawNorthArrow(ar, meta.heading);
    out.arrow = dataURLtoBlob(withExif(canvasToDataURL(ar), meta));
  }
  const th = canvasFrom(clean, clean.width, clean.height, 360);
  out.thumb = dataURLtoBlob(th.toDataURL("image/jpeg", 0.7));
  return out;
}
function shotMeta() {
  const f = GPS.current() || {};
  return { time: Date.now(), lat: f.lat, lon: f.lon, acc: f.acc, heading: Compass.heading, ref: "T",
           arrow: S.arrow, compassBad: Compass.status().bad };
}

/* --------------------------------------------------------- Scheda */
let draft = null;
function newDraft() {
  return { id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2),
    createdAt: Date.now(), dataValutazione: isoToday(), status: "pending", tipoValutazione: "Valutazione",
    strumOn: false, esitoProva: "", foto: [], allegati: [], nomeV: 2, etichetteOn: S.etichette !== false };  // nomeV 2 = nome senza «Valutazione» (dalla 0.5.5)
}
function isoToday() { const d = new Date(), p = (n) => String(n).padStart(2, "0"); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`; }
// vecchi rilievi: "Fitostatica" e "Strumentale" ora sono entrambi "Valutazione"
const tipoNorm = (t) => (t === "Sorveglianza scavo" ? "Sorveglianza scavo" : "Valutazione");
function setSeg(id, v) { $$(`#${id} button`).forEach((b) => b.classList.toggle("on", b.dataset.v === String(v))); }
function bindSeg(id, fn) { $$(`#${id} button`).forEach((b) => b.addEventListener("click", () => { setSeg(id, b.dataset.v); fn(b.dataset.v); })); }

/* ---------------------------------------- Mappe piccole con il pallino della pianta */
// Pallino verde trascinabile: serve a correggere la posizione quando il GPS sotto chioma è impreciso.
function TreeMap(elId, onMove) {
  this.el = elId; this.map = null; this.pin = null; this.onMove = onMove;
}
TreeMap.prototype.ensure = function () {
  if (this.map || !window.L) return;
  this.map = L.map(this.el, { zoomControl: false, attributionControl: true }).setView([41.9, 12.5], 6);
  Basemap.attach(this.map);
};
TreeMap.prototype.set = function (lat, lon, center) {
  this.ensure();
  if (!this.map) return;
  const ll = [lat, lon];
  if (!this.pin) {
    this.pin = L.marker(ll, { draggable: true, autoPan: true,
      icon: L.divIcon({ className: "", iconSize: [28, 28], iconAnchor: [14, 14], html: '<div class="tree-pin"></div>' }) }).addTo(this.map);
    this.pin.on("dragend", () => { const p = this.pin.getLatLng(); this.onMove && this.onMove(p.lat, p.lng); });
  } else this.pin.setLatLng(ll);
  if (center) this.map.setView(ll, Math.max(this.map.getZoom(), 19));
};
TreeMap.prototype.refresh = function () { if (this.map) setTimeout(() => this.map.invalidateSize(), 60); };

const formMap = new TreeMap("formMap", (lat, lon) => {
  if (!draft) return;
  draft.posManuale = { lat, lon };
  $("#formMapHint").innerHTML = 'Posizione spostata a mano. <a href="#" id="lnkBackGps">Torna al GPS</a>';
  $("#lnkBackGps").addEventListener("click", (e) => {
    e.preventDefault(); draft.posManuale = null;
    $("#formMapHint").textContent = "Se il GPS non è preciso, trascina il pallino verde sull'albero.";
    const f = GPS.current(); if (f) formMap.set(f.lat, f.lon, true);
  });
});
function formMapFollowGps(f) {
  if (!draft || draft.posManuale || !$("#form").classList.contains("active")) return;
  formMap.set(f.lat, f.lon, !formMap._centered);
  formMap._centered = true;
}

/* --------------------------------------------------------- Scheda */
// interventi proposti (in ordine alfabetico) e tempi di prescrizione
const INTERVENTI = ["Abbattimento", "Ampliamento formella", "Potatura di contenimento", "Potatura di risanamento", "Prova di trazione",
  "Resistografia base", "Resistografia in quota", "Ricontrollo visivo", "Scavo", "Tomografia base", "Tomografia in quota"];
const TEMPI = ["entro 30 giorni", "entro 90 giorni", "entro 6 mesi", "entro 1 anno", "tra 1 anno", "tra 2 anni", "tra 3 anni", "tra 4 anni"];
// ricontrollo visivo proposto in automatico in base alla nuova CPC (protocollo SIA)
const RICONTROLLO = { A: "tra 4 anni", B: "tra 3 anni", C: "tra 2 anni" };
const CPC_CON_INTERVENTO = ["C/D", "D"];
const TEMPO_DEFAULT = { Abbattimento: "entro 30 giorni" };
let intSel = [];      // [{ nome, tempo, auto }]
let ricTolto = false; // l'utente ha tolto il ricontrollo automatico: non lo ripropongo finché non cambia la CPC
function renderInterventi(sel) {
  if (sel) { intSel = sel.map((x) => ({ ...x })); ricTolto = false; }
  intSel.sort((a, b) => INTERVENTI.indexOf(a.nome) - INTERVENTI.indexOf(b.nome));
  const box = $("#intRows");
  box.innerHTML = intSel.map((x, i) => {
    const tempi = x.tempo && !TEMPI.includes(x.tempo) ? [x.tempo, ...TEMPI] : TEMPI; // tempi delle versioni precedenti
    return `<div class="introw2"><div class="int-name">${esc(x.nome)}${x.auto ? ' <span class="muted">(da CPC)</span>' : ""}</div>
      <button type="button" class="int-del" data-i="${i}" aria-label="Togli">×</button>
      <select data-i="${i}"><option value="">— tempi di prescrizione —</option>${tempi.map((t) => `<option${x.tempo === t ? " selected" : ""}>${t}</option>`).join("")}</select></div>`;
  }).join("");
  $("#intNone").classList.toggle("hidden", intSel.length > 0);
  $("#intAdd").innerHTML = `<option value="">+ Aggiungi intervento</option>` +
    INTERVENTI.filter((n) => !intSel.some((x) => x.nome === n)).map((n) => `<option>${esc(n)}</option>`).join("");
  box.querySelectorAll("select").forEach((s) => s.addEventListener("change", () => { const x = intSel[+s.dataset.i]; x.tempo = s.value; x.auto = false; renderInterventi(); }));
  box.querySelectorAll(".int-del").forEach((b) => b.addEventListener("click", () => {
    const x = intSel[+b.dataset.i]; if (x.nome === "Ricontrollo visivo" && x.auto) ricTolto = true;
    intSel.splice(+b.dataset.i, 1); renderInterventi();
  }));
}
function addIntervento(nome) {
  if (!nome || intSel.some((x) => x.nome === nome)) return;
  intSel.push({ nome, tempo: TEMPO_DEFAULT[nome] || "" }); renderInterventi();
}
// applica la regola del ricontrollo quando cambia la CPC (o il protocollo)
function applyRicontrollo(changed) {
  if (changed) ricTolto = false;
  const cpc = (draft && draft.protocollo === "SIA") ? $("#fCpc").value : "";
  const i = intSel.findIndex((x) => x.nome === "Ricontrollo visivo");
  const t = RICONTROLLO[cpc];
  if (t) {
    if (i >= 0) { if (intSel[i].auto) intSel[i].tempo = t; }
    else if (!ricTolto) intSel.push({ nome: "Ricontrollo visivo", tempo: t, auto: true });
  } else if (i >= 0 && intSel[i].auto) intSel.splice(i, 1);
  renderInterventi();
}
function readInterventi() { return intSel.map((x) => ({ nome: x.nome, tempo: x.tempo, ...(x.auto ? { auto: true } : {}) })); }
// testo per schede e tabelle: "Scavo (entro 90 giorni); Prova di trazione (tra 1 anno)"
const intTxt = (r) => (r.interventiProposti || []).map((x) => x.nome + (x.tempo ? ` (${x.tempo})` : "")).join("; ");
function updateIdReq() { $("#idReq").classList.toggle("ok", !!($("#fNgs").value.trim() || $("#fNcart").value.trim() || $("#fNprog").value.trim())); }
let editingId = null;
// «⧉ Copia» dalla scheda di una pianta: nuovo rilievo nello stesso progetto con specie, protocollo, CPC e interventi
async function copyFrom(id) {
  const t = await DB.get(id); if (!t) return;
  await openForm(t.lavoro); if (!draft) return;
  $("#fSpecie").value = t.specie || "";
  await setProtocollo(t.protocollo || "SIA");
  $("#fCpc").value = t.protocollo === "SIA" ? t.cpc || "" : "";
  renderInterventi((t.interventiProposti || []).map((x) => ({ ...x })));
  toast(`Copiati da ${numeri(t).join(" ") || "pianta"}: specie, protocollo, CPC e interventi. Inserisci il numero della nuova pianta.`, 4500);
}
async function openForm(project) {
  if (!S.name || !S.init) { toast("Prima inserisci il nome del rilevatore e le iniziali"); openWizard(); return; }
  draft = newDraft();
  const last = (await DB.getKV("lastLavoro")) || {};
  editingId = null; $("#form header h1").textContent = "Nuovo rilievo";
  ["fNgs", "fNcart", "fNprog", "fSpecie", "fDiam", "fAlt", "fProb", "fInt", "fCpc"].forEach((i) => ($("#" + i).value = ""));
  renderInterventi([]); updateIdReq();
  $("#fData").value = draft.dataValutazione;
  await fillLavoroSel(project || last.lavoro || "");
  $("#formWho").textContent = `Rilevatore: ${S.name} (${S.init})`;
  draft.tipoValutazione = "Valutazione";
  updateNamePreview();
  $("#btnClino").classList.toggle("hidden", !S.clino);
  GPS.fixed = null; GPS.best = null; GPS.start();
  $("#gpsInfo").textContent = "In attesa del GPS…";
  $("#formMapHint").textContent = "Se il GPS non è preciso, trascina il pallino verde sull'albero.";
  $("#lnkMaps").classList.add("hidden");
  $("#formErr").textContent = ""; $("#dupWarn").classList.add("hidden");
  renderThumbs();
  await fillLavori();
  if (window.ARETE_READY) ARETE.reset($("#areteBox"), (await DB.getKV("areteRemember")) || {});
  await setProtocollo(S.protocollo === "ARETE" ? "ARETE" : "SIA");
  show("form");
  formMap._centered = false;
  formMap.ensure(); formMap.refresh();
  setTimeout(() => { const f = GPS.current(); if (f) { updateGpsUI(); formMapFollowGps(f); } }, 150);
  updateNamePreview();
}
// Apre la scheda già compilata: pianta "da fare" importata da un elenco, oppure rilievo esistente da modificare
async function loadIntoForm(t, edit) {
  await openForm(t.lavoro);
  if (!draft) return;
  draft = { ...draft, ...t, id: t.id, createdAt: edit ? t.createdAt : Date.now(), status: "pending", foto: t.foto || [], allegati: t.allegati || [], daElenco: !edit };
  delete draft.posizione;
  if (edit) { editingId = t.id; $("#form header h1").textContent = "Modifica rilievo"; }
  const set = (i, v) => ($("#" + i).value = v ?? "");
  set("fNgs", t.ngs); set("fNcart", t.ncart); set("fNprog", t.nprog); set("fSpecie", t.specie); set("fDiam", t.diametro120); set("fAlt", t.altezza);
  set("fProb", t.problematiche); set("fInt", t.interventi); set("fCpc", t.cpc);
  renderInterventi(t.interventiProposti || []); updateIdReq();
  await setProtocollo(t.protocollo || (edit ? "SIA" : S.protocollo === "ARETE" ? "ARETE" : "SIA"));
  if (edit && t.arete && window.ARETE_READY) ARETE.fill($("#areteBox"), t.arete);
  set("fData", t.dataValutazione || isoToday());
  draft.tipoValutazione = "Valutazione";
  if (hasPos(t)) {
    draft.posManuale = { lat: t.posizione.lat, lon: t.posizione.lon, keep: edit ? t.posizione : null };
    formMap.set(t.posizione.lat, t.posizione.lon, true); formMap._centered = true;
    $("#formMapHint").innerHTML = (edit ? "Posizione salvata nel rilievo. " : "Posizione presa dall'elenco importato. ") + '<a href="#" id="lnkBackGps">Usa la mia posizione GPS</a>';
    $("#lnkBackGps").addEventListener("click", (e) => {
      e.preventDefault(); draft.posManuale = null;
      $("#formMapHint").textContent = "Se il GPS non è preciso, trascina il pallino verde sull'albero.";
      const f = GPS.current(); if (f) formMap.set(f.lat, f.lon, true);
    });
  }
  renderThumbs(); updateNamePreview(); checkDuplicate();
}
async function completeTodo(id) { const t = await DB.get(id); if (t) await loadIntoForm(t, false); }
async function editRecord(id) { const t = await DB.get(id); if (t) await loadIntoForm(t, true); }
// protocollo di valutazione: SIA (scelta della CPC) oppure ARETÉ (scheda ARETÉ)
async function setProtocollo(v, user) {
  setSeg("segProtocollo", v);
  if (draft) draft.protocollo = v;
  const tpl = await DB.getKV("areteTemplate");
  $("#boxSia").classList.toggle("hidden", v !== "SIA");
  $("#boxArete").classList.toggle("hidden", v !== "ARETE");
  $("#areteMissing").classList.toggle("hidden", !!tpl);
  $("#areteBox").classList.toggle("hidden", !tpl);
  if (user) applyRicontrollo(false);
}

// menu dei progetti nel nuovo rilievo: esistenti (più recenti in alto) oppure «+ Nuovo progetto…»
async function fillLavoroSel(cur) {
  const all = await DB.all(), last = {};
  all.forEach((r) => { if (r.lavoro) last[r.lavoro] = Math.max(last[r.lavoro] || 0, r.createdAt || 0); });
  ((await DB.getKV("projects")) || []).forEach((n) => { if (!(n in last)) last[n] = Date.now(); });
  const names = Object.keys(last).sort((a, b) => last[b] - last[a]);
  const sel = $("#fLavoroSel");
  sel.innerHTML = `<option value="">— scegli il progetto —</option>` + names.map((n) => `<option>${esc(n)}</option>`).join("") + `<option value="__new">+ Nuovo progetto…</option>`;
  if (cur && names.includes(cur)) { sel.value = cur; showNewLavoro(false, cur); }
  else if (cur) { sel.value = "__new"; showNewLavoro(true, cur); }
  else if (!names.length) { sel.value = "__new"; showNewLavoro(true, ""); }
  else { sel.value = ""; showNewLavoro(false, ""); }
}
function showNewLavoro(on, val) {
  const inp = $("#fLavoro");
  inp.classList.toggle("hidden", !on);
  inp.value = val || "";
}
async function fillLavori() {
  const all = await DB.all();
  const names = [...new Set([...all.map((r) => r.lavoro), ...((await DB.getKV("projects")) || [])].filter(Boolean))];
  $("#lavoriList").innerHTML = names.map((n) => `<option value="${esc(n)}">`).join("");
}
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// etichette delle foto (finiscono nella didascalia della scheda Word: «Figura 3 – Colletto»).
// Ordine: vista generale, poi le parti dal basso verso l'alto, poi anomalia e testo libero.
const ETICHETTE = ["Vista generale", "Radici", "Colletto", "Fusto", "Castello", "Branca", "Ramo", "Chioma", "Anomalia", "Altro"];
// piccolo albero stilizzato, sempre uguale, con evidenziata la parte fotografata
function lblIcon(nome) {
  const n = nome === "Difetto" ? "Anomalia" : nome;   // i rilievi vecchi con «Difetto»
  if (n === "Anomalia") return `<svg viewBox="0 0 24 24" class="lico" aria-hidden="true"><path d="M12 3 L22.5 21 H1.5 Z" fill="#e8590c" stroke="#fff" stroke-width="1.2" stroke-linejoin="round"/><rect x="11" y="9" width="2" height="6.5" rx="1" fill="#fff"/><circle cx="12" cy="18" r="1.2" fill="#fff"/></svg>`;
  if (!ETICHETTE.includes(n) || n === "Altro") return "";
  const all = n === "Vista generale", on = (k) => all || n === k;
  const G = all ? "#5b8a63" : "#cdd5cf", H = "#0f7d33", Cg = all ? "#a9d8a9" : "#eef2ee", Ch = "#5cbf68";
  const col = (k) => (on(k) ? H : G), w = (k) => (on(k) && !all ? 3 : 1.6);
  return `<svg viewBox="0 0 24 24" class="lico" aria-hidden="true">
    <ellipse cx="12" cy="8.2" rx="9" ry="6.8" fill="${on("Chioma") ? Ch : Cg}"/>
    <path d="M12 12.5 L6.5 7.5 M12 12.5 L17.5 7.5 M12 12.5 L12 4.5" stroke="${col("Branca")}" stroke-width="${w("Branca")}" stroke-linecap="round" fill="none"/>
    <path d="M8.6 9.4 L5 9.2 M15.4 9.4 L19 9.2 M12 7 L9.6 4.4" stroke="${col("Ramo")}" stroke-width="${w("Ramo")}" stroke-linecap="round" fill="none"/>
    <rect x="${on("Fusto") && !all ? 10.3 : 10.8}" y="12.5" width="${on("Fusto") && !all ? 3.4 : 2.4}" height="6.5" fill="${col("Fusto")}"/>
    <rect x="${on("Colletto") && !all ? 9.2 : 10.1}" y="${on("Colletto") && !all ? 16.8 : 17.6}" width="${on("Colletto") && !all ? 5.6 : 3.8}" height="${on("Colletto") && !all ? 3 : 2.2}" rx=".8" fill="${col("Colletto")}"/>
    <circle cx="12" cy="12.6" r="${on("Castello") && !all ? 2.8 : 1.3}" fill="${col("Castello")}"/>
    <path d="M3.5 19.8 H20.5" stroke="#8a7a66" stroke-width="1" stroke-linecap="round"/>
    <path d="M12 19.8 L7.5 23 M12 19.8 L16.5 23 M12 19.8 L12 23.4" stroke="${col("Radici")}" stroke-width="${w("Radici")}" stroke-linecap="round" fill="none"/></svg>`;
}
const lblNome = (e) => (e === "Difetto" ? "Anomalia" : e);
// icona sulla foto; per «Altro» niente icona, il testo libero sotto
function lblBadge(f, on) {
  if (!on || !f.etichetta) return "";
  if (f.etichettaAltro) return "";
  const ic = lblIcon(f.etichetta);
  return ic ? `<span class="lbadge" title="${esc(lblNome(f.etichetta))}">${ic}</span>` : "";
}
const LblPick = {
  i: -1,
  open(i) {
    this.i = i;
    const cur = draft.foto[i] || {}, sel = cur.etichettaAltro ? "Altro" : lblNome(cur.etichetta || "");
    $("#lblGrid").innerHTML = ETICHETTE.map((e) => `<button type="button" data-e="${e}" class="${sel === e ? "on" : ""}">${e === "Altro" ? `<span class="lico ltxt">Aa</span>` : lblIcon(e)}<span>${e}</span></button>`).join("");
    $("#lblGrid").querySelectorAll("button").forEach((b) => b.addEventListener("click", () => this.set(b.dataset.e)));
    $("#lblPick").classList.remove("hidden");
  },
  set(e) {
    const f = draft.foto[this.i];
    if (f) {
      if (e === "Altro") { if (!f.etichettaAltro) f.etichetta = ""; f.etichettaAltro = true; }
      else { f.etichetta = e; delete f.etichettaAltro; }
    }
    this.close();
    renderThumbs();
    if (e === "Altro") { const t = $(`#thumbs .ltext[data-i="${this.i}"]`); if (t) t.focus(); }
  },
  clear() { const f = draft.foto[this.i]; if (f) { f.etichetta = ""; delete f.etichettaAltro; } this.close(); renderThumbs(); },
  close() { $("#lblPick").classList.add("hidden"); },
  bind() {
    $("#lblNone").addEventListener("click", () => this.clear());
    $("#lblClose").addEventListener("click", () => this.close());
    $("#lblPick").addEventListener("click", (ev) => { if (ev.target.id === "lblPick") this.close(); });
  },
};
function renderThumbs() {
  const box = $("#thumbs");
  box.innerHTML = "";
  const lon = draft.etichetteOn !== false;
  draft.foto.forEach((p, i) => {
    const d = document.createElement("div");
    d.className = "th";
    const url = URL.createObjectURL(p.thumb);
    const chip = !lon ? "" : p.etichettaAltro
      ? `<input class="ltext" data-i="${i}" maxlength="40" placeholder="scrivi l'etichetta…" value="${esc(p.etichetta || "")}"><button type="button" class="lchip lsmall" data-i="${i}">cambia</button>`
      : `<button type="button" class="lchip" data-i="${i}">${p.etichetta ? lblIcon(p.etichetta) + `<span>${esc(lblNome(p.etichetta))}</span>` : "<span>＋ etichetta</span>"}</button>`;
    d.innerHTML = `<div class="pw"><img src="${url}" data-i="${i}">${lblBadge(p, lon)}</div><button class="del" data-i="${i}">×</button>
      ${chip}
      ${S.gradiFoto ? `<div class="cap">${p.heading !== null && p.heading !== undefined ? Math.round(p.heading) + "° " + cardinal(p.heading) : "senza bussola"}${p.compassBad ? " ⚠" : ""}</div>` : ""}`;
    box.appendChild(d);
  });
  box.querySelectorAll("img").forEach((im) => im.addEventListener("click", () => Viewer.open(draft.foto, Number(im.dataset.i))));
  box.querySelectorAll(".lchip").forEach((b) => b.addEventListener("click", () => LblPick.open(Number(b.dataset.i))));
  box.querySelectorAll(".ltext").forEach((t) => t.addEventListener("input", () => { draft.foto[Number(t.dataset.i)].etichetta = t.value.trim(); }));
  box.querySelectorAll(".del").forEach((b) => b.addEventListener("click", () => { draft.foto.splice(Number(b.dataset.i), 1); renderThumbs(); }));
  $("#camCount").textContent = `${draft.foto.length} foto`;
}

/* --------------------------------------------- Nomi dei rilievi e numeri */
const TIPO_NOME = { "Valutazione": "Valutazione", "Fitostatica": "Valutazione", "Strumentale": "Valutazione", "Sorveglianza scavo": "Sorveglianza" };
// tutti i numeri scritti nel rilievo: gs. (Greenspaces), cart. (cartellino), rel. (progressivo)
function numeri(r) {
  return [r.ngs ? "gs." + r.ngs : "", r.ncart ? "cart." + r.ncart : "", r.nprog ? "rel." + r.nprog : ""]
    .filter(Boolean).map((x) => x.replace(/[^\w.]+/g, ""));
}
function idLabel(r) { const n = numeri(r); return n.length ? n.join(" ") : (r.cartella || "—"); }
// "Parco Rossi – Comune X" -> "ParcoRossi-ComuneX"
function compatto(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[–—]/g, "-")
    .split(/[^A-Za-z0-9-]+/).filter(Boolean).map((w) => w[0].toUpperCase() + w.slice(1)).join("")
    .replace(/-+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "SenzaProgetto";
}
function ymd(t) { const d = new Date(t), p = (n) => String(n).padStart(2, "0"); return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`; }
// YYYYMMDD_Tipo_numeri_progetto_XX   es. 20261120_Valutazione_gs.4521-cart.118_ParcoRossi-ComuneX_VA
function nomeRilievo(r) {
  const data = (r.dataValutazione || "").replace(/-/g, "") || ymd(r.createdAt);
  // dalla 0.5.5 i rilievi nuovi non hanno più il tipo nel nome; quelli vecchi lo tengono (così non cambiano nome se modificati)
  const tipo = r.nomeV >= 2 ? [] : [TIPO_NOME[r.tipoValutazione] || "Rilievo"];
  return [data, ...tipo, numeri(r).join("-") || "sn", compatto(r.lavoro), (r.iniziali || "XX")].join("_");
}
function updateNamePreview() {
  if (!draft) return;
  const tmp = { ...draft, dataValutazione: $("#fData").value, lavoro: $("#fLavoro").value.trim(), ngs: $("#fNgs").value.trim(), ncart: $("#fNcart").value.trim(),
    nprog: $("#fNprog").value.trim(), iniziali: (S.init || "XX").toUpperCase() };
  $("#nameprev").textContent = "Nome del rilievo: " + nomeRilievo(tmp);
}
const safe = (s) => String(s).normalize("NFC").replace(/[\\/:*?"<>|#%]+/g, "-").replace(/\s+/g, " ").trim().slice(0, 80);

async function saveForm() {
  const v = (i) => $("#" + i).value.trim();
  const r = draft;
  Object.assign(r, {
    dataValutazione: v("fData") || isoToday(), lavoro: v("fLavoro"), ngs: v("fNgs"), ncart: v("fNcart"), nprog: v("fNprog"), specie: v("fSpecie"),
    diametro120: v("fDiam"), altezza: v("fAlt"), problematiche: v("fProb"), note: "", interventi: v("fInt"),
    tipoValutazione: "Valutazione", protocollo: r.protocollo || "SIA", interventiProposti: readInterventi(),
    operatore: S.name, iniziali: (S.init || "XX").toUpperCase(), appVersion: APP_VERSION,
  });
  const err = [];
  if (!r.lavoro) err.push("progetto");
  if (!r.ngs && !r.ncart && !r.nprog) err.push("almeno un numero (gs, cart o progressivo)");
  if (r.protocollo === "SIA" && CPC_CON_INTERVENTO.includes(v("fCpc")) && !r.interventiProposti.length) err.push(`almeno un intervento proposto (obbligatorio con CPC ${v("fCpc")})`);
  if (err.length) { $("#formErr").textContent = "Mancano: " + err.join(", "); return; }

  const f = GPS.current();
  const gps = f ? { lat: f.lat, lon: f.lon, acc: f.acc, alt: f.alt, time: f.time, fissata: !!GPS.fixed } : null;
  const pm = r.posManuale;
  if (pm && pm.keep && pm.keep.lat === pm.lat && pm.keep.lon === pm.lon) r.posizione = pm.keep; // modifica: posizione invariata
  else r.posizione = pm ? { lat: pm.lat, lon: pm.lon, acc: null, manuale: true, gps } : gps;
  delete r.posManuale;
  // nome univoco: se esiste già lo stesso nome aggiunge -2, -3… (in modifica resta quello di prima se i dati del nome non cambiano)
  const oldName = editingId ? r.cartella : "";
  const base = nomeRilievo(r), used = new Set((await DB.all()).filter((x) => x.id !== r.id).map((x) => x.cartella));
  let name = base, k = 2;
  if (oldName && oldName.replace(/-\d+$/, "") === base && !used.has(oldName)) name = oldName;
  else while (used.has(name)) name = `${base}-${k++}`;
  r.cartella = name;
  r.foto.forEach((p, i) => {
    const n = String(i + 1).padStart(2, "0");
    p.name = `foto_${n}.jpg`;
    p.nameN = p.arrow ? `foto_${n}_N.jpg` : null;
  });
  r.cpc = r.protocollo === "SIA" ? v("fCpc") : "";
  r.arete = null; r.areteXlsx = null;
  const tpl = await DB.getKV("areteTemplate");
  if (r.protocollo === "ARETE" && tpl && window.ARETE_READY) {
    r.arete = ARETE.collect($("#areteBox"));
    // dalla 0.6.0 non si compila più la scheda ufficiale: il foglio con i dati si crea al momento dell'invio
    if (r.arete.ubicaz) await DB.setKV("areteRemember", { ubicaz: r.arete.ubicaz });
  }
  r.status = "pending"; delete r.daElenco; r.error = "";
  if (editingId) { r.modificatoIl = Date.now(); r.onlyData = false; } // dopo una modifica si reinvia tutto, foto comprese
  await DB.put(r);
  await DB.setKV("lastLavoro", { lavoro: r.lavoro });
  S.protocollo = r.protocollo; saveSettings();
  const wasEdit = editingId; editingId = null;
  draft = null;
  GPS.fixed = null;
  if (wasEdit) {
    toast((oldName && oldName !== r.cartella ? "Rilievo modificato. Ha un nome nuovo: i file con il vecchio nome restano nel cloud." : "Rilievo modificato.") + " Tocca «Sincronizza» per aggiornarlo nel cloud.", 5000);
    openDetail(r.id);
  } else { toast("Rilievo salvato sul telefono. Per inviarlo al cloud tocca «Sincronizza ora».", 4000); show("home"); }
  refreshCounts();
}

async function checkDuplicate() {
  if (!draft) return;
  updateNamePreview();
  const lav = $("#fLavoro").value.trim().toLowerCase();
  const gs = $("#fNgs").value.trim(), cart = $("#fNcart").value.trim(), rel = $("#fNprog").value.trim();
  const all = (await DB.all()).filter((r) => (r.lavoro || "").toLowerCase() === lav && !isTodo(r) && r.id !== draft.id);
  const hit = [];
  if (gs && all.some((r) => r.ngs === gs)) hit.push("gs." + gs);
  if (cart && all.some((r) => r.ncart === cart)) hit.push("cart." + cart);
  if (rel && all.some((r) => String(r.nprog) === rel)) hit.push("rel." + rel);
  $("#dupWarn").textContent = hit.length ? `Attenzione: in questo progetto esiste già un rilievo con ${hit.join(", ")}.` : "";
  $("#dupWarn").classList.toggle("hidden", !hit.length);
}

/* --------------------------------------------------- File in uscita */
const num = (s) => (s === "" || s === undefined || s === null ? null : Number(String(s).replace(",", ".")));
// cartella del progetto: YYYYMMDD_Cliente_Valutazione (data del primo rilievo), decisa una volta e poi fissa
async function projectFolder(name) {
  const map = (await DB.getKV("projFolders")) || {};
  if (map[name]) return map[name];
  const list = (await DB.all()).filter((r) => r.lavoro === name).sort((a, b) => (a.dataValutazione || "").localeCompare(b.dataValutazione || "") || a.createdAt - b.createdAt);
  const first = list[0] || {};
  // cartelle nuove (dalla 0.5.5): data_Progetto, senza «_Valutazione»; quelle già create restano come sono
  map[name] = [(first.dataValutazione || isoToday()).replace(/-/g, ""), compatto(name)].join("_");
  await DB.setKV("projFolders", map);
  return map[name];
}
const fotoName = (r, i, arrow) => `${r.cartella}_${String(i + 1).padStart(2, "0")}${arrow ? "_N" : ""}.jpg`;
// file di una pianta, secondo le spunte: schede (Word + ARETÉ + allegati) e foto
async function treeFiles(r, o, onlyData) {
  const out = [];
  if (o.schede) {
    out.push({ path: `Schede/${r.cartella}.docx`, blob: await OUT.schedaDocx(r) });
    if (!onlyData) r.allegati.filter((a) => a.blob).forEach((a) => out.push({ path: `Schede/${r.cartella}_allegato_${safe(a.name)}`, blob: a.blob }));
  }
  // foglio con i dati ARETÉ (Sezione | Voce | Valore), per compilare poi la scheda ufficiale
  if (o.arete !== false && r.protocollo === "ARETE" && r.arete && window.ARETE_READY) out.push({ path: `Fogli ARETE/${r.cartella}_dati_ARETE.xlsx`, blob: await ARETE.dataXlsx(r) });
  if (o.foto && !onlyData) r.foto.forEach((f, i) => {
    if (f.ridotta) return; // copia leggera dopo «Libera spazio»: l'originale è già nel cloud
    out.push({ path: "Foto/" + fotoName(r, i), blob: f.clean });
    if (f.arrow) out.push({ path: "Foto/" + fotoName(r, i, true), blob: f.arrow });
  });
  return out;
}
// file del progetto: riepilogo (Excel + CSV) e, a richiesta, il vettoriale per QGIS / Google Earth
async function projectFiles(name, o) {
  const list = (await DB.all()).filter((r) => r.lavoro === name && !isTodo(r))
    .sort((a, b) => (a.dataValutazione || "").localeCompare(b.dataValutazione || "") || a.createdAt - b.createdAt);
  const c = compatto(name), out = [];
  if (o.riepilogo) {
    out.push({ path: `Riepilogo_${c}.xlsx`, blob: await OUT.xlsx(list, name) });
    out.push({ path: `Riepilogo_${c}.csv`, blob: OUT.csv(list) });
  }
  if (o.vettoriale) {
    out.push({ path: `Mappa/${c}.geojson`, blob: OUT.geojson(list) });
    out.push({ path: `Mappa/${c}.qml`, blob: OUT.qml() });
    out.push({ path: `Mappa/${c}.kml`, blob: OUT.kml(list, name) });
  }
  if (o.albx && list.length) {
    // tabella per albx_ricoltura: Località = nome del progetto; il file viene sovrascritto a ogni invio
    const t = await OUT.albxXlsx(list, name);
    if (t.n) out.push({ path: `albx_${c}.xlsx`, blob: t.blob });
    if (t.avvisi.length) albxAvvisi.push(`«${name}»: ` + t.avvisi.join("; "));
  }
  return out;
}
let albxAvvisi = [];   // piante escluse dalla tabella per albx (mostrate dopo la sincronizzazione)
const OUT_LABELS = [["schede", "Schede di valutazione"], ["arete", "Foglio per scheda ARETÉ"], ["foto", "Foto"],
  ["riepilogo", "Tabella riassuntiva"], ["vettoriale", "File per QGIS e Google Earth"], ["albx", "Tabella per albx"]];
function renderOutChecks(box, o, onChange) {
  if (o.arete === undefined) o.arete = true;
  box.innerHTML = OUT_LABELS.map(([k, t]) => `<label><input type="checkbox" data-o="${k}" ${o[k] ? "checked" : ""}><span>${t}</span></label>`).join("");
  box.querySelectorAll("input").forEach((i) => i.addEventListener("change", () => onChange(i.dataset.o, i.checked)));
}

/* ------------------------------------------------- OneDrive (Graph) */
const msClientId = () => S.msClientId || CFG.clientId || "";
const Auth = {
  msal: null, account: null,
  configured() { return !!(msClientId() && window.msal); },
  async init() {
    if (!this.configured()) return;
    this.msal = new msal.PublicClientApplication({
      auth: { clientId: msClientId(), authority: CFG.authority, redirectUri: location.origin + location.pathname },
      cache: { cacheLocation: "localStorage" },
    });
    await this.msal.initialize();
    try {
      const res = await this.msal.handleRedirectPromise();
      if (res && res.account) this.msal.setActiveAccount(res.account);
    } catch (e) { toast("Accesso Microsoft non riuscito: " + e.message, 6000); }
    this.account = this.msal.getActiveAccount() || this.msal.getAllAccounts()[0] || null;
    if (this.account) this.msal.setActiveAccount(this.account);
  },
  async login() {
    if (!this.configured()) { toast("Manca il codice dell'app Microsoft", 5000); return; }
    if (!this.msal) await this.init();
    await this.msal.loginRedirect({ scopes: CFG.scopes, prompt: "select_account" });
  },
  async logout() { if (this.msal) await this.msal.logoutRedirect({ account: this.account }); },
  connected() { return !!this.account; },
  label() { return this.account ? this.account.username : ""; },
  async token(interactive) {
    if (!this.account) throw new Error("OneDrive non collegato");
    try {
      return (await this.msal.acquireTokenSilent({ scopes: CFG.scopes, account: this.account })).accessToken;
    } catch (e) {
      if (interactive && e instanceof msal.InteractionRequiredAuthError) await this.msal.acquireTokenRedirect({ scopes: CFG.scopes });
      throw new Error("Accesso Microsoft scaduto: tocca Sincronizza ora");
    }
  },
};

const G = "https://graph.microsoft.com/v1.0";
const enc = (path) => path.split("/").map(encodeURIComponent).join("/");
const OneDrive = {
  name: "OneDrive",
  async fetch(url, opt = {}, interactive) {
    const t = await Auth.token(interactive);
    const go = () => fetch(url, { ...opt, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } });
    let r = await go();
    if (r.status === 429 || r.status >= 500) { await new Promise((ok) => setTimeout(ok, 2000)); r = await go(); }
    return r;
  },
  async upload(path, blob, interactive) {
    if (blob.size < 4 * 1024 * 1024) {
      const r = await this.fetch(`${G}/me/drive/root:/${enc(path)}:/content`, { method: "PUT", body: blob }, interactive);
      if (!r.ok) throw new Error(`Upload ${path}: ${r.status}`);
      return;
    }
    const s = await this.fetch(`${G}/me/drive/root:/${enc(path)}:/createUploadSession`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }),
    }, interactive);
    if (!s.ok) throw new Error(`Sessione upload ${path}: ${s.status}`);
    const { uploadUrl } = await s.json();
    const CH = 10 * 327680;
    for (let i = 0; i < blob.size; i += CH) {
      const part = blob.slice(i, Math.min(i + CH, blob.size));
      const r = await fetch(uploadUrl, { method: "PUT", body: part, headers: { "Content-Range": `bytes ${i}-${i + part.size - 1}/${blob.size}` } });
      if (!r.ok && r.status !== 202) throw new Error(`Upload ${path}: ${r.status}`);
    }
  },
  async readText(path, interactive) {
    const r = await this.fetch(`${G}/me/drive/root:/${enc(path)}:/content`, {}, interactive);
    if (r.status === 404) return null;
    if (!r.ok) throw new Error("Lettura " + path + ": " + r.status);
    return r.text();
  },
};

/* --------------------------------------------- Google Drive (OAuth) */
const gClientId = () => S.gClientId || CFG.googleClientId || "";
const GOOGLE_SCOPE = "https://www.googleapis.com/auth/drive.file openid email";
const GAuth = {
  tok: null,  // { access, exp, email }
  configured() { return !!gClientId(); },
  load() { try { this.tok = JSON.parse(localStorage.getItem("gtoken") || "null"); } catch (e) { this.tok = null; } },
  save() { localStorage.setItem("gtoken", JSON.stringify(this.tok)); },
  // ritorno dalla pagina di accesso Google: il token arriva nell'indirizzo (#access_token=…)
  async handleRedirect() {
    if (!location.hash.includes("access_token") || !location.hash.includes("state=gdrive")) {
      if (location.hash.includes("state=gdrive") && location.hash.includes("error")) {
        history.replaceState(null, "", location.pathname);
        toast("Accesso Google non completato", 5000);
      }
      return false;
    }
    const p = new URLSearchParams(location.hash.slice(1));
    history.replaceState(null, "", location.pathname);
    this.tok = { access: p.get("access_token"), exp: Date.now() + (Number(p.get("expires_in")) - 60) * 1000, email: (this.tok && this.tok.email) || "" };
    try {
      const r = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", { headers: { Authorization: "Bearer " + this.tok.access } });
      if (r.ok) this.tok.email = (await r.json()).email || this.tok.email;
    } catch (e) { /* offline */ }
    this.save();
    return true;
  },
  login(silent) {
    if (!this.configured()) { toast("Manca il codice dell'app Google", 5000); return; }
    const q = new URLSearchParams({
      client_id: gClientId(), redirect_uri: location.origin + location.pathname, response_type: "token",
      scope: GOOGLE_SCOPE, include_granted_scopes: "true", state: "gdrive",
      prompt: silent ? "none" : "select_account",
    });
    if (this.tok && this.tok.email) q.set("login_hint", this.tok.email);
    location.href = "https://accounts.google.com/o/oauth2/v2/auth?" + q.toString();
  },
  logout() { this.tok = null; localStorage.removeItem("gtoken"); },
  connected() { return !!(this.tok && this.tok.access); },
  label() { return this.tok ? this.tok.email || "account Google" : ""; },
  token(interactive) {
    if (!this.tok) throw new Error("Google Drive non collegato");
    if (Date.now() < this.tok.exp) return this.tok.access;
    if (interactive) { this.login(true); throw new Error("Rinnovo accesso Google…"); }
    throw new Error("Accesso Google scaduto: tocca Sincronizza ora");
  },
};
const GD = "https://www.googleapis.com/drive/v3";
const GDU = "https://www.googleapis.com/upload/drive/v3";
const GoogleDrive = {
  name: "Google Drive",
  folders: {},
  async fetch(url, opt = {}, interactive) {
    const t = GAuth.token(interactive);
    const r = await fetch(url, { ...opt, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } });
    if (r.status === 401) { GAuth.tok.exp = 0; GAuth.save(); throw new Error("Accesso Google scaduto: tocca Sincronizza ora"); }
    return r;
  },
  q: (s) => s.replace(/\\/g, "\\\\").replace(/'/g, "\\'"),
  async find(name, parent, folder, interactive) {
    const q = `name='${this.q(name)}' and '${parent}' in parents and trashed=false` +
      (folder ? " and mimeType='application/vnd.google-apps.folder'" : "");
    const r = await this.fetch(`${GD}/files?q=${encodeURIComponent(q)}&fields=files(id,name)&spaces=drive`, {}, interactive);
    if (!r.ok) throw new Error("Ricerca su Drive: " + r.status);
    const j = await r.json();
    return j.files && j.files[0] ? j.files[0].id : null;
  },
  async folder(parts, interactive) {
    let parent = "root", key = "";
    for (const name of parts) {
      key += "/" + name;
      if (!this.folders[key]) {
        let id = await this.find(name, parent, true, interactive);
        if (!id) {
          const r = await this.fetch(`${GD}/files?fields=id`, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ name, mimeType: "application/vnd.google-apps.folder", parents: [parent] }),
          }, interactive);
          if (!r.ok) throw new Error("Creazione cartella " + name + ": " + r.status);
          id = (await r.json()).id;
        }
        this.folders[key] = id;
      }
      parent = this.folders[key];
    }
    return parent;
  },
  split(path) { const p = path.split("/"); return { dirs: p.slice(0, -1), name: p[p.length - 1] }; },
  async upload(path, blob, interactive) {
    const { dirs, name } = this.split(path);
    const parent = await this.folder(dirs, interactive);
    const existing = await this.find(name, parent, false, interactive);
    const meta = existing ? {} : { name, parents: [parent] };
    const url = existing ? `${GDU}/files/${existing}?uploadType=resumable` : `${GDU}/files?uploadType=resumable`;
    const s = await this.fetch(url, {
      method: existing ? "PATCH" : "POST",
      headers: { "Content-Type": "application/json; charset=UTF-8", "X-Upload-Content-Type": blob.type || "application/octet-stream" },
      body: JSON.stringify(meta),
    }, interactive);
    if (!s.ok) throw new Error(`Upload ${name}: ${s.status}`);
    const loc = s.headers.get("Location");
    const r = await fetch(loc, { method: "PUT", body: blob });
    if (!r.ok) throw new Error(`Upload ${name}: ${r.status}`);
  },
  async readText(path, interactive) {
    const { dirs, name } = this.split(path);
    const parent = await this.folder(dirs, interactive);
    const id = await this.find(name, parent, false, interactive);
    if (!id) return null;
    const r = await this.fetch(`${GD}/files/${id}?alt=media`, {}, interactive);
    if (!r.ok) throw new Error("Lettura " + name + ": " + r.status);
    return r.text();
  },
};

/* ------------------------------------------- Cartella locale del dispositivo */
// Chrome/Edge su computer: scrive direttamente nella cartella scelta (File System Access API).
// iPhone e Android non lo permettono: l'app prepara uno ZIP e lo si salva con "Condividi → Salva su File".
const Local = {
  name: "Cartella locale",
  handle: null, zip: null, files: 0,
  get dirMode() { return typeof window.showDirectoryPicker === "function"; },
  get zipMode() { return !this.dirMode; },
  async load() { try { this.handle = (await DB.getKV("localDir")) || null; } catch (e) { this.handle = null; } },
  configured() { return true; },
  connected() { return this.zipMode || !!this.handle; },
  label() { return this.zipMode ? "sul telefono (ZIP da salvare con Condividi → Salva su File)" : this.handle ? "cartella «" + this.handle.name + "»" : ""; },
  async login() {
    if (this.zipMode) { toast("Su questo dispositivo i file vengono salvati come ZIP"); renderSettings(); return; }
    try {
      this.handle = await window.showDirectoryPicker({ mode: "readwrite" });
      await DB.setKV("localDir", this.handle);
      toast("Cartella scelta: " + this.handle.name);
    } catch (e) { /* annullato */ }
    renderSettings();
  },
  async logout() { this.handle = null; await DB.setKV("localDir", null); renderSettings(); },
  // permesso di scrittura: va chiesto con un tocco dell'utente
  async ensure(interactive) {
    if (this.zipMode) return true;
    if (!this.handle) throw new Error("Cartella locale non scelta");
    let st = await this.handle.queryPermission({ mode: "readwrite" });
    if (st !== "granted" && interactive) st = await this.handle.requestPermission({ mode: "readwrite" });
    if (st !== "granted") throw new Error("Permesso sulla cartella da rinnovare: tocca Sincronizza ora");
    return true;
  },
  begin() { if (this.zipMode) { this.zip = new JSZip(); this.files = 0; } },
  async upload(path, blob) {
    if (this.zipMode) { this.zip.file(path, blob); this.files++; return; }
    const parts = path.split("/"), name = parts.pop();
    let dir = this.handle;
    for (const d of parts) dir = await dir.getDirectoryHandle(d, { create: true });
    const fh = await dir.getFileHandle(name, { create: true });
    const w = await fh.createWritable();
    await w.write(blob); await w.close();
  },
  async readText() { return null; },
  // ZIP pronto: si salva con un tocco (il menu Condividi va aperto da un gesto dell'utente)
  async finish() {
    if (!this.zipMode || !this.zip || !this.files) return;
    const blob = await this.zip.generateAsync({ type: "blob" });
    const file = new File([blob], `Rilievi_${isoToday().replace(/-/g, "")}.zip`, { type: "application/zip" });
    this.zip = null;
    showZipReady(file, this.files);
  },
};
function showZipReady(file, n) {
  const box = $("#zipReady");
  box.classList.remove("hidden");
  $("#zipInfo").textContent = `${file.name} · ${n} file · ${(file.size / 1048576).toFixed(1)} MB`;
  $("#btnZipClose").onclick = () => box.classList.add("hidden");
  $("#btnZipSave").onclick = async () => {
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: file.name }); box.classList.add("hidden"); return; } catch (e) { if (e.name === "AbortError") return; }
    }
    const a = document.createElement("a"); a.href = URL.createObjectURL(file); a.download = file.name; a.click();
    box.classList.add("hidden");
  };
}

const provider = () => (S.provider === "google" ? GoogleDrive : S.provider === "onedrive" ? OneDrive : S.provider === "local" ? Local : null);
const providerAuth = () => (S.provider === "google" ? GAuth : S.provider === "onedrive" ? Auth : S.provider === "local" ? Local : null);

const Sync = {
  running: false,
  // invia i rilievi in attesa; poi rigenera i file dei progetti toccati
  async run(manual = false, only = null) {   // only = { ids, projects }: sincronizza solo i selezionati
    if (this.running) return;
    const P = provider(), A = providerAuth();
    if (!P || !A.connected()) { setSyncInfo("Archivio non collegato (Impostazioni)"); if (manual) toast("Collega OneDrive o Google Drive nelle Impostazioni"); return; }
    if (P === Local) {
      if (Local.zipMode && !manual) { setSyncInfo("Tocca «Sincronizza ora» per preparare lo ZIP da salvare sul telefono"); return; }
      try { await Local.ensure(manual); } catch (e) { setSyncInfo(e.message); return; }
    } else if (!navigator.onLine) { if (manual) toast("Nessuna connessione: i rilievi restano sul telefono"); return; }
    const pending = (await DB.all()).filter((r) => r.status !== "synced" && !isTodo(r) && (!only || only.ids.has(String(r.id))));
    // progetti con tabella riassuntiva / vettoriale mancanti o non aggiornati (es. spunte cambiate dopo l'invio)
    const stale = (await staleProjects()).filter((n) => !only || only.projects.has(n));
    if (!pending.length && !stale.length) { if (manual) toast("Tutto sincronizzato"); setSyncInfo(`Tutto sincronizzato su ${P.name}`); return; }
    this.running = true;
    let ok = 0, lastErr = "", filesOk = 0;
    const touched = new Set();
    if (P === Local) Local.begin();
    for (const r of pending) {
      try {
        setSyncInfo(`Invio ${r.cartella} su ${P.name}…`);
        const base = `${safe(S.folder || "Rilievi")}/${await projectFolder(r.lavoro)}`;
        for (const f of await treeFiles(r, S.out, r.onlyData)) await P.upload(`${base}/${f.path}`, f.blob, manual);
        r.status = "synced"; r.syncedAt = Date.now(); r.error = null; r.archivio = P.name; r.onlyData = false;
        touched.add(r.lavoro);
        ok++;
        await DB.put(r);
      } catch (e) {
        r.status = "error"; r.error = lastErr = errIt(e);
        await DB.put(r);
        if (/scadut|Rinnovo|non collegato/.test(lastErr)) break;
      }
    }
    stale.forEach((n) => touched.add(n));
    const sigs = (await DB.getKV("projSig")) || {};
    albxAvvisi = [];
    for (const name of touched) {
      try {
        setSyncInfo(`Aggiorno i file del progetto ${name}…`);
        const base = `${safe(S.folder || "Rilievi")}/${await projectFolder(name)}`;
        for (const f of await projectFiles(name, S.out)) await P.upload(`${base}/${f.path}`, f.blob, manual);
        sigs[name] = await projSig(name);
        filesOk++;
      } catch (e) { lastErr = errIt(e); }
    }
    await DB.setKV("projSig", sigs);
    if (P === Local) await Local.finish();
    this.running = false;
    const left = (await DB.all()).filter((r) => r.status !== "synced" && !isTodo(r)).length;
    const ora = new Date().toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
    if (lastErr) setSyncInfo(`${ok ? ok + (ok === 1 ? " rilievo inviato" : " rilievi inviati") + ", " : ""}${left ? left + " da reinviare · " : ""}Errore: ${lastErr}`);
    else if (!ok && filesOk) setSyncInfo(`Tabella e vettoriale aggiornati su ${P.name} alle ${ora}`);
    else setSyncInfo(`Sincronizzato su ${P.name} alle ${ora}`);
    if (albxAvvisi.length) setSyncInfo($("#syncInfo").textContent + "\n\nTabella per albx, piante escluse o da controllare: " + albxAvvisi.join(" · "));
    refreshCounts();
  },
  // "Crea / aggiorna i file scelti" dalla pagina del progetto: rifà tutto il progetto con le spunte scelte
  async exportProject(name, o) {
    const P = provider(), A = providerAuth();
    if (!P || !A.connected()) { toast("Collega OneDrive o Google Drive nelle Impostazioni"); return; }
    if (P === Local) { try { await Local.ensure(true); } catch (e) { toast(e.message); return; } }
    else if (!navigator.onLine) { toast("Serve la connessione"); return; }
    if (this.running) { toast("Sincronizzazione in corso, riprova tra poco"); return; }
    this.running = true;
    if (P === Local) Local.begin();
    const info = (t) => ($("#prjExportInfo").textContent = t);
    try {
      const base = `${safe(S.folder || "Rilievi")}/${await projectFolder(name)}`;
      const list = (await DB.all()).filter((r) => r.lavoro === name && !isTodo(r));
      let i = 0;
      for (const r of list) {
        info(`Pianta ${++i} di ${list.length}…`);
        for (const f of await treeFiles(r, o, false)) await P.upload(`${base}/${f.path}`, f.blob, true);
        if (r.status !== "synced") { r.status = "synced"; r.syncedAt = Date.now(); r.error = null; await DB.put(r); }
      }
      info("File del progetto…");
      albxAvvisi = [];
      for (const f of await projectFiles(name, o)) await P.upload(`${base}/${f.path}`, f.blob, true);
      if (P === Local) await Local.finish();
      if (o.riepilogo === !!S.out.riepilogo && o.vettoriale === !!S.out.vettoriale && !!o.albx === !!S.out.albx) { const sg = (await DB.getKV("projSig")) || {}; sg[name] = await projSig(name); await DB.setKV("projSig", sg); }
      info(`Fatto: ${base}` + (albxAvvisi.length ? `\nTabella per albx, piante escluse o da controllare: ${albxAvvisi.join(" · ")}` : ""));
      toast("File del progetto aggiornati");
    } catch (e) { info("Errore: " + errIt(e)); }
    this.running = false;
    refreshCounts();
  },
};
// "firma" dei file di progetto: cambia se cambiano le spunte, il numero di piante o un invio
async function projSig(name) {
  const list = (await DB.all()).filter((r) => r.lavoro === name && !isTodo(r));
  return JSON.stringify([!!S.out.riepilogo, !!S.out.vettoriale, !!S.out.albx, list.length, Math.max(0, ...list.map((r) => r.syncedAt || 0))]);
}
async function staleProjects() {
  if (!S.out.riepilogo && !S.out.vettoriale && !S.out.albx) return [];
  const sigs = (await DB.getKV("projSig")) || {}, all = await DB.all(), out = [];
  for (const name of new Set(all.filter((r) => !isTodo(r) && r.status === "synced" && r.lavoro).map((r) => r.lavoro)))
    if (sigs[name] !== (await projSig(name))) out.push(name);
  return out;
}
function setSyncInfo(t) { $("#syncInfo").textContent = t; }
/* ------------------------------------------------- Controllo dei file dell'app (fase di sviluppo) */
// Verifica che sul sito ci siano tutti i pezzi dell'app: se ne manca uno (es. un file non caricato su GitHub)
// lo segnala in Impostazioni con il nome del file da ricaricare.
const FILE_CHECK = {
  libs: [["vendor/jszip.min.js", () => window.JSZip, "riepilogo Excel, ARETÉ, importazione, ZIP"],
    ["vendor/docx.min.js", () => window.docx, "schede Word"],
    ["vendor/leaflet.js", () => window.L, "mappe"],
    ["vendor/piexif.js", () => window.piexif, "dati GPS e direzione nelle foto"],
    ["vendor/msal-browser.min.js", () => window.msal, "accesso OneDrive"],
    ["config.js", () => window.APP_CONFIG, "codici di OneDrive, Google ed Esri"]],
  files: [["arete_lists.json", "elenchi ARETÉ e specie"], ["specie.json", "elenco delle specie"], ["brand/logo_orizzontale.png", "logo delle schede Word"],
    ["modello/Modello_rilievi_GabTree.xlsx", "modello Excel da scaricare"], ["vendor/leaflet.css", "aspetto delle mappe"], ["install.js", "istruzioni di installazione"],
    ["import.js", "importazione degli elenchi"], ["output.js", "schede, tabelle e vettoriale"], ["arete.js", "scheda ARETÉ"], ["manifest.webmanifest", "installazione"]],
  missing: null,
};
async function checkAppFiles() {
  const miss = FILE_CHECK.libs.filter(([, ok]) => !ok()).map(([f, , use]) => [f, use]);
  if (navigator.onLine) {
    await Promise.all(FILE_CHECK.files.map(async ([f, use]) => {
      try { const r = await fetch(f + "?check=" + Date.now(), { method: "HEAD", cache: "no-store" }); if (!r.ok) miss.push([f, use]); }
      catch (e) { /* rete instabile: non segnalo */ }
    }));
  }
  FILE_CHECK.missing = miss;
  const box = $("#fileCheck");
  box.classList.toggle("hidden", !miss.length);
  box.innerHTML = miss.length ? `<b>⚠️ Mancano dei file dell'app</b><div class="muted" style="margin-top:4px">Ricaricali su GitHub nella stessa cartella indicata:</div>
    <ul class="ins-list">${miss.map(([f, use]) => `<li><code>${esc(f)}</code> – serve per: ${esc(use)}</li>`).join("")}</ul>` : "";
  return miss;
}
// errori più comprensibili nella riga sotto «Sincronizza ora»
function errIt(e) {
  const m = String((e && e.message) || e || "");
  if (/Can't find variable|is not defined|undefined is not|reading '/.test(m)) return "manca un file dell'app (vedi Impostazioni)";
  if (/Failed to fetch|Load failed|NetworkError|network/i.test(m)) return "connessione assente o interrotta";
  if (/QuotaExceeded|quota/i.test(m)) return "spazio esaurito sul telefono o nel cloud";
  return m;
}


/* ----------------------------------------------------- Lista/dettaglio */
async function refreshCounts() {
  const all = (await DB.all()).filter((r) => !isTodo(r));
  const nPend = all.filter((r) => r.status !== "synced").length;
  $("#pendingCount").textContent = nPend;
  $("#btnSync").classList.toggle("attn", nPend > 0);
  // contatore delle valutazioni (data della valutazione)
  const now = Date.now(), day = 86400000;
  const age = (r) => (now - new Date((r.dataValutazione || isoToday()) + "T12:00:00").getTime()) / day;
  $("#cntTot").textContent = all.length;
  $("#cntW").textContent = all.filter((r) => age(r) < 7).length;
  $("#cntM").textContent = all.filter((r) => age(r) < 30).length;
  $("#cntY").textContent = all.filter((r) => age(r) < 365).length;
}
const STATO = { pending: ["in attesa", "pending"], synced: ["inviato", "synced"], error: ["errore", "error"], todo: ["da fare", "todo"] };
// Progetti = rilievi raggruppati per cliente/località, come le cartelle sul cloud
async function renderList() {
  const all = await DB.all();
  const box = $("#listBox");
  const groups = {};
  all.forEach((r) => { const k = r.lavoro || "(senza progetto)"; (groups[k] = groups[k] || []).push(r); });
  const empty = ((await DB.getKV("projects")) || []).filter((n) => !groups[n]);
  box.classList.toggle("hidden", !all.length && !empty.length);
  const names = Object.keys(groups).sort((x, y) => Math.max(...groups[y].map((r) => r.createdAt)) - Math.max(...groups[x].map((r) => r.createdAt)));
  box.innerHTML = "";
  names.forEach((n) => {
    const list = groups[n], last = Math.max(...list.map((r) => r.createdAt)), pend = list.filter((r) => r.status !== "synced" && !isTodo(r)).length;
    const todo = list.filter(isTodo).length;
    const withFoto = list.find((r) => r.foto && r.foto[0]);
    const d = document.createElement("div");
    d.className = "item";
    d.innerHTML = `<img src="${withFoto ? URL.createObjectURL(withFoto.foto[0].thumb) : "icons/icon-192.png"}"><div class="t"><b>${esc(n)}</b>
      <span class="muted">${list.length} ${list.length === 1 ? "pianta" : "piante"} · ultimo ${new Date(last).toLocaleDateString("it-IT")}</span></div>
      <span class="badge ${todo ? "todo" : pend ? "pending" : "synced"}">${todo ? todo + " da fare" : pend ? pend + " da inviare" : "inviato"}</span>`;
    const sel = selBox(d, n);
    d.addEventListener("click", () => (sel ? Sel.toggle(n) : openProject(n)));
    box.appendChild(d);
  });
  empty.reverse().forEach((n) => {
    const d = document.createElement("div");
    d.className = "item";
    d.innerHTML = `<img src="icons/icon-192.png"><div class="t"><b>${esc(n)}</b><span class="muted">nuovo progetto · nessuna pianta</span></div>`;
    const sel = selBox(d, n);
    d.addEventListener("click", () => (sel ? Sel.toggle(n) : openProject(n)));
    box.prepend(d);
  });
  Sel.all = [...empty, ...names];
  $("#btnSelList").textContent = Sel.mode === "list" ? "Fine" : "Seleziona";
  $("#btnSelList").classList.toggle("hidden", !Sel.all.length);
  if (Sel.mode === "list") Sel.bar();
}
let currentProject = null;
async function openProject(name) {
  currentProject = name;
  $("#prjTitle").textContent = name;
  const list = (await DB.all()).filter((r) => (r.lavoro || "(senza progetto)") === name)
    .sort((a, b) => isTodo(b) - isTodo(a) || b.createdAt - a.createdAt);
  const nTodo = list.filter(isTodo).length;
  $("#prjProg").textContent = nTodo ? `${list.length - nTodo} di ${list.length} piante rilevate · ${nTodo} da fare` : "";
  $("#btnPrjTodoDel").classList.toggle("hidden", !nTodo);
  const box = $("#prjBox");
  box.innerHTML = "";
  list.forEach((r) => {
    const d = document.createElement("div");
    d.className = "item";
    const st = STATO[r.status] || ["", ""];
    const classe = isTodo(r) ? "tocca per completare il rilievo" : r.protocollo === "ARETE" ? "Protocollo ARETÉ" : r.cpc ? "Classe " + r.cpc : "Classe —";
    d.innerHTML = `<img src="${r.foto[0] ? URL.createObjectURL(r.foto[0].thumb) : "icons/icon-192.png"}"><div class="t"><b>${esc(numeri(r).join(" · ") || "—")}</b>
      <span class="muted">${esc(classe)}</span></div><span class="badge ${st[1]}">${st[0]}</span>`;
    const sel = selBox(d, String(r.id));
    if (r.foto.some((f) => f.ridotta)) d.querySelector(".muted").insertAdjacentHTML("beforeend", " · foto alleggerite");
    d.addEventListener("click", () => (sel ? Sel.toggle(String(r.id)) : isTodo(r) ? completeTodo(r.id) : openDetail(r.id)));
    box.appendChild(d);
  });
  Sel.all = list.map((r) => String(r.id));
  $("#btnSelPrj").textContent = Sel.mode === "prj" ? "Fine" : "Seleziona";
  $("#btnSelPrj").classList.toggle("hidden", !list.length);
  if (Sel.mode === "prj") Sel.bar();
  const pc = $("#prjColor"); pc.innerHTML = swatches(name); bindSwatches(pc, name);
  $("#prjExportInfo").textContent = "";
  show("project");
}
/* ------------------------------------------------- Seleziona: sincronizza, libera spazio, elimina */
// Funziona sia nell'elenco dei progetti (seleziono progetti) sia dentro un progetto (seleziono rilievi).
const Sel = {
  mode: null, keys: new Set(), all: [],
  on(mode) { this.mode = mode; this.keys.clear(); this.bar(); },
  off() { this.mode = null; this.keys.clear(); $("#selBar").classList.add("hidden"); document.body.classList.remove("selecting"); },
  toggle(k) { this.keys.has(k) ? this.keys.delete(k) : this.keys.add(k); this.bar(); },
  bar() {
    const n = this.keys.size, what = this.mode === "list" ? (n === 1 ? "progetto" : "progetti") : (n === 1 ? "rilievo" : "rilievi");
    $("#selBar").classList.remove("hidden"); document.body.classList.add("selecting");
    $("#selCount").textContent = n ? `${n} ${what} selezionat${n === 1 ? "o" : "i"}` : `Tocca per selezionare`;
    $("#selAll").textContent = n && n === this.all.length ? "Nessuno" : "Tutti";
    ["selSync", "selFree", "selDel"].forEach((id) => ($("#" + id).disabled = !n));
    document.querySelectorAll(".item[data-key]").forEach((d) => {
      const on = this.keys.has(d.dataset.key); d.classList.toggle("sel", on);
      const c = d.querySelector(".selbox"); if (c) c.checked = on;
    });
  },
  async records() {
    const all = await DB.all();
    return this.mode === "list" ? all.filter((r) => this.keys.has(r.lavoro || "(senza progetto)")) : all.filter((r) => this.keys.has(String(r.id)));
  },
  refresh(mode = this.mode) { mode === "list" ? renderList() : mode === "prj" ? openProject(currentProject) : null; },
};
// selettore all'inizio di una riga
function selBox(d, key) {
  d.dataset.key = key;
  if (!Sel.mode) return false;
  d.insertAdjacentHTML("afterbegin", `<input type="checkbox" class="selbox"${Sel.keys.has(key) ? " checked" : ""} tabindex="-1">`);
  d.classList.toggle("sel", Sel.keys.has(key));
  return true;
}
// riduce le foto di un rilievo già inviato a una copia leggera (1200 px): basta per la scheda Word,
// gli originali restano nel cloud e non vengono più ricaricati (non verrebbero sovrascritti da copie ridotte)
async function riduciFoto(blob) {
  const bmp = await createImageBitmap(blob);
  const k = Math.min(1, 1200 / Math.max(bmp.width, bmp.height));
  const c = document.createElement("canvas"); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
  c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
  return new Promise((res) => c.toBlob(res, "image/jpeg", 0.8));
}
async function liberaSpazio(list) {
  let n = 0, saved = 0;
  for (const r of list) {
    if (r.status !== "synced" || isTodo(r)) continue;
    let did = false;
    for (const f of r.foto) {
      if (f.ridotta) continue;
      const before = (f.clean ? f.clean.size : 0) + (f.arrow ? f.arrow.size : 0);
      if (f.arrow) { f.arrow = await riduciFoto(f.arrow); f.clean = null; } // la scheda usa la foto con la freccia
      else if (f.clean) f.clean = await riduciFoto(f.clean);
      f.ridotta = true; did = true;
      saved += before - ((f.clean ? f.clean.size : 0) + (f.arrow ? f.arrow.size : 0));
    }
    (r.allegati || []).forEach((a) => { if (a.blob) { saved += a.blob.size || 0; a.blob = null; a.suCloud = true; did = true; } });
    if (did) { await DB.put(r); n++; }
  }
  return { n, mb: saved / 1048576 };
}
async function selAction(kind) {
  const list = await Sel.records();
  if (!list.length) return;
  const done = list.filter((r) => !isTodo(r));
  const mai = done.filter((r) => r.status !== "synced"), todo = list.filter(isTodo);
  if (kind === "sync") {
    const projects = new Set(done.map((r) => r.lavoro));
    if (!done.length) { toast("Nei selezionati ci sono solo piante da fare: non c'è niente da inviare"); return; }
    const mode = Sel.mode; Sel.off(); Sel.refresh(mode);
    await Sync.run(true, { ids: new Set(done.map((r) => String(r.id))), projects });
    Sel.refresh(mode);
    return;
  }
  if (kind === "free") {
    const ok = done.filter((r) => r.status === "synced");
    if (!ok.length) { toast("Libera spazio funziona solo sui rilievi già inviati: prima sincronizza"); return; }
    if (!confirm(`Libera spazio su ${ok.length} ${ok.length === 1 ? "rilievo inviato" : "rilievi inviati"}?\n\nSul telefono le foto restano in una copia leggera (basta per rifare la scheda Word) e gli allegati vengono tolti. Gli originali restano nel cloud. I dati, la mappa e la tabella non cambiano.` + (mai.length ? `\n\n${mai.length} non ancora inviati verranno lasciati come sono.` : ""))) return;
    toast("Riduco le foto…");
    const res = await liberaSpazio(ok);
    const mode = Sel.mode; Sel.off(); Sel.refresh(mode);
    toast(res.n ? `Spazio liberato: circa ${res.mb.toFixed(res.mb < 10 ? 1 : 0)} MB su ${res.n} rilievi` : "Questi rilievi erano già alleggeriti", 4500);
    return;
  }
  // elimina
  let msg = Sel.mode === "list"
    ? `Eliminare dal telefono ${Sel.keys.size === 1 ? "il progetto" : Sel.keys.size + " progetti"} (${list.length} ${list.length === 1 ? "pianta" : "piante"})?`
    : `Eliminare dal telefono ${list.length} ${list.length === 1 ? "rilievo" : "rilievi"}?`;
  const nInv = done.length - mai.length;
  if (nInv) msg += `\n\n${nInv === 1 ? "1 rilievo già inviato: resta" : nInv + " rilievi già inviati: restano"} nel cloud, ${nInv === 1 ? "lo togli" : "li togli"} solo dal telefono.`;
  if (todo.length) msg += `\n${todo.length === 1 ? "1 pianta da fare: si può" : todo.length + " piante da fare: si possono"} reimportare dall'elenco.`;
  if (Sel.mode === "prj" && done.length - mai.length) {
    const rest = (await DB.all()).filter((r) => r.lavoro === currentProject && !isTodo(r) && !Sel.keys.has(String(r.id)));
    if (rest.length) msg += `\n\nAttenzione: la tabella e il vettoriale del progetto nel cloud vengono rifatti con le piante rimaste sul telefono, quindi alla prossima sincronizzazione queste righe sparirebbero. Per recuperare spazio senza perdere righe usa «Libera spazio».`;
  }
  if (!confirm(msg)) return;
  if (mai.length && !confirm(`⚠️ ${mai.length} ${mai.length === 1 ? "rilievo non è mai stato inviato" : "rilievi non sono mai stati inviati"} al cloud.\nSe ${mai.length === 1 ? "lo elimini andrà perso" : "li elimini andranno persi"} per sempre, foto comprese.\n\nEliminarli comunque?`)) return;
  for (const r of list) await DB.del(r.id);
  if (Sel.mode === "list") {
    const names = [...Sel.keys];
    const prj = ((await DB.getKV("projects")) || []).filter((n) => !names.includes(n)); await DB.setKV("projects", prj);
    const sig = (await DB.getKV("projSig")) || {}; names.forEach((n) => delete sig[n]); await DB.setKV("projSig", sig);
  }
  const mode = Sel.mode; Sel.off();
  toast(`Eliminati ${list.length} ${list.length === 1 ? "elemento" : "elementi"} dal telefono`, 3500);
  refreshCounts(); MapView.plot();
  if (mode === "prj") { const left = (await DB.all()).some((r) => r.lavoro === currentProject); left ? openProject(currentProject) : show("list"); }
  else renderList();
}

let current = null, detPos = null;
const detMap = new TreeMap("detMap", (lat, lon) => { detPos = { lat, lon }; $("#btnSavePos").classList.remove("hidden"); });
async function openDetail(id) {
  current = await DB.get(id);
  detPos = null;
  $("#btnSavePos").classList.add("hidden");
  const r = current, p = r.posizione || {};
  $("#detTitle").textContent = idLabel(r);
  const pos = p.lat ? `${p.lat.toFixed(6)}, ${p.lon.toFixed(6)}` + (p.manuale ? " (corretta a mano)" : p.acc ? ` (±${Math.round(p.acc)} m)` : "") : "—";
  const kv = [["Nome", r.cartella], ["Data", r.dataValutazione ? r.dataValutazione.split("-").reverse().join("/") : ""], ["Progetto", r.lavoro],
    ["Tipo di rilievo", r.tipoValutazione === "Sorveglianza scavo" ? "Sorveglianza scavo" : ""], ["Numeri", numeri(r).join("  ")],
    ["Rilevatore", r.operatore], ["Specie", r.specie], ["Diametro 120 cm", r.diametro120 ? r.diametro120 + " cm" : ""], ["Altezza", r.altezza ? r.altezza + " m" : ""],
    ["Posizione", pos], ["Note di campo", r.problematiche], ["Protocollo", r.protocollo === "ARETE" ? "ARETÉ" : r.protocollo || ""], ["Nuova CPC", r.cpc],
    ["Interventi proposti", intTxt(r) || "Nessun intervento proposto"], ["Note sugli interventi", r.interventi],
    ["Indagine strumentale", r.strumOn ? `${r.strumento || "sì"}${r.esitoProva ? " · esito " + r.esitoProva.toLowerCase() : ""}` : ""],
    ["Stato", (STATO[r.status] || [r.status])[0] + (r.error ? " – " + r.error : "")]];
  $("#detKv").innerHTML = kv.filter((x) => x[1]).map(([k, v]) => `<div>${k}</div><div>${esc(v)}</div>`).join("");
  const lon = r.etichetteOn !== false;
  $("#detThumbs").innerHTML = r.foto.map((f, i) => `<div class="th"><div class="pw"><img src="${URL.createObjectURL(f.thumb || f.arrow || f.clean)}" data-i="${i}">${lblBadge(f, lon)}</div>
    ${lon && f.etichetta && f.etichettaAltro ? `<div class="cap ltxtcap">${esc(f.etichetta)}</div>` : ""}
    ${S.gradiFoto && f.heading !== null && f.heading !== undefined ? `<div class="cap">${Math.round(f.heading)}° ${cardinal(f.heading)}</div>` : ""}</div>`).join("");
  $("#detThumbs").querySelectorAll("img").forEach((im) => im.addEventListener("click", () => Viewer.open(r.foto, Number(im.dataset.i))));
  show("detail");
  detMap.ensure(); detMap.refresh();
  if (p.lat) setTimeout(() => detMap.set(p.lat, p.lon, true), 150);
}
async function saveDetailPos() {
  if (!current || !detPos) return;
  const old = current.posizione || {};
  current.posizione = { lat: detPos.lat, lon: detPos.lon, acc: null, manuale: true, gps: old.manuale ? old.gps : old };
  current.status = "pending";
  current.onlyData = current.syncedAt ? true : false;   // già inviato: si reinviano solo dati e registro
  await DB.put(current);
  detPos = null;
  $("#btnSavePos").classList.add("hidden");
  toast("Posizione aggiornata");
  MapView.plot(); refreshCounts();
  openDetail(current.id);
}
async function shareCurrent() {
  toast("Preparo la scheda…");
  const files = (await treeFiles(current, { schede: true, foto: true }, false))
    .map((f) => new File([f.blob], f.path.split("/").pop(), { type: f.blob.type || "application/octet-stream" }));
  if (navigator.canShare && navigator.canShare({ files })) {
    try { await navigator.share({ files, title: current.cartella }); } catch (e) { /* annullato */ }
  } else toast("Condivisione di file non supportata da questo browser");
}
async function deleteCurrent() {
  const msg = current.status === "synced" ? "Eliminare il rilievo dal telefono? Su OneDrive resta." :
    "Questo rilievo NON è ancora su OneDrive. Eliminarlo definitivamente?";
  if (!confirm(msg)) return;
  await DB.del(current.id);
  current = null;
  show("list");
  refreshCounts();
}

/* ------------------------------------------------------- Navigazione */
function show(id) {
  $$(".screen").forEach((s) => s.classList.toggle("active", s.id === id));
  const tab = id === "project" || id === "detail" || id === "imp" ? "list" : id;
  if (id === "home") refreshCounts();
  $$("nav.tabs button").forEach((b) => b.classList.toggle("on", b.dataset.go === tab));
  $("nav.tabs").classList.toggle("hidden", id === "camera" || id === "welcome");
  if (id === "list") renderList();
  if (id === "home") MapView.refresh();
  if (id === "pmap") ProjMap.show();
  if (Sel.mode && !((Sel.mode === "list" && id === "list") || (Sel.mode === "prj" && id === "project"))) Sel.off();
  if (id === "settings") renderSettings();
  window.scrollTo(0, 0);
}
async function renderSettings() {
  $("#sFolder").value = S.folder; $("#sName").value = S.name; $("#sInit").value = S.init; $("#sDecl").value = S.decl;
  setSeg("segQuality", [2048, 3000, 0].includes(Number(S.quality)) ? Number(S.quality) : 2048); setSeg("segProv", S.provider); setSeg("segGradi", S.gradiFoto ? "1" : "0");
  setSeg("segEtichette", S.etichette ? "1" : "0"); setSeg("segClino", S.clino ? "1" : "0"); setSeg("segArrowSet", S.arrow ? "1" : "0");
  $("#sStudio").value = S.studio || ""; $("#sQual").value = S.qualifica || "";
  renderOutChecks($("#setOut"), S.out, (k, v) => { S.out[k] = v; saveSettings(); });
  const A = providerAuth();
  $("#acctInfo").textContent = !A ? "" :
    !A.configured() ? "Manca il codice dell'app: riapri la configurazione iniziale" :
    A.connected() ? `${provider().name}: ${A.label()}` : `${provider().name} non collegato`;
  $("#btnLogin").textContent = A && !A.configured() ? "Configura" : A === Local ? "Scegli la cartella" : "Collega";
  $("#btnLogin").classList.toggle("hidden", !A || A.connected());
  $("#btnLogout").classList.toggle("hidden", !A || !A.connected() || (A === Local && Local.zipMode));
  const tpl = await DB.getKV("areteTemplate");
  $("#areteTplInfo").textContent = tpl ? `Caricato: ${tpl.name} (${new Date(tpl.date).toLocaleDateString("it-IT")})` : "Nessun modello caricato";
  let perm = "?";
  try { if (navigator.permissions) perm = (await navigator.permissions.query({ name: "geolocation" })).state; } catch (e) { /* Safari vecchi */ }
  const sw = navigator.serviceWorker && navigator.serviceWorker.controller ? "attivo" : "no";
  checkAppFiles();
  $("#verInfo").innerHTML = `App ${APP_VERSION} · mappa ${window.L ? "ok" : "NON caricata"} · satellite ${Basemap.hasSat() ? "attivo" : "chiave mancante"} · posizione: ${perm} · ` +
    `ultimo GPS: ${GPS.last ? "±" + Math.round(GPS.last.acc) + " m" : "nessuno"} · offline: ${sw}`;
}
function setNet() {
  const on = navigator.onLine;
  $("#netState").textContent = on ? "online" : "offline";
  $("#netState").className = "badge " + (on ? "synced" : "pending");
}

/* --------------------------------------------------------------- Avvio */
function wizardCode() {
  const A = providerAuth();
  $("#wCodeBox").classList.toggle("hidden", !A || A.configured());
}
function openWizard() {
  $("#wName").value = S.name; $("#wInit").value = S.init;
  setSeg("wProv", S.provider); wizardCode();
  show("welcome");
}

async function main() {
  await DB.open();
  await loadSettings();
  S.north = "T";   // sempre nord geografico
  GAuth.load();
  await Local.load();
  const gBack = await GAuth.handleRedirect();
  if (gBack) S.provider = "google";
  drawRose();
  $("#camArrow").innerHTML = arrowSvg(0);
  $("#camArrow").classList.toggle("hidden", !S.arrow);
  Compass.onChange(updateCompassUI);
  setNet();
  window.addEventListener("online", setNet);
  window.addEventListener("offline", setNet);
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) GPS.stop();          // risparmio batteria quando l'app è in secondo piano
    else GPS.start();
  });
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});

  // specie (elenco del Protocollo ARETÉ)
  fetch("specie.json").then((r) => r.json()).then((list) => {
    $("#specieList").innerHTML = list.map((s) => `<option value="${esc(s)}">`).join("");
  }).catch(() => {});

  $$("nav.tabs button").forEach((b) => b.addEventListener("click", () => show(b.dataset.go)));
  $("#btnCompass").addEventListener("click", async () => {
    if (await Compass.start()) { $("#btnCompass").classList.add("hidden"); updateCompassUI(Compass); }
  });
  $("#btnNew").addEventListener("click", async () => { Compass.start().then(() => $("#btnCompass").classList.add("hidden")); openForm(); });
  // Annulla: serve un secondo tocco entro 4 secondi, così non si perde il rilievo per sbaglio
  let cancelArm = 0, cancelT = null;
  const disarmCancel = () => { cancelArm = 0; clearTimeout(cancelT); const b = $("#btnFormBack"); b.classList.remove("armed"); b.textContent = "‹ Annulla"; };
  $("#btnFormBack").addEventListener("click", () => {
    if (draft && Date.now() - cancelArm > 4000) {
      cancelArm = Date.now();
      const b = $("#btnFormBack");
      b.classList.add("armed"); b.textContent = "Tocca ancora per annullare";
      toast(editingId ? "Tocca ancora «Annulla» per uscire senza salvare le modifiche" : "Tocca ancora «Annulla» per eliminare questo rilievo", 3500);
      clearTimeout(cancelT); cancelT = setTimeout(disarmCancel, 4000);
      return;
    }
    disarmCancel();
    const back = editingId; editingId = null;
    draft = null; GPS.fixed = null;
    back ? openDetail(back) : show("home");
  });
  $("#fLavoroSel").addEventListener("change", (e) => {
    const v = e.target.value;
    if (v === "__new") { showNewLavoro(true, ""); $("#fLavoro").focus(); } else showNewLavoro(false, v);
    updateNamePreview(); checkDuplicate();
  });
  $("#btnClino").addEventListener("click", (e) => { e.preventDefault(); Clino.open((h) => { $("#fAlt").value = h; toast(`Altezza inserita: ${h} m`); }); });
  Clino.bind();
  document.querySelectorAll(".helpbtn[data-help]").forEach((b) => b.addEventListener("click", (e) => { e.preventDefault(); $("#" + b.dataset.help).classList.toggle("hidden"); }));
  bindSeg("segEtichette", (v) => { S.etichette = v === "1"; saveSettings(); });
  bindSeg("segArrowSet", (v) => { S.arrow = v === "1"; saveSettings(); $("#btnArrow").textContent = "Freccia N: " + (S.arrow ? "sì" : "no"); $("#btnArrow").classList.toggle("on", S.arrow); $("#camArrow").classList.toggle("hidden", !S.arrow); });
  $("#btnCopy").addEventListener("click", () => current && copyFrom(current.id));
  $("#btnPmMe").addEventListener("click", () => ProjMap.locate());  bindSeg("segClino", (v) => { S.clino = v === "1"; saveSettings(); });
  bindSeg("segColor", (v) => { S.mapColor = v; saveSettings(); ProjMap.renderControls(); ProjMap.plot(); });
  $("#pmFilter").addEventListener("change", (e) => { S.mapFilter = e.target.value; saveSettings(); ProjMap.plot(); });
  $("#btnAreteHelp").addEventListener("click", (e) => { e.preventDefault(); $("#areteHelp").classList.toggle("hidden"); });
  bindSeg("segGradi", (v) => { S.gradiFoto = v === "1"; saveSettings(); setSeg("segGradi", v); });
  Viewer.bind(); LblPick.bind();
  $("#intAdd").addEventListener("change", (e) => { addIntervento(e.target.value); e.target.value = ""; });
  $("#fCpc").addEventListener("change", () => applyRicontrollo(true));
  $("#btnCal").addEventListener("click", () => { const d = $("#fData"); try { d.showPicker(); } catch (e) { d.focus(); d.click(); } });
  ["fNgs", "fNcart", "fNprog"].forEach((id) => $("#" + id).addEventListener("input", updateIdReq));
  $("#btnEdit").addEventListener("click", () => current && editRecord(current.id));
  $("#btnNewPrj").addEventListener("click", () => { $("#newPrjBox").classList.remove("hidden"); $("#newPrjName").value = ""; $("#newPrjName").focus(); });
  $("#btnNewPrjNo").addEventListener("click", () => $("#newPrjBox").classList.add("hidden"));
  $("#btnNewPrjOk").addEventListener("click", async () => {
    const n = $("#newPrjName").value.trim().replace(/\s+/g, " ");
    if (!n) { toast("Scrivi il nome del progetto"); return; }
    const list = (await DB.getKV("projects")) || [];
    if (!list.includes(n)) { list.push(n); await DB.setKV("projects", list); }
    $("#newPrjBox").classList.add("hidden");
    openProject(n);
  });
  $("#fData").addEventListener("change", updateNamePreview);
  $("#btnPrjReload").addEventListener("click", () => {
    if (confirm("Ricaricare tutto il progetto sul cloud?\n\nSchede, foto, schede ARETÉ, tabella e vettoriale vengono rifatti e sovrascritti nel drive scelto in Impostazioni. Serve per esempio dopo un cambio di drive.")) Sync.exportProject(currentProject, S.out);
  });
  $("#btnSelList").addEventListener("click", () => { Sel.mode ? Sel.off() : Sel.on("list"); renderList(); });
  $("#btnSelPrj").addEventListener("click", () => { Sel.mode ? Sel.off() : Sel.on("prj"); openProject(currentProject); });
  $("#selCancel").addEventListener("click", () => { const m = Sel.mode; Sel.off(); Sel.refresh(m); });
  $("#selAll").addEventListener("click", () => { if (Sel.keys.size === Sel.all.length) Sel.keys.clear(); else Sel.all.forEach((k) => Sel.keys.add(k)); Sel.bar(); });
  $("#selSync").addEventListener("click", () => selAction("sync"));
  $("#selFree").addEventListener("click", () => selAction("free"));
  $("#selDel").addEventListener("click", () => selAction("del"));
  $("#sStudio").addEventListener("change", (e) => { S.studio = e.target.value.trim(); saveSettings(); });
  $("#sQual").addEventListener("change", (e) => { S.qualifica = e.target.value.trim(); saveSettings(); });
  bindSeg("segProtocollo", (v) => setProtocollo(v, true));

  $("#fLavoro").addEventListener("change", () => checkDuplicate());
  ["fLavoro", "fNgs", "fNcart", "fNprog"].forEach((id) => $("#" + id).addEventListener("input", updateNamePreview));
  ["fNgs", "fNcart", "fNprog"].forEach((id) => $("#" + id).addEventListener("change", checkDuplicate));
  $("#btnPrjBack").addEventListener("click", () => show("list"));
  $("#btnPmFit").addEventListener("click", () => ProjMap.fit());
  $("#btnListImport").addEventListener("click", () => Importer.open(""));
  $("#btnPrjTodoDel").addEventListener("click", async () => {
    const todo = (await DB.all()).filter((r) => r.lavoro === currentProject && isTodo(r));
    if (!todo.length || !confirm(`Eliminare le ${todo.length} piante ancora da fare di questo progetto? I rilievi già fatti restano.`)) return;
    for (const r of todo) await DB.del(r.id);
    toast("Piante da fare eliminate"); openProject(currentProject); MapView.plot();
  });
  $("#btnPmAll").addEventListener("click", () => ProjMap.all(true));
  $("#btnPmNone").addEventListener("click", () => ProjMap.all(false));
  $("#btnPrjNew").addEventListener("click", () => { Compass.start(); openForm(currentProject); });
  $("#btnSavePos").addEventListener("click", saveDetailPos);
  $("#btnGpsFix").addEventListener("click", () => {
    const f = GPS.best || GPS.last;
    if (!f) { toast("Nessuna posizione ancora"); return; }
    GPS.fixed = { ...f }; updateGpsUI(); toast(`Posizione fissata (±${Math.round(f.acc)} m)`);
  });

  // fotocamera interna
  $("#btnCamera").addEventListener("click", async () => {
    await Compass.start();
    show("camera");
    $("#btnArrow").textContent = "Freccia N: " + (S.arrow ? "sì" : "no");
    $("#btnArrow").classList.toggle("on", S.arrow);
    renderThumbs();
    if (!(await Cam.open())) show("form");
  });
  $("#btnArrow").addEventListener("click", () => {
    S.arrow = !S.arrow; saveSettings();
    $("#camArrow").classList.toggle("hidden", !S.arrow);
    $("#btnArrow").textContent = "Freccia N: " + (S.arrow ? "sì" : "no");
    $("#btnArrow").classList.toggle("on", S.arrow);
  });
  $("#btnShot").addEventListener("click", () => {
    const v = $("#video");
    if (!v.videoWidth) return;
    const meta = shotMeta();
    const fl = $("#flash"); fl.style.opacity = 0.8; setTimeout(() => (fl.style.opacity = 0), 120);
    const p = makePhotos(v, v.videoWidth, v.videoHeight, meta);
    draft.foto.push({ ...meta, clean: p.clean, arrow: p.arrow, thumb: p.thumb });
    renderThumbs();
  });
  $("#btnCamDone").addEventListener("click", () => { Cam.close(); show("form"); renderThumbs(); });

  // fotocamera di sistema (qualità piena): la direzione viene letta al momento del tocco
  let sysMeta = null;
  $("#btnSysCam").addEventListener("click", () => {
    if (!Compass.started) Compass.start();   // senza await: l'apertura della fotocamera deve restare legata al tocco
    sysMeta = shotMeta();
    $("#sysCamInput").click();
  });
  $("#sysCamInput").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
    const meta = sysMeta || shotMeta();
    const p = makePhotos(bmp, bmp.width, bmp.height, meta);
    draft.foto.push({ ...meta, clean: p.clean, arrow: p.arrow, thumb: p.thumb });
    e.target.value = "";
    renderThumbs();
  });
  $("#btnSave").addEventListener("click", saveForm);

  $("#btnSync").addEventListener("click", () => Sync.run(true));
  $("#btnDetBack").addEventListener("click", () => (current ? openProject(current.lavoro || "(senza progetto)") : show("list")));
  $("#btnShare").addEventListener("click", shareCurrent);
  $("#btnDelete").addEventListener("click", deleteCurrent);

  // impostazioni
  const bindS = (id, key, f = (x) => x) => $("#" + id).addEventListener("change", (e) => { S[key] = f(e.target.value); saveSettings(); });
  bindS("sFolder", "folder", (x) => x.trim() || "Rilievi");
  bindS("sName", "name", (x) => x.trim());
  bindS("sInit", "init", (x) => x.trim().toUpperCase());
  bindS("sDecl", "decl", (x) => Number(String(x).replace(",", ".")) || 0);
  bindSeg("segQuality", (v) => { S.quality = Number(v); saveSettings(); });
  bindSeg("segProv", (v) => { S.provider = v; saveSettings(); renderSettings(); });
  $("#btnLogin").addEventListener("click", () => {
    const A = providerAuth();
    if (!A) return;
    if (!A.configured()) { openWizard(); return; }
    A.login();
  });
  $("#btnLogout").addEventListener("click", async () => {
    const A = providerAuth();
    if (A === GAuth || A === Local) { await A.logout(); renderSettings(); } else if (A) await A.logout();
  });
  $("#sAreteTpl").addEventListener("change", async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      await ARETE.checkTemplate(f);
      await DB.setKV("areteTemplate", { blob: f, name: f.name, date: Date.now() });
      toast("Modello ARETÉ caricato");
    } catch (err) { toast(err.message, 6000); }
    e.target.value = "";
    renderSettings();
  });

  Install.render($("#wInstall"), { hideIfInstalled: true });
  Install.render($("#sInstall"), { hideIfInstalled: true });
  // configurazione iniziale
  bindSeg("wProv", (v) => { S.provider = v; saveSettings(); wizardCode(); });
  $("#wName").addEventListener("change", (e) => { S.name = e.target.value.trim(); saveSettings(); });
  $("#wInit").addEventListener("change", (e) => { S.init = e.target.value.trim().toUpperCase(); saveSettings(); });
  $("#wConnect").addEventListener("click", async () => {
    S.name = $("#wName").value.trim(); S.init = $("#wInit").value.trim().toUpperCase();
    if (!S.name || !S.init) { $("#wErr").textContent = "Scrivi nome e iniziali"; return; }
    if (!S.provider) { $("#wErr").textContent = "Scegli OneDrive o Google Drive"; return; }
    const code = $("#wCode").value.trim();
    if (code) { if (S.provider === "google") S.gClientId = code; else S.msClientId = code; }
    await saveSettings();
    const A = providerAuth();
    if (!A.configured()) { $("#wErr").textContent = "Inserisci il codice dell'app"; return; }
    if (A === Local) { await Local.login(); if (Local.connected()) { S.setupDone = true; await saveSettings(); show("home"); } return; }
    if (A === Auth && !Auth.msal) await Auth.init();
    A.login();
  });
  $("#wSkip").addEventListener("click", async () => {
    S.name = $("#wName").value.trim() || S.name; S.init = $("#wInit").value.trim().toUpperCase() || S.init;
    S.setupDone = true; await saveSettings(); show("home");
  });

  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
  await Auth.init().catch((e) => console.warn(e));
  try { MapView.init(); } catch (e) { showErr("mappa: " + e.message); }
  GPS.start();
  $("#btnGpsOn").addEventListener("click", () => {
    $("#mapInfo").textContent = "Richiesta della posizione…";
    navigator.geolocation.getCurrentPosition(() => { GPS.stop(); GPS.start(); }, gpsError, { enableHighAccuracy: true, timeout: 20000 });
  });
  ARETE.init().then(() => { ARETE.build($("#areteBox")); window.ARETE_READY = true; }).catch((e) => console.warn("ARETÉ", e));
  const A = providerAuth();
  if (A && A.connected() && !S.setupDone) {
    S.setupDone = true; toast(`${provider().name} collegato: ${A.label()}`, 4000);
  }
  await saveSettings();
  if (!S.setupDone) openWizard();
  refreshCounts();
  if (gBack) Sync.run(true);   // appena collegato (o rinnovato) l'accesso Google: invio quello che c'è
  checkAppFiles().then((m) => { if (m.length) toast("Mancano dei file dell'app: vedi Impostazioni", 5000); });
}
main();
