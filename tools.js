/* Rilievi di Campo – strumenti in campo: misura dell'altezza con il telefono (inclinometro).
   Metodo delle due letture: angolo verso la cima e verso la base, più la distanza orizzontale dal tronco.
   H = d × (tan α_cima − tan α_base). Non serve l'altezza degli occhi. */
"use strict";
const Clino = (() => {
  let stream = null, beta = null, gamma = null, sm = null, top = null, base = null, onUse = null, listening = false;
  const $c = (id) => document.getElementById(id);
  const deg = (r) => (r * 180) / Math.PI, rad = (d) => (d * Math.PI) / 180;
  function onOri(e) {
    if (typeof e.beta !== "number") return;
    beta = e.beta; gamma = e.gamma || 0;
    // angolo dell'asse della fotocamera sopra l'orizzonte (telefono in verticale): beta − 90
    const a = beta - 90;
    sm = sm === null ? a : sm + (a - sm) * 0.25;   // smorzo il tremolio della mano
    draw();
  }
  function draw() {
    if (!$c("clino") || $c("clino").classList.contains("hidden")) return;
    const ok = sm !== null;
    $c("clAngle").textContent = ok ? `${sm >= 0 ? "+" : ""}${sm.toFixed(1)}°` : "—";
    $c("clWarn").textContent = !ok ? "In attesa del sensore di inclinazione…"
      : Math.abs(gamma) > 25 ? "Tieni il telefono dritto (non inclinato di lato)" : "";
    $c("clTop").textContent = top === null ? "Segna cima ▲" : `Cima: ${top.toFixed(1)}°`;
    $c("clBase").textContent = base === null ? "Segna base ▼" : `Base: ${base.toFixed(1)}°`;
    const d = parseFloat(String($c("clDist").value).replace(",", "."));
    let h = null;
    if (top !== null && base !== null && d > 0 && top > base && Math.abs(top) < 85 && Math.abs(base) < 85)
      h = d * (Math.tan(rad(top)) - Math.tan(rad(base)));
    $c("clRes").textContent = h ? `Altezza ≈ ${h.toFixed(1)} m` : top !== null && base !== null && !(d > 0) ? "Inserisci la distanza dal tronco" : "";
    $c("clUse").disabled = !h;
    $c("clUse").dataset.h = h ? h.toFixed(1) : "";
  }
  async function open(useFn) {
    onUse = useFn; top = null; base = null; sm = null;
    if (typeof Compass !== "undefined") await Compass.start();   // su iPhone chiede il permesso ai sensori
    if (!listening) { window.addEventListener("deviceorientation", onOri); listening = true; }
    $c("clino").classList.remove("hidden"); document.body.classList.add("viewing");
    try {
      stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: "environment" } }, audio: false });
      $c("clVideo").srcObject = stream; await $c("clVideo").play().catch(() => {});
    } catch (e) { $c("clWarn").textContent = "Fotocamera non disponibile: mira lungo il bordo del telefono"; }
    draw();
  }
  function close() {
    $c("clino").classList.add("hidden"); document.body.classList.remove("viewing");
    if (stream) stream.getTracks().forEach((t) => t.stop());
    stream = null; $c("clVideo").srcObject = null;
  }
  function bind() {
    $c("clClose").addEventListener("click", close);
    $c("clTop").addEventListener("click", () => { if (sm !== null) { top = sm; draw(); } });
    $c("clBase").addEventListener("click", () => { if (sm !== null) { base = sm; draw(); } });
    $c("clDist").addEventListener("input", draw);
    $c("clReset").addEventListener("click", () => { top = null; base = null; draw(); });
    $c("clUse").addEventListener("click", () => { const h = $c("clUse").dataset.h; if (h && onUse) onUse(h); close(); });
  }
  return { open, close, bind, _calc: (d, t, b) => d * (Math.tan(rad(t)) - Math.tan(rad(b))) };
})();
