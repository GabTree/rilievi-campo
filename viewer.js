/* Rilievi di Campo – galleria a schermo intero: due dita per ingrandire, doppio tocco, scorrimento tra le foto.
   Viewer.open(foto, indice): foto = elenco di foto del rilievo ({clean, arrow, ...}). */
"use strict";
const Viewer = (() => {
  let list = [], i = 0, withArrow = false, url = null;
  let sc = 1, tx = 0, ty = 0;                    // ingrandimento e spostamento
  const pts = new Map(); let pinch0 = null, pan0 = null, lastTap = 0, swipe0 = null;
  const el = () => document.getElementById("viewer");
  const img = () => document.getElementById("vImg");
  function apply(anim) {
    const im = img();
    im.style.transition = anim ? "transform .2s" : "none";
    im.style.transform = `translate(${tx}px, ${ty}px) scale(${sc})`;
  }
  function reset(anim) { sc = 1; tx = 0; ty = 0; apply(anim); }
  function blobOf(f) { return withArrow ? f.arrow || f.clean : f.clean || f.arrow; }
  function draw() {
    const f = list[i];
    if (url) URL.revokeObjectURL(url);
    url = URL.createObjectURL(blobOf(f));
    img().src = url;
    document.getElementById("vCount").textContent = `${i + 1} / ${list.length}${f.etichetta ? " · " + (f.etichetta === "Difetto" ? "Anomalia" : f.etichetta) : ""}`;
    const both = !!(f.arrow && f.clean);
    const b = document.getElementById("vArrow");
    b.classList.toggle("hidden", !both);
    b.textContent = withArrow ? "Senza freccia" : "Con freccia N";
    document.getElementById("vNote").textContent = f.ridotta ? "Copia leggera: l'originale è nel cloud" : "";
    reset(false);
  }
  function go(d) { if (list.length < 2) return; i = (i + d + list.length) % list.length; draw(); }
  function open(foto, idx = 0) {
    list = foto.filter((f) => f.clean || f.arrow); if (!list.length) return;
    i = Math.min(idx, list.length - 1); withArrow = false;
    el().classList.remove("hidden"); document.body.classList.add("viewing");
    draw();
  }
  function close() { el().classList.add("hidden"); document.body.classList.remove("viewing"); if (url) { URL.revokeObjectURL(url); url = null; } }
  function clamp() {
    const r = el().getBoundingClientRect(), m = img();
    const w = m.clientWidth * sc, h = m.clientHeight * sc;
    const mx = Math.max(0, (w - r.width) / 2), my = Math.max(0, (h - r.height) / 2);
    tx = Math.max(-mx, Math.min(mx, tx)); ty = Math.max(-my, Math.min(my, ty));
  }
  function bind() {
    const stage = document.getElementById("vStage");
    document.getElementById("vBack").addEventListener("click", close);
    document.getElementById("vPrev").addEventListener("click", () => go(-1));
    document.getElementById("vNext").addEventListener("click", () => go(1));
    document.getElementById("vArrow").addEventListener("click", () => { withArrow = !withArrow; draw(); });
    stage.addEventListener("pointerdown", (e) => {
      stage.setPointerCapture(e.pointerId);
      pts.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pts.size === 2) {
        const [a, b] = [...pts.values()];
        pinch0 = { d: Math.hypot(a.x - b.x, a.y - b.y), sc, tx, ty, cx: (a.x + b.x) / 2, cy: (a.y + b.y) / 2 };
        swipe0 = null;
      } else if (pts.size === 1) {
        pan0 = { x: e.clientX, y: e.clientY, tx, ty }; swipe0 = { x: e.clientX, y: e.clientY, t: Date.now() };
      }
    });
    stage.addEventListener("pointermove", (e) => {
      if (!pts.has(e.pointerId)) return;
      pts.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pts.size === 2 && pinch0) {
        const [a, b] = [...pts.values()];
        const d = Math.hypot(a.x - b.x, a.y - b.y);
        sc = Math.max(1, Math.min(6, pinch0.sc * d / pinch0.d));
        const cx = (a.x + b.x) / 2, cy = (a.y + b.y) / 2;
        tx = pinch0.tx + (cx - pinch0.cx); ty = pinch0.ty + (cy - pinch0.cy);
        clamp(); apply(false);
      } else if (pts.size === 1 && pan0 && sc > 1) {
        tx = pan0.tx + (e.clientX - pan0.x); ty = pan0.ty + (e.clientY - pan0.y);
        clamp(); apply(false);
      }
    });
    const up = (e) => {
      pts.delete(e.pointerId);
      if (pts.size < 2) pinch0 = null;
      if (pts.size === 0) {
        if (sc <= 1.02) reset(true);
        // scorrimento orizzontale (solo a foto non ingrandita): foto precedente / successiva
        if (swipe0 && sc <= 1.02) {
          const dx = e.clientX - swipe0.x, dy = e.clientY - swipe0.y;
          if (Math.abs(dx) > 60 && Math.abs(dx) > Math.abs(dy) * 1.5 && Date.now() - swipe0.t < 700) go(dx < 0 ? 1 : -1);
          else if (Math.abs(dx) < 10 && Math.abs(dy) < 10) {
            // doppio tocco: ingrandisce dove tocchi, poi torna indietro
            const now = Date.now();
            if (now - lastTap < 300) {
              const r = stage.getBoundingClientRect();
              sc = 2.5; tx = (r.width / 2 - (e.clientX - r.left)) * 1.5; ty = (r.height / 2 - (e.clientY - r.top)) * 1.5; clamp(); apply(true);
              lastTap = 0;
            } else lastTap = now;
          }
        } else if (swipe0 && sc > 1.02) {
          const dx = e.clientX - swipe0.x, dy = e.clientY - swipe0.y;
          if (Math.abs(dx) < 10 && Math.abs(dy) < 10) {
            const now = Date.now();
            if (now - lastTap < 300) { reset(true); lastTap = 0; } else lastTap = now;
          }
        }
        swipe0 = null; pan0 = null;
      }
    };
    stage.addEventListener("pointerup", up);
    stage.addEventListener("pointercancel", up);
    document.addEventListener("keydown", (e) => {
      if (el().classList.contains("hidden")) return;
      if (e.key === "Escape") close(); else if (e.key === "ArrowRight") go(1); else if (e.key === "ArrowLeft") go(-1);
    });
  }
  return { open, close, bind };
})();
