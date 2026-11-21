/* Rilievi di Campo – istruzioni per installare l'app sulla schermata Home, adattate a telefono e browser.
   Usato dall'app (benvenuto e Impostazioni) e dalla pagina installa.html da mandare ai colleghi. */
"use strict";
const Install = (() => {
  let deferred = null; // evento "beforeinstallprompt" di Chrome/Edge su Android e PC: permette il pulsante "Installa"
  const listeners = [];
  window.addEventListener("beforeinstallprompt", (e) => { e.preventDefault(); deferred = e; listeners.forEach((f) => f()); });
  window.addEventListener("appinstalled", () => { deferred = null; listeners.forEach((f) => f()); });

  function env() {
    const ua = navigator.userAgent || "";
    const ios = /iPhone|iPad|iPod/i.test(ua) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
    const android = /Android/i.test(ua);
    const standalone = (window.matchMedia && matchMedia("(display-mode: standalone)").matches) || navigator.standalone === true;
    // browser interni di altre app (WhatsApp, Facebook, Instagram, Telegram, Gmail…): da lì non si può installare
    const inApp = /FBAN|FBAV|FB_IAB|Instagram|WhatsApp|Line\/|Telegram|GSA\/|; wv\)/i.test(ua);
    let browser = "altro";
    if (ios) browser = /CriOS/i.test(ua) ? "chrome" : /FxiOS/i.test(ua) ? "firefox" : /EdgiOS/i.test(ua) ? "edge" : "safari";
    else if (/SamsungBrowser/i.test(ua)) browser = "samsung";
    else if (/Edg\//i.test(ua)) browser = "edge";
    else if (/Firefox/i.test(ua)) browser = "firefox";
    else if (/Chrome/i.test(ua)) browser = "chrome";
    else if (/Safari/i.test(ua)) browser = "safari";
    return { ios, android, standalone, inApp, browser, desktop: !ios && !android };
  }
  const SHARE = '<span class="ins-ic" aria-hidden="true">⬆︎</span>';
  function steps(e) {
    if (e.inApp) return {
      title: "Apri il link nel browser",
      list: [e.ios ? "Hai aperto il link dentro un'altra app (es. WhatsApp): da qui non si può installare."
                   : "Hai aperto il link dentro un'altra app (es. WhatsApp): da qui non si può installare.",
        e.ios ? "Tocca il menu (⋯ oppure l'icona della bussola) e scegli «Apri in Safari»." : "Tocca il menu ⋮ e scegli «Apri in Chrome» (o «Apri nel browser»).",
        "Poi segui le istruzioni che compaiono lì."],
    };
    if (e.ios) {
      if (e.browser === "safari") return { title: "Installa su iPhone (Safari)", list: [
        `Tocca il pulsante <b>Condividi</b> ${SHARE} (in basso, o in alto a destra su iPad).`,
        "Scorri e tocca <b>«Aggiungi alla schermata Home»</b>.",
        "Tocca <b>Aggiungi</b>: comparirà l'icona di Rilievi di Campo.",
        "Da ora apri sempre l'app dall'icona."] };
      if (e.browser === "chrome") return { title: "Installa su iPhone (Chrome)", list: [
        `Tocca <b>Condividi</b> ${SHARE} nella barra dell'indirizzo (o menu ⋯ → Condividi).`,
        "Tocca <b>«Aggiungi alla schermata Home»</b> → <b>Aggiungi</b>.",
        "Se la voce non c'è, apri lo stesso indirizzo in Safari e installa da lì.",
        "Da ora apri sempre l'app dall'icona."] };
      return { title: "Installa su iPhone", list: [
        `Tocca <b>Condividi</b> ${SHARE} → <b>«Aggiungi alla schermata Home»</b>.`,
        "Se la voce non c'è, apri lo stesso indirizzo in <b>Safari</b> e installa da lì.",
        "Da ora apri sempre l'app dall'icona."] };
    }
    if (e.android) {
      if (e.browser === "samsung") return { title: "Installa su Android (Samsung Internet)", list: [
        "Tocca il menu ≡ in basso → <b>«Aggiungi pagina a»</b> → <b>«Schermata Home»</b>.",
        "Da ora apri sempre l'app dall'icona."] };
      if (e.browser === "firefox") return { title: "Installa su Android (Firefox)", list: [
        "Tocca il menu ⋮ → <b>«Installa»</b> (o «Aggiungi alla schermata Home»).",
        "Se non c'è, apri lo stesso indirizzo in <b>Chrome</b>.", "Da ora apri sempre l'app dall'icona."] };
      return { title: "Installa su Android", list: [
        "Tocca il pulsante <b>Installa l'app</b> qui sotto, se c'è.",
        "Altrimenti: menu ⋮ in alto a destra → <b>«Installa app»</b> (o «Aggiungi a schermata Home»).",
        "Da ora apri sempre l'app dall'icona."] };
    }
    if (e.browser === "safari") return { title: "Installa sul Mac (Safari)", list: [
      "Menu <b>File → «Aggiungi al Dock»</b> (macOS Sonoma o successivo)."] };
    return { title: "Installa sul computer", list: [
      "Clicca il pulsante <b>Installa l'app</b> qui sotto, se c'è.",
      "Altrimenti: in Chrome o Edge clicca l'icona di installazione nella barra dell'indirizzo (a destra) o menu ⋮ → «Installa Rilievi di Campo»."] };
  }
  // disegna le istruzioni dentro el; opts.compact = solo stato se già installata
  function render(el, opts = {}) {
    const draw = () => {
      const e = env();
      if (e.standalone) { el.innerHTML = opts.hideIfInstalled ? "" : `<div class="ins-ok">✓ App installata: la stai usando dall'icona.</div>`; el.classList.toggle("hidden", !!opts.hideIfInstalled); return; }
      el.classList.remove("hidden");
      const s = steps(e);
      el.innerHTML = `<b class="ins-title">${s.title}</b><ol class="ins-list">${s.list.map((x) => `<li>${x}</li>`).join("")}</ol>
        ${deferred && !e.inApp ? `<button type="button" class="big ins-btn">Installa l'app</button>` : ""}
        <div class="ins-note">Usa sempre l'icona: i rilievi vengono salvati dentro l'app installata, separati dal browser.</div>`;
      const b = el.querySelector(".ins-btn");
      if (b) b.addEventListener("click", async () => { const d = deferred; deferred = null; d.prompt(); try { await d.userChoice; } catch (er) { /* annullato */ } draw(); });
    };
    listeners.push(draw);
    draw();
  }
  return { env, render };
})();
