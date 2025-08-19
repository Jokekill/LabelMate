// taskpane.js — lean orchestrator (bez Tailwindu)
(function () {
  function init() {
    // 1) Téma (světlý/tmavý) – bezpečně, pokud existuje
    try { window.Theme?.init?.(); } catch (e) { console.warn("Theme.init:", e); }

    // 2) Překlady / statické texty – pokud máš inicializátor v app.js nebo i18n.js
    try { window.LM?.app?.init?.(); } catch (e) { /* volitelné, nemusí existovat */ }

    // 3) Vyrenderuj klasifikační tlačítka
    try { window.renderButtons?.(); } catch (e) { console.warn("renderButtons:", e); }

    // 4) Banner „chybí klasifikace“ + heartbeat
    try { window.updateMissingBanner?.(); } catch (e) { console.warn("updateMissingBanner:", e); }
    try { window.startBannerHeartbeat?.(); } catch (e) { console.warn("startBannerHeartbeat:", e); }
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => init());
  } else if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
