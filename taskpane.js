// taskpane.js — orchestrator
(function () {
  "use strict";

  let started = false;

  async function init() {
    if (started) return;
    started = true;

    try {
      await window.LM?.classification?.ensureOfficeReady?.();
    } catch (error) {
      console.error("Office bootstrap failed:", error);
      return;
    }

    try { window.Theme?.init?.(); } catch (e) { console.warn("Theme.init:", e); }
    try { window.LM?.app?.init?.(); } catch (e) { console.warn("LM.app.init:", e); }
    try { window.Labels?.renderButtons?.(); } catch (e) { console.warn("Labels.renderButtons:", e); }
    try { await window.Banner?.init?.(); } catch (e) { console.warn("Banner.init:", e); }
    try { window.Banner?.startHeartbeat?.(); } catch (e) { console.warn("Banner.startHeartbeat:", e); }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      init().catch((e) => console.error("Taskpane init failed:", e));
    }, { once: true });
  } else {
    init().catch((e) => console.error("Taskpane init failed:", e));
  }
})();
