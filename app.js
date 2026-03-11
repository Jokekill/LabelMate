// ===== app.js – add-in bootstrap, language selection, UI text refresh =====
//
// Small improvements in this version:
// - Sets <html lang="..."> based on effective language (accessibility).
// - Keeps existing behavior (AUTO + manual language override, heartbeat banner).

(function () {
  function setDocumentLangAttr(effective) {
    // Map your internal codes to BCP-47.
    const map = { EN: "en", CZ: "cs", SK: "sk" };
    const lang = map[effective] || "en";
    document.documentElement.setAttribute("lang", lang);
  }

  function applyStaticTexts() {
    const L = window.LM.i18n.T();
    const byId = (id, val) => {
      const el = document.getElementById(id);
      if (el) el.textContent = val;
    };

    byId("appTitle", L.appTitle);
    byId("langLabel", L.langLabel);
    byId("choosePrompt", L.choosePrompt);
  }

  function initLanguageUI() {
    const sel = document.getElementById("langSelect");
    const status = document.getElementById("langStatus");
    if (!sel) return;

    sel.value = window.LM.i18n.getSavedLangOverride();
    const effective = (sel.value === "AUTO")
      ? window.LM.i18n.langFromOfficeContext()
      : sel.value;

    setDocumentLangAttr(effective);

    if (status) {
      status.textContent = (sel.value === "AUTO")
        ? `Auto: ${effective}`
        : `Manual: ${effective}`;
    }

    applyStaticTexts();
    window.Labels.renderButtons();
    window.Banner.refresh().catch(() => {});

    sel.addEventListener("change", () => {
      const val = sel.value;
      window.LM.i18n.saveLangOverride(val);

      const lang = (val === "AUTO")
        ? window.LM.i18n.langFromOfficeContext()
        : val;

      setDocumentLangAttr(lang);

      if (status) {
        status.textContent = (val === "AUTO")
          ? `Auto: ${lang}`
          : `Manual: ${lang}`;
      }

      applyStaticTexts();
      window.Labels.renderButtons();
      window.Banner.refresh().catch(() => {});
    });
  }

  function boot() {
    initLanguageUI();
    window.Banner.startHeartbeat();
  }

  // Run in Office and also outside Office (browser debugging)
  // Office.onReady returns a Promise / can be called in multiple spots. citeturn8search4
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => boot());
  } else if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
