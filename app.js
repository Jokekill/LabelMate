// ===== app.js – bootstrap doplňku, jazyk, texty (bez Theme.initUI) =====

(function () {
  function applyStaticTexts() {
    const L = window.LM.i18n.T();
    const byId = (id, val) => {
      const el = document.getElementById(id);
      if (el) el.textContent = val;
    };
    byId("appTitle", L.appTitle);
    byId("themeLabel", L.themeLabel);
    byId("langLabel", L.langLabel);
    byId("choosePrompt", L.choosePrompt);

    // Lokalizace textu voleb motivu (jen light/dark)
    const themeSel = document.getElementById("themeSelect");
    if (themeSel) {
      const optL = themeSel.querySelector('option[value="light"]');
      const optD = themeSel.querySelector('option[value="dark"]');
      if (optL) optL.textContent = L.themeOptions.light;
      if (optD) optD.textContent = L.themeOptions.dark;
    }
  }

  function initLanguageUI() {
    const sel = document.getElementById("langSelect");
    const status = document.getElementById("langStatus");
    if (!sel) return;

    sel.value = window.LM.i18n.getSavedLangOverride();
    const effective = (sel.value === "AUTO")
      ? window.LM.i18n.langFromOfficeContext()
      : sel.value;

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

  // Spuštění v Office i mimo Office (pro ladění v prohlížeči)
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => boot());
  } else if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
