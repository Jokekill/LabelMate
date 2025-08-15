// ===== app.js – bootstrap doplňku, jazyk, texty, theme init =====

(function () {
  function applyStaticTexts() {
    const L = window.LM.i18n.T();
    const byId = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    byId("appTitle", L.appTitle);
    byId("themeLabel", L.themeLabel);
    byId("langLabel", L.langLabel);
    byId("choosePrompt", L.choosePrompt);

    // lokalizace textu voleb motivu (hodnoty zůstávají auto/light/dark)
    const themeSel = document.getElementById("themeSelect");
    if (themeSel) {
      const optA = themeSel.querySelector('option[value="auto"]');
      const optL = themeSel.querySelector('option[value="light"]');
      const optD = themeSel.querySelector('option[value="dark"]');
      if (optA) optA.textContent = L.themeOptions.auto;
      if (optL) optL.textContent = L.themeOptions.light;
      if (optD) optD.textContent = L.themeOptions.dark;
    }
  }

  function initLanguageUI() {
    const sel = document.getElementById("langSelect");
    const status = document.getElementById("langStatus");
    if (!sel) return;

    sel.value = window.LM.i18n.getSavedLangOverride();
    const effective = (sel.value === "AUTO") ? window.LM.i18n.langFromOfficeContext() : sel.value;
    if (status) status.textContent = (sel.value === "AUTO") ? `Auto: ${effective}` : `Manual: ${effective}`;

    applyStaticTexts();
    window.Labels.renderButtons();
    window.Banner.refresh().catch(()=>{});

    sel.addEventListener("change", () => {
      const val = sel.value;
      window.LM.i18n.saveLangOverride(val);
      const lang = (val === "AUTO") ? window.LM.i18n.langFromOfficeContext() : val;
      if (status) status.textContent = (val === "AUTO") ? `Auto: ${lang}` : `Manual: ${lang}`;
      applyStaticTexts();
      window.Labels.renderButtons();
      window.Banner.refresh().catch(()=>{});
    });
  }

  Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      // Theme UI
      window.Theme.initUI();
      // Language, labels, banner
      initLanguageUI();
      window.Banner.startHeartbeat();
    }
  });
})();
