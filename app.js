(function () {
  "use strict";

  let started = false;

  function getBannerApi() {
    return window.Banner || window.LMBanner || null;
  }

  function setDocumentLangAttr(effective) {
    const map = { EN: "en", CZ: "cs", SK: "sk" };
    document.documentElement.setAttribute("lang", map[effective] || "en");
  }

  function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  }

  function applyStaticTexts() {
    const L = window.LM?.i18n?.T?.();
    if (!L) return;

    setText("appTitle", L.appTitle || "LabelMate");
    setText("langLabel", L.langLabel || "Language");
    setText("choosePrompt", L.choosePrompt || "Choose classification level:");
    setText("bnrTitle", L.banner?.title || "");
    setText("bnrDesc", L.banner?.desc || "");
  }

  function updateLangStatus(selectedValue) {
    const status = document.getElementById("langStatus");
    if (!status || !window.LM?.i18n) return;

    const L = window.LM.i18n.T?.() || {};
    const effective = selectedValue === "AUTO"
      ? window.LM.i18n.langFromOfficeContext()
      : selectedValue;

    status.textContent =
      selectedValue === "AUTO"
        ? (typeof L.langStatusAuto === "function"
            ? L.langStatusAuto(effective)
            : `Auto: ${effective}`)
        : (typeof L.langStatusManual === "function"
            ? L.langStatusManual(effective)
            : `Manual: ${effective}`);
  }

  function rerenderUi() {
    applyStaticTexts();
    window.Labels?.renderButtons?.();

    try {
      window.dispatchEvent(new CustomEvent("labelmate:rerender-labels"));
    } catch (_) {}

    const banner = getBannerApi();
    if (banner?.refresh) {
      banner.refresh().catch(() => {});
    }
  }

  function initLanguageUI() {
    const sel = document.getElementById("langSelect");
    if (!sel || !window.LM?.i18n) return;

    const saved = window.LM.i18n.getSavedLangOverride();
    sel.value = saved;

    const effective = saved === "AUTO"
      ? window.LM.i18n.langFromOfficeContext()
      : saved;

    setDocumentLangAttr(effective);
    updateLangStatus(saved);
    rerenderUi();

    sel.addEventListener("change", () => {
      const val = sel.value;
      window.LM.i18n.saveLangOverride(val);

      const lang = val === "AUTO"
        ? window.LM.i18n.langFromOfficeContext()
        : val;

      setDocumentLangAttr(lang);
      updateLangStatus(val);
      rerenderUi();
    });
  }

  async function boot() {
    if (started) return;
    started = true;

    try {
      window.Theme?.init?.();
    } catch (e) {
      console.warn("Theme.init failed:", e);
    }

    try {
      await window.LM?.classification?.ensureOfficeReady?.();
    } catch (e) {
      console.error("Office bootstrap failed:", e);
      return;
    }

    try {
      initLanguageUI();
    } catch (e) {
      console.error("Language UI init failed:", e);
    }

    try {
      await getBannerApi()?.init?.();
    } catch (e) {
      console.error("Banner init failed:", e);
    }

    try {
      getBannerApi()?.startHeartbeat?.();
    } catch (e) {
      console.warn("Banner heartbeat failed:", e);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      boot().catch((e) => console.error("App boot failed:", e));
    }, { once: true });
  } else {
    boot().catch((e) => console.error("App boot failed:", e));
  }
})();