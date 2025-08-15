// ===== core.js – společné utility, i18n a UI helpers =====

window.LM = window.LM || {};

(function () {
  const LS_LANG = "labelmate_lang_override"; // 'AUTO' | 'EN' | 'CZ' | 'SK'

  function getSavedLangOverride() {
    try { return localStorage.getItem(LS_LANG) || "AUTO"; } catch { return "AUTO"; }
  }
  function saveLangOverride(v) {
    try { localStorage.setItem(LS_LANG, v); } catch {}
  }
  function langFromOfficeContext() {
    const ctx = (typeof Office !== "undefined" && Office.context) || {};
    const content = (ctx && ctx.contentLanguage) || "";
    const display = (ctx && ctx.displayLanguage) || "";
    const combined = (content || display).toLowerCase();
    if (combined.includes("cs")) return "CZ";
    if (combined.includes("sk")) return "SK";
    if (combined.includes("en")) return "EN";
    return window.LM_DEFAULT_LANG || "EN";
  }
  function getCurrentLangCode() {
    const sel = document.getElementById("langSelect");
    const v = sel?.value || "AUTO";
    return (v === "AUTO") ? langFromOfficeContext() : v;
  }
  function T() {
    const code = getCurrentLangCode();
    return window.LM_I18N[code] || window.LM_I18N[window.LM_DEFAULT_LANG];
  }

  function setStatusOk(msg) {
    const el = document.getElementById("status");
    if (!el) return;
    el.classList.remove("text-rose-600","dark:text-rose-400");
    el.classList.add("text-emerald-600","dark:text-emerald-400");
    el.textContent = msg || "";
  }
  function setStatusError(msg) {
    const el = document.getElementById("status");
    if (!el) return;
    el.classList.remove("text-emerald-600","dark:text-emerald-400");
    el.classList.add("text-rose-600","dark:text-rose-400");
    el.textContent = msg || "";
  }
  function clearStatus() {
    const el = document.getElementById("status");
    if (el) el.textContent = "";
  }
  function setBusy(isBusy) {
    document.querySelectorAll("#labelsContainer button").forEach(b => b.disabled = isBusy);
  }

  window.LM.i18n = {
    getSavedLangOverride, saveLangOverride, langFromOfficeContext, getCurrentLangCode, T
  };
  window.LM.ui = { setStatusOk, setStatusError, clearStatus, setBusy };
})();
