// banner.js
// Sjednocení starého a nového banner API: používá existující DOM v taskpane.html,
// lokalizuje texty z i18n a vystaví window.Banner i window.LMBanner.
(function () {
  "use strict";

  let heartbeatId = null;
  let refreshInFlight = null;

  function getElements() {
    return {
      banner: document.getElementById("missingBanner") || document.getElementById("classification-banner"),
      title: document.getElementById("bnrTitle"),
      desc: document.getElementById("bnrDesc")
    };
  }

  function getTexts() {
    const L = window.LM?.i18n?.T?.() || {};
    return L.banner || {
      title: "No classification set",
      desc: "Choose a classification below. The banner will disappear after it is set."
    };
  }

  function show() {
    const { banner, title, desc } = getElements();
    if (!banner) return;

    const text = getTexts();
    if (title) title.textContent = text.title || "";
    if (desc) desc.textContent = text.desc || "";

    banner.classList.remove("hidden");
    banner.hidden = false;
  }

  function hide() {
    const { banner } = getElements();
    if (!banner) return;

    banner.classList.add("hidden");
    banner.hidden = true;
  }

  async function hasClassificationSafe() {
    const checker = window.LM?.classification?.hasClassification;
    if (typeof checker !== "function") return false;

    try {
      return await checker();
    } catch (error) {
      console.error("LabelMate banner classification check failed:", error);
      return false;
    }
  }

  async function refresh() {
    if (refreshInFlight) return refreshInFlight;

    refreshInFlight = (async () => {
      const exists = await hasClassificationSafe();
      if (exists) hide();
      else show();
    })();

    try {
      await refreshInFlight;
    } finally {
      refreshInFlight = null;
    }
  }

  function startHeartbeat(intervalMs) {
    if (heartbeatId) return;
    heartbeatId = window.setInterval(() => {
      refresh().catch(() => {});
    }, intervalMs || 2500);
  }

  function stopHeartbeat() {
    if (!heartbeatId) return;
    window.clearInterval(heartbeatId);
    heartbeatId = null;
  }

  function init() {
    refresh().catch(() => {});
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init, { once: true });
  } else {
    init();
  }

  if (window.Office?.onReady) {
    Office.onReady(() => {
      init();
      window.setTimeout(() => refresh().catch(() => {}), 300);
      window.setTimeout(() => refresh().catch(() => {}), 1000);
    });
  }

  window.addEventListener("labelmate:classification-changed", () => refresh().catch(() => {}));
  window.addEventListener("labelmate:rerender-labels", () => refresh().catch(() => {}));
  window.addEventListener("focus", () => refresh().catch(() => {}));
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) refresh().catch(() => {});
  });

  const api = { init, refresh, show, hide, startHeartbeat, stopHeartbeat };
  window.Banner = api;
  window.LMBanner = api;
  window.refreshBanner = refresh;
})();