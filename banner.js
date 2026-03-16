// banner.js
(function () {
  "use strict";

  window.LM = window.LM || {};

  let initialized = false;
  let refreshInFlight = null;
  let lastKnownState = null;

  function getHostName() {
    try {
      return window.LM?.classification?.getHost?.() || Office?.context?.diagnostics?.host || null;
    } catch (_) {
      return null;
    }
  }

  function getDefaultBannerMessage() {
    const host = getHostName();

    if (host === Office?.HostType?.Word) {
      return "This Word document is not classified yet.";
    }

    if (host === Office?.HostType?.Excel) {
      return "This Excel workbook is not classified yet.";
    }

    if (host === Office?.HostType?.PowerPoint) {
      return "This PowerPoint presentation is not classified yet.";
    }

    return "This file is not classified yet.";
  }

  function resolveBannerElement() {
    return (
      document.getElementById("classification-banner") ||
      document.getElementById("banner") ||
      document.getElementById("label-banner") ||
      document.getElementById("missing-classification-banner") ||
      document.getElementById("lm-banner") ||
      document.querySelector("[data-role='classification-banner']")
    );
  }

  function resolveBannerTextElement(banner) {
    if (!banner) {
      return null;
    }

    return (
      banner.querySelector("[data-role='banner-text']") ||
      banner.querySelector(".banner-text") ||
      banner.querySelector(".classification-banner__text") ||
      banner.querySelector(".labelmate-banner__text") ||
      banner.querySelector("span") ||
      banner.querySelector("div")
    );
  }

  function resolveRootElement() {
    return (
      document.getElementById("app") ||
      document.querySelector("main") ||
      document.body
    );
  }

  function applyBannerStyles(banner) {
    banner.style.display = "none";
    banner.style.boxSizing = "border-box";
    banner.style.width = "100%";
    banner.style.marginBottom = "12px";
    banner.style.padding = "12px 14px";
    banner.style.borderRadius = "8px";
    banner.style.background = "#fef3c7";
    banner.style.color = "#92400e";
    banner.style.border = "1px solid #f59e0b";
    banner.style.fontSize = "14px";
    banner.style.fontWeight = "600";
    banner.style.lineHeight = "1.4";
  }

  function ensureBannerElement() {
    let banner = resolveBannerElement();
    if (banner) {
      if (!banner.id) {
        banner.id = "classification-banner";
      }
      banner.setAttribute("role", "status");
      banner.setAttribute("aria-live", "polite");
      banner.dataset.role = "classification-banner";
      applyBannerStyles(banner);

      if (!resolveBannerTextElement(banner)) {
        const text = document.createElement("div");
        text.className = "classification-banner__text";
        text.setAttribute("data-role", "banner-text");
        text.textContent = getDefaultBannerMessage();
        banner.innerHTML = "";
        banner.appendChild(text);
      }

      return banner;
    }

    const root = resolveRootElement();

    banner = document.createElement("div");
    banner.id = "classification-banner";
    banner.dataset.role = "classification-banner";
    banner.dataset.state = "warning";
    banner.setAttribute("role", "status");
    banner.setAttribute("aria-live", "polite");

    applyBannerStyles(banner);

    const text = document.createElement("div");
    text.className = "classification-banner__text";
    text.setAttribute("data-role", "banner-text");
    text.textContent = getDefaultBannerMessage();

    banner.appendChild(text);

    if (root.firstChild) {
      root.insertBefore(banner, root.firstChild);
    } else {
      root.appendChild(banner);
    }

    return banner;
  }

  function show(message) {
    const banner = ensureBannerElement();
    const textEl = resolveBannerTextElement(banner);

    if (textEl) {
      textEl.textContent = message || getDefaultBannerMessage();
    } else {
      banner.textContent = message || getDefaultBannerMessage();
    }

    banner.dataset.state = "warning";
    banner.hidden = false;
    banner.style.display = "block";
    lastKnownState = "shown";
  }

  function hide() {
    const banner = ensureBannerElement();
    banner.hidden = true;
    banner.style.display = "none";
    lastKnownState = "hidden";
  }

  async function hasClassificationSafe() {
    const checker = window.LM?.classification?.hasClassification;

    if (typeof checker !== "function") {
      return false;
    }

    try {
      return await checker();
    } catch (error) {
      console.error("LabelMate banner classification check failed:", error);
      return false;
    }
  }

  async function refresh() {
    if (refreshInFlight) {
      return refreshInFlight;
    }

    refreshInFlight = (async function () {
      const exists = await hasClassificationSafe();

      if (exists) {
        hide();
      } else {
        show(getDefaultBannerMessage());
      }
    })();

    try {
      await refreshInFlight;
    } finally {
      refreshInFlight = null;
    }
  }

  function init() {
    if (initialized) {
      return;
    }

    initialized = true;
    ensureBannerElement();
    refresh();
  }

  function scheduleRefresh(delayMs) {
    window.setTimeout(function () {
      refresh();
    }, delayMs);
  }

  function wireLifecycle() {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
      init();
    }

    if (window.Office?.onReady) {
      Office.onReady(function () {
        init();
        refresh();

        // Best-effort extra refreshes because Office host objects
        // can sometimes become fully available slightly after onReady.
        scheduleRefresh(300);
        scheduleRefresh(1000);
      });
    }

    window.addEventListener("labelmate:classification-changed", function () {
      refresh();
    });

    window.addEventListener("labelmate:rerender-labels", function () {
      refresh();
    });

    window.addEventListener("focus", function () {
      refresh();
    });

    document.addEventListener("visibilitychange", function () {
      if (!document.hidden) {
        refresh();
      }
    });
  }

  window.LMBanner = {
    init,
    refresh,
    show,
    hide,
    getState: function () {
      return lastKnownState;
    }
  };

  // Backward-compatible alias
  window.refreshBanner = refresh;

  wireLifecycle();
})();