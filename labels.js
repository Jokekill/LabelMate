// labels.js
window.Labels = window.Labels || {};

(function () {
  "use strict";

  let openTooltip = null;
  let tooltipGlobalsBound = false;
  let running = false;

  function getBannerApi() {
    return window.Banner || window.LMBanner || null;
  }

  function bindTooltipGlobalsOnce() {
    if (tooltipGlobalsBound) return;
    tooltipGlobalsBound = true;

    document.addEventListener(
      "click",
      (ev) => {
        if (!openTooltip) return;
        const target = ev.target;
        if (target && openTooltip.wrap.contains(target)) return;
        closeTooltip();
      },
      true
    );

    document.addEventListener("keydown", (ev) => {
      if (!openTooltip) return;
      if (ev.key === "Escape" || ev.key === "Esc") {
        ev.preventDefault();
        closeTooltip();
      }
    });
  }

  function closeTooltip() {
    if (!openTooltip) return;
    try {
      openTooltip.wrap.classList.remove("tooltip-open");
      openTooltip.btn.setAttribute("aria-expanded", "false");
    } catch (_) {
      // ignore cleanup issues
    }
    openTooltip = null;
  }

  function toggleTooltip(wrap, btn) {
    bindTooltipGlobalsOnce();

    const isOpen = wrap.classList.contains("tooltip-open");
    if (isOpen) {
      closeTooltip();
      return;
    }

    closeTooltip();
    wrap.classList.add("tooltip-open");
    btn.setAttribute("aria-expanded", "true");
    openTooltip = { wrap, btn };
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (m) => {
      return {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[m];
    });
  }

  async function applyClassification(label) {
    if (running) return;
    running = true;

    const L = window.LM?.i18n?.T?.() || {};
    const banner = getBannerApi();

    closeTooltip();
    window.LM?.ui?.clearStatus?.();
    window.LM?.ui?.setBusy?.(true);

    try {
      if (!window.LM?.classification?.apply || typeof window.LM.classification.apply !== "function") {
        throw new Error("Classification engine is not available.");
      }

      const result = await window.LM.classification.apply(label);
      const successMessage = typeof L.statusOk === "function"
        ? L.statusOk(label)
        : `Classification "${label}" was set successfully.`;

      if (typeof result === "number" && result > 0) {
        window.LM?.ui?.setStatusOk?.(`${successMessage} (${result})`);
      } else {
        window.LM?.ui?.setStatusOk?.(successMessage);
      }
    } catch (err) {
      console.error(err);
      const fallback = typeof L.statusErrVerify === "function"
        ? L.statusErrVerify(label, "", err?.message || "Unknown error")
        : (err?.message || String(err));
      window.LM?.ui?.setStatusError?.(fallback);
    } finally {
      try {
        window.dispatchEvent(new CustomEvent("labelmate:classification-changed"));
      } catch (_) {
        // ignore custom event issues
      }

      try {
        if (banner?.refresh) {
          await banner.refresh();
        } else if (typeof window.refreshBanner === "function") {
          await window.refreshBanner();
        }
      } catch (_) {
        // ignore banner refresh errors
      }

      window.LM?.ui?.setBusy?.(false);
      running = false;
    }
  }

  function renderButtons() {
    closeTooltip();

    const container = document.getElementById("labelsContainer");
    if (!container) return;
    container.innerHTML = "";

    const L = window.LM?.i18n?.T?.() || { labels: [], docLinkText: "More in documentation" };
    const infoAria = L.tooltipMoreInfoAria || "More information";

    (L.labels || []).forEach((item, idx) => {
      const row = document.createElement("div");
      row.className = "label-row";

      const wrap = document.createElement("div");
      wrap.className = "classify-wrap";

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn primary pill classify-btn";
      btn.textContent = item.text;
      btn.addEventListener("click", () => {
        closeTooltip();
        applyClassification(item.text);
      });

      const infoWrap = document.createElement("div");
      infoWrap.className = "info-in-btn";

      const infoBtn = document.createElement("button");
      infoBtn.type = "button";
      infoBtn.className = "btn icon info-btn";
      infoBtn.setAttribute("aria-label", infoAria);
      infoBtn.setAttribute("aria-expanded", "false");

      const tooltipId = `lm-tooltip-${idx}`;
      infoBtn.setAttribute("aria-controls", tooltipId);
      infoBtn.innerHTML = `<span aria-hidden="true">ℹ️</span>`;
      infoBtn.tabIndex = 0;

      const bubble = document.createElement("div");
      bubble.id = tooltipId;
      bubble.className = "tooltip-bubble";
      bubble.setAttribute("role", "tooltip");

      let inner;
      if (item.helpHtml) {
        inner = item.helpHtml;
      } else {
        const tip = escapeHtml(item.tip || "");
        const link = item.docUrl
          ? ` <a href="${item.docUrl}" target="_blank" rel="noopener">${escapeHtml(L.docLinkText || "More in documentation")}</a>.`
          : "";
        inner = `${tip}${link}`;
      }

      bubble.innerHTML = `<strong>${escapeHtml(item.text)}</strong><br>${inner}`;

      infoBtn.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        toggleTooltip(wrap, infoBtn);
      });

      infoWrap.appendChild(infoBtn);
      wrap.appendChild(btn);
      wrap.appendChild(infoWrap);
      wrap.appendChild(bubble);
      row.appendChild(wrap);
      container.appendChild(row);
    });

    bindTooltipGlobalsOnce();

    try {
      window.dispatchEvent(new CustomEvent("labelmate:rerender-labels"));
    } catch (_) {
      // ignore custom event issues
    }
  }

  window.Labels = {
    renderButtons,
    applyClassification,
  };
})();
