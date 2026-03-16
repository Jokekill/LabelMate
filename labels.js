// labels.js
(function () {
  "use strict";

  window.LM = window.LM || {};

  const DEFAULT_LABELS = [
    { id: "public", text: "PUBLIC", color: "#166534", textColor: "#ffffff" },
    { id: "internal", text: "INTERNAL", color: "#1d4ed8", textColor: "#ffffff" },
    { id: "confidential", text: "CONFIDENTIAL", color: "#b45309", textColor: "#ffffff" },
    { id: "strictly-confidential", text: "STRICTLY CONFIDENTIAL", color: "#b91c1c", textColor: "#ffffff" }
  ];

  let initialized = false;
  let busy = false;

  function slugify(value) {
    return String(value || "")
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "");
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function getHostName() {
    try {
      return window.LM?.classification?.getHost?.() || Office?.context?.diagnostics?.host || null;
    } catch (_) {
      return null;
    }
  }

  function resolveContainer() {
    return (
      document.getElementById("labels-container") ||
      document.getElementById("labelsContainer") ||
      document.getElementById("labels") ||
      document.getElementById("classification-list") ||
      document.getElementById("classificationList")
    );
  }

  function resolveStatusElement() {
    return (
      document.getElementById("labels-status") ||
      document.getElementById("labelsStatus") ||
      document.getElementById("status")
    );
  }

  function resolveTitleElement() {
    return (
      document.getElementById("labels-title") ||
      document.getElementById("labelsTitle") ||
      document.getElementById("classification-title")
    );
  }

  function resolveConfiguredLabels() {
    if (Array.isArray(window.LM_CONFIG?.labels) && window.LM_CONFIG.labels.length > 0) {
      return window.LM_CONFIG.labels;
    }

    if (Array.isArray(window.LM_LABELS) && window.LM_LABELS.length > 0) {
      return window.LM_LABELS;
    }

    if (Array.isArray(window.LABELS) && window.LABELS.length > 0) {
      return window.LABELS;
    }

    return DEFAULT_LABELS;
  }

  function normalizeLabelItem(item) {
    if (typeof item === "string") {
      return {
        id: slugify(item),
        text: item,
        color: "#b91c1c",
        textColor: "#ffffff"
      };
    }

    const text =
      item.text ||
      item.label ||
      item.name ||
      item.value ||
      item.classification ||
      "";

    return {
      id: item.id || slugify(text),
      text,
      color: item.color || item.backgroundColor || "#b91c1c",
      textColor: item.textColor || item.foregroundColor || "#ffffff",
      description: item.description || ""
    };
  }

  function getLabels() {
    return resolveConfiguredLabels()
      .map(normalizeLabelItem)
      .filter((item) => item.text && item.text.trim().length > 0);
  }

  function ensureContainerExists() {
    let container = resolveContainer();
    if (container) {
      return container;
    }

    const root =
      document.getElementById("app") ||
      document.querySelector("main") ||
      document.body;

    container = document.createElement("div");
    container.id = "labels-container";
    root.appendChild(container);
    return container;
  }

  function setTitle() {
    const titleEl = resolveTitleElement();
    if (!titleEl) {
      return;
    }

    const host = getHostName();
    let suffix = "";

    if (host === Office.HostType.Word) suffix = "Word";
    if (host === Office.HostType.Excel) suffix = "Excel";
    if (host === Office.HostType.PowerPoint) suffix = "PowerPoint";

    titleEl.textContent = suffix ? `Classification (${suffix})` : "Classification";
  }

  function setStatus(message, type) {
    const el = resolveStatusElement();
    if (!el) {
      return;
    }

    el.textContent = message || "";
    el.dataset.state = type || "";

    el.style.display = message ? "block" : "none";
  }

  function setBusyState(nextBusy) {
    busy = !!nextBusy;

    const buttons = document.querySelectorAll("[data-lm-label-button]");
    buttons.forEach((button) => {
      button.disabled = busy;
      button.setAttribute("aria-busy", busy ? "true" : "false");
    });
  }

  function buildButton(item) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "labelmate-label-button";
    button.dataset.lmLabelButton = "true";
    button.dataset.labelId = item.id;
    button.dataset.labelText = item.text;

    button.style.background = item.color;
    button.style.color = item.textColor;
    button.style.border = "0";
    button.style.borderRadius = "8px";
    button.style.padding = "12px 14px";
    button.style.cursor = "pointer";
    button.style.fontWeight = "700";
    button.style.width = "100%";
    button.style.textAlign = "left";
    button.style.marginBottom = "10px";
    button.style.boxSizing = "border-box";

    const title = document.createElement("div");
    title.className = "labelmate-label-button__title";
    title.textContent = item.text;

    button.appendChild(title);

    if (item.description) {
      const desc = document.createElement("div");
      desc.className = "labelmate-label-button__description";
      desc.textContent = item.description;
      desc.style.fontWeight = "400";
      desc.style.fontSize = "12px";
      desc.style.opacity = "0.95";
      desc.style.marginTop = "4px";
      button.appendChild(desc);
    }

    button.addEventListener("click", async () => {
      await applyLabel(item);
    });

    return button;
  }

  function renderLabels() {
    const container = ensureContainerExists();
    const labels = getLabels();

    container.innerHTML = "";
    container.setAttribute("data-host", escapeHtml(getHostName() || ""));

    if (labels.length === 0) {
      const empty = document.createElement("div");
      empty.textContent = "No labels configured.";
      container.appendChild(empty);
      return;
    }

    labels.forEach((item) => {
      container.appendChild(buildButton(item));
    });
  }

  async function refreshBannerIfPossible() {
    window.dispatchEvent(new CustomEvent("labelmate:classification-changed"));

    if (window.LMBanner?.refresh && typeof window.LMBanner.refresh === "function") {
      try {
        await window.LMBanner.refresh();
      } catch (_) {
        // Ignore banner refresh issues.
      }
    }

    if (window.refreshBanner && typeof window.refreshBanner === "function") {
      try {
        await window.refreshBanner();
      } catch (_) {
        // Ignore banner refresh issues.
      }
    }
  }

  async function applyLabel(item) {
    if (busy) {
      return;
    }

    const classificationText = String(item?.text || "").trim();
    if (!classificationText) {
      setStatus("Selected label is empty.", "error");
      return;
    }

    if (!window.LM?.classification?.apply || typeof window.LM.classification.apply !== "function") {
      setStatus("Classification engine is not available.", "error");
      return;
    }

    try {
      setBusyState(true);
      setStatus(`Applying "${classificationText}"...`, "info");

      await window.LM.classification.apply(classificationText);

      try {
        localStorage.setItem("labelmate.lastLabel", classificationText);
      } catch (_) {
        // Ignore localStorage issues.
      }

      setStatus(`Applied: ${classificationText}`, "success");
      await refreshBannerIfPossible();
    } catch (error) {
      const message =
        error?.message ||
        error?.toString?.() ||
        "Unknown error while applying classification.";

      console.error("LabelMate apply failed:", error);
      setStatus(message, "error");
    } finally {
      setBusyState(false);
    }
  }

  function init() {
    if (initialized) {
      return;
    }

    initialized = true;
    setTitle();
    renderLabels();
  }

  function wireLifecycle() {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
      init();
    }

    if (window.Office?.onReady) {
      Office.onReady(() => {
        setTitle();
        renderLabels();
      });
    }

    window.addEventListener("labelmate:rerender-labels", () => {
      setTitle();
      renderLabels();
    });
  }

  window.LM.labels = {
    init,
    render: renderLabels,
    applyLabel
  };

  wireLifecycle();
})();