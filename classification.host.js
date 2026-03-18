// classification.host.js
window.LM = window.LM || {};

(function () {
  "use strict";

  let officeReadyPromise = null;
  let officeInfoCache = null;

  function normalizeHost(host) {
    return String(host || "").trim().toLowerCase();
  }

  function hostLooksLike(host, expected) {
    return normalizeHost(host) === normalizeHost(expected);
  }

  function getHostFromContext() {
    try {
      return (
        Office?.context?.diagnostics?.host ||
        Office?.context?.host ||
        officeInfoCache?.host ||
        null
      );
    } catch (_) {
      return officeInfoCache?.host || null;
    }
  }

  function ensureOfficeReady() {
    if (officeReadyPromise) return officeReadyPromise;

    officeReadyPromise = new Promise((resolve, reject) => {
      if (typeof Office === "undefined" || !Office || typeof Office.onReady !== "function") {
        reject(new Error("Office.js is not loaded."));
        return;
      }

      Office.onReady((info) => {
        officeInfoCache = info || officeInfoCache || {};
        resolve(officeInfoCache);
      });
    });

    return officeReadyPromise;
  }

  async function getHost() {
    await ensureOfficeReady();
    return getHostFromContext();
  }

  async function getEngine() {
    const host = await getHost();

    if (hostLooksLike(host, "Word")) return window.LMWordClassification || null;
    if (hostLooksLike(host, "Excel")) return window.LMExcelClassification || null;
    if (hostLooksLike(host, "PowerPoint")) return window.LMPowerPointClassification || null;

    return null;
  }

  async function apply(label) {
    await ensureOfficeReady();

    const engine = await getEngine();
    if (!engine || typeof engine.apply !== "function") {
      throw new Error("Unsupported host or missing classification engine.");
    }

    const result = await engine.apply(label);

    try {
      window.dispatchEvent(
        new CustomEvent("labelmate:classification-changed", {
          detail: {
            label,
            host: getHostFromContext(),
          },
        })
      );
    } catch (_) {
      // ignore custom event issues
    }

    return result;
  }

  async function hasClassification() {
    await ensureOfficeReady();

    const engine = await getEngine();
    if (!engine || typeof engine.hasClassification !== "function") {
      return false;
    }

    return engine.hasClassification();
  }

  window.LM.classification = {
    apply,
    hasClassification,
    getHost,
    ensureOfficeReady,
  };
})();
