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
        officeInfoCache?.host ||
        Office?.context?.diagnostics?.host ||
        Office?.context?.host ||
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

  async function markDocumentForAutoOpen() {
  try {
    await ensureOfficeReady();
    const settings = Office?.context?.document?.settings;
    if (!settings || typeof settings.set !== "function") return;

    if (settings.get("Office.AutoShowTaskpaneWithDocument") === true) return;

    settings.set("Office.AutoShowTaskpaneWithDocument", true);

    await new Promise((resolve) => {
      settings.saveAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("AutoShowTaskpane save failed:", result.error);
        }
        resolve(); 
      });
    });
  } catch (e) {
    console.warn("markDocumentForAutoOpen:", e);
  }
}

  async function apply(label) {
    await ensureOfficeReady();

    const engine = await getEngine();
    if (!engine || typeof engine.apply !== "function") {
      throw new Error("Unsupported host or missing classification engine.");
    }

    const result = await engine.apply(label);

    markDocumentForAutoOpen().catch(() => {});

    try {
      window.dispatchEvent(
        new CustomEvent("labelmate:classification-changed", {
          detail: {
            label,
            host: getHostFromContext()
          }
        })
      );
    } catch (_) {}

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
    ensureOfficeReady
  };
})();