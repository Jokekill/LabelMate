// classification.host.js
window.LM = window.LM || {};

(function () {
  const KEY = "labelmate.classification"; // optional: store state in document settings

  function getHost() {
    try {
      // Office.context.diagnostics.host is a standard way to read host type. citeturn10search14
      return Office?.context?.diagnostics?.host || null;
    } catch {
      return null;
    }
  }

  function getEngine() {
    const host = getHost();
    if (host === Office.HostType.Word) return window.LMWordClassification;
    if (host === Office.HostType.Excel) return window.LMExcelClassification;
    if (host === Office.HostType.PowerPoint) return window.LMPowerPointClassification;
    return null;
  }

  async function apply(label) {
    const engine = getEngine();
    if (!engine) throw new Error("Unsupported host.");
    return engine.apply(label);
  }

  async function hasClassification() {
    const engine = getEngine();
    if (!engine) return false;
    return engine.hasClassification();
  }

  window.LM.classification = { apply, hasClassification, getHost };
})();
