window.LMExcelClassification = (function () {
  "use strict";

  function formatHeaderText(label) {
    return `&B&14&KFF0000${label}`;
  }

  async function ensureReady() {
    if (window.LM?.classification?.ensureOfficeReady) {
      await window.LM.classification.ensureOfficeReady();
      return;
    }

    if (typeof Office === "undefined" || typeof Office.onReady !== "function") {
      throw new Error("Office.js is not loaded.");
    }

    await Office.onReady();
  }

  function getKnownLabelTexts() {
    const values = new Set();
    const i18n = window.LM_I18N || {};

    Object.values(i18n).forEach((lang) => {
      (lang?.labels || []).forEach((item) => {
        if (item?.text) values.add(String(item.text).trim());
      });
    });

    return Array.from(values);
  }

  async function apply(label) {
    await ensureReady();

    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      throw new Error("ExcelApi 1.9 is not supported in this client.");
    }

    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const headerText = formatHeaderText(label);

      for (const ws of sheets.items) {
        const hf = ws.pageLayout.headersFooters.defaultForAllPages;
        hf.set({ centerHeader: headerText });
      }

      await context.sync();
    });

    return true;
  }

  async function hasClassification() {
    await ensureReady();

    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      return false;
    }

    const knownLabels = new Set(getKnownLabelTexts());
    let found = false;

    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const hf = ws.pageLayout.headersFooters.defaultForAllPages;
      hf.load("centerHeader,leftHeader,rightHeader");
      await context.sync();

      const values = [hf.leftHeader, hf.centerHeader, hf.rightHeader]
        .map((value) => String(value || "").trim())
        .filter(Boolean);

      if (!values.length) {
        found = false;
        return;
      }

      if (!knownLabels.size) {
        found = true;
        return;
      }

      found = values.some((headerValue) => {
        const normalized = headerValue.replace(/&[^A-Z0-9]?/gi, " ").replace(/\s+/g, " ").trim();
        for (const label of knownLabels) {
          if (normalized.includes(label)) return true;
        }
        return false;
      });
    });

    return found;
  }

  return { apply, hasClassification };
})();
