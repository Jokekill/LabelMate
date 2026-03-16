// classification.excel.js
window.LMExcelClassification = (function () {
  // Optional: styling via format codes. citeturn9search1
  function formatHeaderText(label) {
    // Example: bold + 14pt + red (hex). Exact rendering follows Excel header/footer code rules. citeturn9search1
    return `&B&14&KFF0000${label}`;
  }

  async function apply(label) {
    // Guard: ExcelApi 1.9 is required for headers/footers. citeturn9search4turn13search8
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
        // Set center header. API set ExcelApi 1.9. citeturn9search4turn13search8
        hf.set({ centerHeader: headerText });
      }
      await context.sync();
    });
  }

  async function hasClassification() {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) return false;

    let found = false;
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const hf = ws.pageLayout.headersFooters.defaultForAllPages;
      hf.load("centerHeader,leftHeader,rightHeader");
      await context.sync();

      const any = [hf.leftHeader, hf.centerHeader, hf.rightHeader]
        .map((s) => (s || "").trim())
        .join(" ");
      found = any.length > 0; // or stricter: check for one of known label texts
    });
    return found;
  }

  return { apply, hasClassification };
})();
