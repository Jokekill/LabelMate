// classification.powerpoint.js
window.LMPowerPointClassification = (function () {
  const SHAPE_NAME = "LABELMATE_CLASSIFICATION_FOOTER";

  async function apply(label) {
    // Slide masters exist in PowerPointApi 1.3. citeturn7search5turn7search8
    await PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters;
      masters.load("items");
      await context.sync();

      for (const master of masters.items) {
        const shapes = master.shapes;
        shapes.load("items/name,items/type,items/left,items/top,items/width,items/height,items/textFrame/hasText");
        await context.sync();

        // Find existing LabelMate shape by name
        const existing = shapes.items.find((s) => s.name === SHAPE_NAME);
        if (existing) {
          existing.textFrame.textRange.text = label;
          continue;
        }

        // Create a master-level text box as a “fake footer”. citeturn5search0turn6search0
        const tb = shapes.addTextBox(label);
        tb.name = SHAPE_NAME;

        // Heuristic placement (bottom-ish, centered). You may tune after testing on your corporate slide size.
        tb.left = 0;
        tb.top = 510;
        tb.width = 960;
        tb.height = 24;

        // Optional formatting (font APIs are part of the PowerPoint text model). citeturn6search0
        tb.textFrame.textRange.font.bold = true;
        tb.textFrame.textRange.font.color = "b91c1c";
      }

      await context.sync();
    });
  }

  async function hasClassification() {
    let found = false;

    await PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters;
      masters.load("items");
      await context.sync();

      for (const master of masters.items) {
        const shapes = master.shapes;
        shapes.load("items/name");
        await context.sync();

        if (shapes.items.some((s) => s.name === SHAPE_NAME)) {
          found = true;
          return;
        }
      }
    });

    return found;
  }

  return { apply, hasClassification };
})();
