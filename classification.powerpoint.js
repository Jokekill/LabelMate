// classification.powerpoint.js
// PowerPoint implementation for LabelMate.
// Goal: keep Word/Excel untouched and apply the classification on slide masters,
// ideally into an existing footer placeholder, otherwise into a named master text box.
window.LMPowerPointClassification = (function () {
  "use strict";

  const SHAPE_NAME = "LABELMATE_CLASSIFICATION_FOOTER";
  const FALLBACK_SLIDE_WIDTH = 960;
  const FALLBACK_SLIDE_HEIGHT = 540;
  const FONT_COLOR = "#B91C1C";
  const FONT_SIZE = 14;

  function supportsPowerPointApi(version) {
    try {
      return Office.context.requirements.isSetSupported("PowerPointApi", version);
    } catch (_) {
      return false;
    }
  }

  function normalizeText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function getAllLabelTexts() {
    const all = new Set();
    const i18n = window.LM_I18N || {};

    Object.values(i18n).forEach((lang) => {
      (lang?.labels || []).forEach((label) => {
        if (label?.text) {
          all.add(normalizeText(label.text));
        }
      });
    });

    return Array.from(all);
  }

  async function getSlideSize(context) {
    if (!supportsPowerPointApi("1.10")) {
      return {
        width: FALLBACK_SLIDE_WIDTH,
        height: FALLBACK_SLIDE_HEIGHT,
      };
    }

    const pageSetup = context.presentation.pageSetup;
    pageSetup.load("slideWidth,slideHeight");
    await context.sync();

    return {
      width: pageSetup.slideWidth || FALLBACK_SLIDE_WIDTH,
      height: pageSetup.slideHeight || FALLBACK_SLIDE_HEIGHT,
    };
  }

  function getFallbackFooterBox(size) {
    return {
      left: 24,
      top: Math.max(0, size.height - 28),
      width: Math.max(240, size.width - 48),
      height: 20,
    };
  }

  function applyTextFormatting(shape, label) {
    const textFrame = shape.textFrame;
    const textRange = textFrame.textRange;

    textRange.text = label;
    textRange.font.bold = true;
    textRange.font.size = FONT_SIZE;
    textRange.font.color = FONT_COLOR;

    try { textFrame.leftMargin = 0; } catch (_) {}
    try { textFrame.rightMargin = 0; } catch (_) {}
    try { textFrame.topMargin = 0; } catch (_) {}
    try { textFrame.verticalAlignment = "Bottom"; } catch (_) {}
  }

  async function loadMasterShapes(master) {
    const shapes = master.shapes;
    shapes.load("items/name,items/type");
    await master.context.sync();
    return shapes.items || [];
  }

  async function findNamedLabelMateShape(master) {
    const shapes = await loadMasterShapes(master);
    return shapes.find((shape) => shape.name === SHAPE_NAME) || null;
  }

  async function findFooterPlaceholder(master) {
    if (!supportsPowerPointApi("1.8")) {
      return null;
    }

    const shapes = await loadMasterShapes(master);
    const placeholderShapes = shapes.filter((shape) => normalizeText(shape.type).toLowerCase() === "placeholder");

    for (const shape of placeholderShapes) {
      try {
        shape.placeholderFormat.load("type");
      } catch (_) {
        // ignore non-placeholder edge cases
      }
    }

    await master.context.sync();

    for (const shape of placeholderShapes) {
      try {
        const placeholderType = normalizeText(shape.placeholderFormat.type).toLowerCase();
        if (placeholderType === "footer") {
          return shape;
        }
      } catch (_) {
        // ignore shapes that don't expose placeholderFormat properly
      }
    }

    return null;
  }

  async function upsertOnMaster(master, label, slideSize) {
    const existingNamedShape = await findNamedLabelMateShape(master);
    if (existingNamedShape) {
      applyTextFormatting(existingNamedShape, label);
      return;
    }

    const footerPlaceholder = await findFooterPlaceholder(master);
    if (footerPlaceholder) {
      applyTextFormatting(footerPlaceholder, label);
      return;
    }

    const box = getFallbackFooterBox(slideSize);
    const newShape = master.shapes.addTextBox(label, box);
    newShape.name = SHAPE_NAME;
    applyTextFormatting(newShape, label);
  }

  async function apply(label) {
    const normalizedLabel = normalizeText(label);
    if (!normalizedLabel) {
      throw new Error("Classification label is empty.");
    }

    if (!supportsPowerPointApi("1.4")) {
      throw new Error("PowerPointApi 1.4 is not supported in this client.");
    }

    await PowerPoint.run(async (context) => {
      const slideSize = await getSlideSize(context);
      const masters = context.presentation.slideMasters;
      masters.load("items");
      await context.sync();

      if (!masters.items || masters.items.length === 0) {
        throw new Error("No slide masters were found in the presentation.");
      }

      for (const master of masters.items) {
        await upsertOnMaster(master, normalizedLabel, slideSize);
      }

      await context.sync();
    });

    return true;
  }

  async function hasClassification() {
    if (!supportsPowerPointApi("1.4")) {
      return false;
    }

    const knownLabels = new Set(getAllLabelTexts());
    let found = false;

    await PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters;
      masters.load("items");
      await context.sync();

      for (const master of masters.items) {
        const namedShape = await findNamedLabelMateShape(master);
        if (namedShape) {
          found = true;
          return;
        }

        const footerPlaceholder = await findFooterPlaceholder(master);
        if (!footerPlaceholder) {
          continue;
        }

        try {
          footerPlaceholder.textFrame.load("hasText,textRange/text");
          await context.sync();

          const footerText = normalizeText(footerPlaceholder.textFrame.textRange.text);
          if (footerText && knownLabels.has(footerText)) {
            found = true;
            return;
          }
        } catch (_) {
          // ignore placeholder text read failures and continue checking other masters
        }
      }
    });

    return found;
  }

  return {
    apply,
    hasClassification,
    constants: {
      SHAPE_NAME,
      FONT_COLOR,
      FONT_SIZE,
    },
  };
})();
