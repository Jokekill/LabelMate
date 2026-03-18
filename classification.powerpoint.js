window.LMPowerPointClassification = (function () {
  "use strict";

  const SHAPE_NAME = "LABELMATE_CLASSIFICATION_FOOTER";
  const FALLBACK_SLIDE_WIDTH = 960;
  const FALLBACK_SLIDE_HEIGHT = 540;
  const FONT_COLOR = "#B91C1C";
  const FONT_SIZE = 12;
  const BOX_MARGIN_X = 24;
  const BOX_HEIGHT = 18;
  const BOX_BOTTOM_OFFSET = 24;

  function normalizeText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function supportsPowerPointApi(version) {
    try {
      return Office.context.requirements.isSetSupported("PowerPointApi", version);
    } catch (_) {
      return false;
    }
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
        height: FALLBACK_SLIDE_HEIGHT
      };
    }

    const pageSetup = context.presentation.pageSetup;
    pageSetup.load("slideWidth,slideHeight");
    await context.sync();

    return {
      width: pageSetup.slideWidth || FALLBACK_SLIDE_WIDTH,
      height: pageSetup.slideHeight || FALLBACK_SLIDE_HEIGHT
    };
  }

  function getFooterBox(size) {
    const width = Math.max(260, (size.width || FALLBACK_SLIDE_WIDTH) - BOX_MARGIN_X * 2);
    const height = BOX_HEIGHT;
    const left = BOX_MARGIN_X;
    const top = Math.max(0, (size.height || FALLBACK_SLIDE_HEIGHT) - BOX_BOTTOM_OFFSET - height);

    return { left, top, width, height };
  }

  function applyTextFormatting(shape, label) {
    shape.textFrame.textRange.text = label;

    try { shape.textFrame.textRange.font.bold = true; } catch (_) {}
    try { shape.textFrame.textRange.font.size = FONT_SIZE; } catch (_) {}
    try { shape.textFrame.textRange.font.color = FONT_COLOR; } catch (_) {}

    try { shape.textFrame.leftMargin = 0; } catch (_) {}
    try { shape.textFrame.rightMargin = 0; } catch (_) {}
    try { shape.textFrame.topMargin = 0; } catch (_) {}
    try { shape.textFrame.bottomMargin = 0; } catch (_) {}
    try { shape.textFrame.wordWrap = false; } catch (_) {}

    try { shape.fill.transparency = 100; } catch (_) {}
    try { shape.lineFormat.transparency = 100; } catch (_) {}
  }

  function loadSlides(context) {
    const slides = context.presentation.slides;
    slides.load("items");
    return slides;
  }

  async function loadShapeNamesForSlides(context, slides) {
    for (const slide of slides.items) {
      slide.shapes.load("items/name");
    }
    await context.sync();
  }

  function findManagedShape(slide) {
    return (slide.shapes.items || []).find((shape) => shape.name === SHAPE_NAME) || null;
  }

  function ensureShapeOnSlide(slide, label, size) {
    let shape = findManagedShape(slide);

    if (!shape) {
      shape = slide.shapes.addTextBox(label, getFooterBox(size));
      shape.name = SHAPE_NAME;
    }

    applyTextFormatting(shape, label);
    return shape;
  }

  async function apply(label) {
    await ensureReady();

    const normalizedLabel = normalizeText(label);
    if (!normalizedLabel) {
      throw new Error("Classification label is empty.");
    }

    if (!supportsPowerPointApi("1.4")) {
      throw new Error("PowerPointApi 1.4 is not supported in this client.");
    }

    let updatedCount = 0;

    await PowerPoint.run(async (context) => {
      const size = await getSlideSize(context);
      const slides = loadSlides(context);
      await context.sync();

      if (!slides.items || slides.items.length === 0) {
        throw new Error("Presentation has no slides.");
      }

      await loadShapeNamesForSlides(context, slides);

      for (const slide of slides.items) {
        ensureShapeOnSlide(slide, normalizedLabel, size);
        updatedCount += 1;
      }

      await context.sync();
    });

    if (!updatedCount) {
      throw new Error("No slide was updated.");
    }

    return updatedCount;
  }

  async function hasClassification() {
    await ensureReady();

    if (!supportsPowerPointApi("1.4")) {
      return false;
    }

    const knownLabels = new Set(getAllLabelTexts());
    let allSlidesClassified = true;

    await PowerPoint.run(async (context) => {
      const slides = loadSlides(context);
      await context.sync();

      if (!slides.items || slides.items.length === 0) {
        allSlidesClassified = false;
        return;
      }

      await loadShapeNamesForSlides(context, slides);

      const managedShapes = [];

      for (const slide of slides.items) {
        const shape = findManagedShape(slide);
        if (!shape) {
          allSlidesClassified = false;
          return;
        }

        try {
          shape.textFrame.load("hasText,textRange/text");
          managedShapes.push(shape);
        } catch (_) {
          allSlidesClassified = false;
          return;
        }
      }

      await context.sync();

      for (const shape of managedShapes) {
        const text = normalizeText(shape.textFrame?.textRange?.text);
        if (!text) {
          allSlidesClassified = false;
          return;
        }

        if (knownLabels.size && !knownLabels.has(text)) {
          allSlidesClassified = false;
          return;
        }
      }
    });

    return allSlidesClassified;
  }

  return {
    apply,
    hasClassification,
    constants: {
      SHAPE_NAME,
      FONT_COLOR,
      FONT_SIZE
    }
  };
})();