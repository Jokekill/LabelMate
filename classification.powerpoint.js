// classification.powerpoint.js
window.LMPowerPointClassification = (function () {
  "use strict";

  const SHAPE_NAME_PREFIX = "LABELMATE_CLASSIFICATION_FOOTER";
  const FALLBACK_SLIDE_WIDTH = 960;
  const FALLBACK_SLIDE_HEIGHT = 540;

  const FONT_COLOR = "#B91C1C";
  const FONT_SIZE = 12;

  const BOX_LEFT = 24;
  const BOX_HEIGHT = 18;
  const BOX_BOTTOM_MARGIN = 20;
  const BOX_RIGHT_MARGIN = 24;

  function normalizeText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function getKnownLabelTexts() {
    const values = new Set();
    const i18n = window.LM_I18N || {};

    Object.values(i18n).forEach((lang) => {
      (lang?.labels || []).forEach((item) => {
        if (item?.text) values.add(normalizeText(item.text));
      });
    });

    return values;
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

  async function getSlideSize(context) {
    if (!supportsPowerPointApi("1.10")) {
      return {
        width: FALLBACK_SLIDE_WIDTH,
        height: FALLBACK_SLIDE_HEIGHT
      };
    }

    try {
      const pageSetup = context.presentation.pageSetup;
      pageSetup.load("slideWidth,slideHeight");
      await context.sync();

      return {
        width: pageSetup.slideWidth || FALLBACK_SLIDE_WIDTH,
        height: pageSetup.slideHeight || FALLBACK_SLIDE_HEIGHT
      };
    } catch (_) {
      return {
        width: FALLBACK_SLIDE_WIDTH,
        height: FALLBACK_SLIDE_HEIGHT
      };
    }
  }

  function getFooterBox(slideSize) {
    const width = Math.max(
      220,
      (slideSize.width || FALLBACK_SLIDE_WIDTH) - BOX_LEFT - BOX_RIGHT_MARGIN
    );
    const height = BOX_HEIGHT;
    const left = BOX_LEFT;
    const top = Math.max(
      0,
      (slideSize.height || FALLBACK_SLIDE_HEIGHT) - BOX_BOTTOM_MARGIN - height
    );

    return { left, top, width, height };
  }

  function getShapeNameForSlide(index) {
    return `${SHAPE_NAME_PREFIX}_${index + 1}`;
  }

  function isManagedShapeName(name) {
    const n = normalizeText(name);
    return n === SHAPE_NAME_PREFIX || n.startsWith(`${SHAPE_NAME_PREFIX}_`);
  }

  function applyTextFormatting(shape, label) {
    const textFrame = shape.textFrame;
    const textRange = textFrame.textRange;

    textRange.text = label;

    try { textRange.font.bold = true; } catch (_) {}
    try { textRange.font.size = FONT_SIZE; } catch (_) {}
    try { textRange.font.color = FONT_COLOR; } catch (_) {}

    try { textFrame.leftMargin = 0; } catch (_) {}
    try { textFrame.rightMargin = 0; } catch (_) {}
    try { textFrame.topMargin = 0; } catch (_) {}
    try { textFrame.bottomMargin = 0; } catch (_) {}
    try { textFrame.wordWrap = false; } catch (_) {}

    // Bezpečnější než transparency = 100
    try { shape.fill.clear(); } catch (_) {}
    try { shape.lineFormat.visible = false; } catch (_) {}
  }

  async function loadSlidesAndShapeNames(context) {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    if (!slides.items || slides.items.length === 0) {
      return slides;
    }

    for (const slide of slides.items) {
      slide.shapes.load("items/name");
    }
    await context.sync();

    return slides;
  }

  function findManagedShape(slide) {
    const items = slide.shapes?.items || [];
    for (const shape of items) {
      if (isManagedShapeName(shape.name)) {
        return shape;
      }
    }
    return null;
  }

  function upsertOnSlide(slide, slideIndex, label, slideSize) {
    const existing = findManagedShape(slide);
    const box = getFooterBox(slideSize);

    let shape = existing;
    if (!shape) {
      shape = slide.shapes.addTextBox(label, box);
    } else {
      try { shape.left = box.left; } catch (_) {}
      try { shape.top = box.top; } catch (_) {}
      try { shape.width = box.width; } catch (_) {}
      try { shape.height = box.height; } catch (_) {}
    }

    try { shape.name = getShapeNameForSlide(slideIndex); } catch (_) {}

    applyTextFormatting(shape, label);
    return true;
  }

  async function collectSlideShapeData(context, slides) {
    const slideData = [];

    for (const slide of slides.items) {
      const shape = findManagedShape(slide);
      if (shape) {
        try {
          shape.textFrame.textRange.load("text");
        } catch (_) {}
      }
      slideData.push({ slide, shape, text: "" });
    }

    await context.sync();

    for (const data of slideData) {
      if (!data.shape) continue;
      try {
        data.text = normalizeText(data.shape.textFrame.textRange.text);
      } catch (_) {
        data.text = "";
      }
    }

    return slideData;
  }

  function findFirstValidLabel(slideData, knownLabels) {
    for (const data of slideData) {
      if (data.shape && data.text && knownLabels.has(data.text)) {
        return data.text;
      }
    }
    return null;
  }

  function syncAllSlidesWithLabel(slideData, label, slideSize) {
    let changedCount = 0;

    slideData.forEach((data, index) => {
      const needsUpdate = !data.shape || data.text !== label;
      if (!needsUpdate) return;

      upsertOnSlide(data.slide, index, label, slideSize);
      changedCount += 1;
    });

    return changedCount;
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
      const slideSize = await getSlideSize(context);
      const slides = await loadSlidesAndShapeNames(context);

      if (!slides.items || slides.items.length === 0) {
        throw new Error("Presentation has no slides.");
      }

      slides.items.forEach((slide, index) => {
        const ok = upsertOnSlide(slide, index, normalizedLabel, slideSize);
        if (ok) updatedCount += 1;
      });

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

    let classified = false;

    await PowerPoint.run(async (context) => {
      const slides = await loadSlidesAndShapeNames(context);

      if (!slides.items || slides.items.length === 0) {
        classified = false;
        return;
      }

      const slideData = await collectSlideShapeData(context, slides);
      const knownLabels = getKnownLabelTexts();
      const validLabel = findFirstValidLabel(slideData, knownLabels);

      if (!validLabel) {
        classified = false;
        return;
      }

      const slideSize = await getSlideSize(context);
      const changed = syncAllSlidesWithLabel(slideData, validLabel, slideSize);
      if (changed > 0) {
        await context.sync();
      }

      classified = true;
    });

    return classified;
  }

  return {
    apply,
    hasClassification,
    constants: {
      SHAPE_NAME_PREFIX,
      FONT_COLOR,
      FONT_SIZE
    }
  };
})();