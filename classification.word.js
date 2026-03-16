// classification.word.js
window.LMWordClassification = (function () {
  "use strict";

  const CC_TAG = "LABELMATE_CLASSIFICATION";
  const CC_TITLE = "LabelMate Classification";
  const CC_COLOR = "#C00000";

  function normalizeLabel(label) {
    return String(label || "").replace(/\s+/g, " ").trim();
  }

  async function getSections(context) {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    return sections.items || [];
  }

  async function getPrimaryHeaderBodies(context) {
    const sections = await getSections(context);
    return sections.map((section) => section.getHeader(Word.HeaderFooterType.primary).body);
  }

  async function loadContentControls(context, body, tag) {
    const controls = body.contentControls.getByTag(tag);
    controls.load("items");
    await context.sync();

    for (const item of controls.items) {
      item.load(["id", "text", "tag", "title", "cannotEdit", "cannotDelete"]);
    }

    await context.sync();
    return controls.items || [];
  }

  function applyStandardProperties(contentControl) {
    contentControl.tag = CC_TAG;
    contentControl.title = CC_TITLE;
    contentControl.cannotEdit = true;
    contentControl.cannotDelete = true;

    if (Word.ContentControlAppearance && Word.ContentControlAppearance.hidden) {
      contentControl.appearance = Word.ContentControlAppearance.hidden;
    }
  }

  function applyStandardFormattingToControl(contentControl) {
    const range = contentControl.getRange(Word.RangeLocation.content);
    range.font.bold = true;
    range.font.color = CC_COLOR;
    range.font.size = 11;
  }

  function createHeaderControl(headerBody, label) {
    const paragraph = headerBody.insertParagraph(label, Word.InsertLocation.start);
    paragraph.alignment = Word.Alignment.centered;

    const contentControl = paragraph.insertContentControl();
    applyStandardProperties(contentControl);
    applyStandardFormattingToControl(contentControl);

    return contentControl;
  }

  function updateExistingControl(contentControl, label) {
    contentControl.insertText(label, Word.InsertLocation.replace);
    applyStandardProperties(contentControl);
    applyStandardFormattingToControl(contentControl);

    try {
      const firstParagraph = contentControl.getRange(Word.RangeLocation.content).paragraphs.getFirst();
      firstParagraph.alignment = Word.Alignment.centered;
    } catch (_) {
      // Alignment is best-effort only.
    }
  }

  async function cleanupBodyControls(context) {
    const bodyControls = await loadContentControls(context, context.document.body, CC_TAG);

    for (const control of bodyControls) {
      control.delete(false);
    }

    if (bodyControls.length > 0) {
      await context.sync();
    }
  }

  async function cleanupDuplicateControlsInHeader(context, headerBody) {
    const controls = await loadContentControls(context, headerBody, CC_TAG);

    if (controls.length <= 1) {
      return controls;
    }

    const keeper = controls[0];
    for (let i = 1; i < controls.length; i += 1) {
      controls[i].delete(false);
    }

    await context.sync();
    return [keeper];
  }

  async function upsertClassificationIntoAllPrimaryHeaders(context, label) {
    const headerBodies = await getPrimaryHeaderBodies(context);

    for (const headerBody of headerBodies) {
      const dedupedControls = await cleanupDuplicateControlsInHeader(context, headerBody);

      if (dedupedControls.length > 0) {
        updateExistingControl(dedupedControls[0], label);
      } else {
        createHeaderControl(headerBody, label);
      }

      await context.sync();
    }
  }

  async function verifyClassificationSet(context, label) {
    const normalizedExpected = normalizeLabel(label);
    const headerBodies = await getPrimaryHeaderBodies(context);

    if (headerBodies.length === 0) {
      return false;
    }

    for (const headerBody of headerBodies) {
      const controls = await loadContentControls(context, headerBody, CC_TAG);

      if (controls.length === 0) {
        return false;
      }

      const controlText = normalizeLabel(controls[0].text);
      if (controlText !== normalizedExpected) {
        return false;
      }
    }

    const bodyControls = await loadContentControls(context, context.document.body, CC_TAG);
    return bodyControls.length === 0;
  }

  async function hasClassification() {
    let exists = false;

    await Word.run(async (context) => {
      const bodyControls = await loadContentControls(context, context.document.body, CC_TAG);
      if (bodyControls.length > 0) {
        exists = true;
        return;
      }

      const headerBodies = await getPrimaryHeaderBodies(context);

      for (const headerBody of headerBodies) {
        const controls = await loadContentControls(context, headerBody, CC_TAG);
        if (controls.length > 0) {
          exists = true;
          return;
        }
      }
    });

    return exists;
  }

  async function apply(label) {
    const normalizedLabel = normalizeLabel(label);

    if (!normalizedLabel) {
      throw new Error("Classification label is empty.");
    }

    await Word.run(async (context) => {
      await cleanupBodyControls(context);
      await upsertClassificationIntoAllPrimaryHeaders(context, normalizedLabel);

      const ok = await verifyClassificationSet(context, normalizedLabel);
      if (!ok) {
        throw new Error("Classification was not written correctly into all primary headers.");
      }
    });

    return true;
  }

  return {
    apply,
    hasClassification,
    constants: {
      CC_TAG,
      CC_TITLE,
      CC_COLOR
    }
  };
})();