// classification.word.js
// Obnova původního Word chování: header-first content control, cleanup starých body labelů,
// styling jako ve staré verzi a striktní ověření po zápisu.
window.LMWordClassification = (function () {
  "use strict";

  const CC_TAG = "LABELMATE_CLASSIFICATION";
  const CC_TITLE = "Document Classification";
  const HEADER_TYPE = "Primary";

  function normalizeLabel(label) {
    return String(label || "").replace(/\s+/g, " ").trim();
  }

  function getAllLabelTexts() {
    const all = new Set();
    const i18n = window.LM_I18N || {};

    Object.values(i18n).forEach((lang) => {
      (lang?.labels || []).forEach((label) => {
        if (label?.text) all.add(String(label.text).trim());
      });
    });

    return Array.from(all);
  }

  async function cleanupOrphanLabels(context) {
    const paras = context.document.body.paragraphs;
    paras.load("items");
    await context.sync();

    const limit = Math.min(paras.items.length, 8);
    const allLabels = getAllLabelTexts();
    const checks = [];

    for (let i = 0; i < limit; i += 1) {
      const p = paras.items[i];
      p.load("text");
      const firstCC = p.contentControls.getFirstOrNullObject();
      firstCC.load("isNullObject");
      checks.push({ p, firstCC });
    }

    await context.sync();

    for (const item of checks) {
      const txt = normalizeLabel(item.p.text);
      const hasAnyCC = !item.firstCC.isNullObject;
      if (!hasAnyCC && allLabels.includes(txt)) {
        item.p.delete();
      }
    }

    await context.sync();
  }

  function applyClassificationStylingToRange(range) {
    range.font.bold = true;
    range.font.size = 14;
    range.font.color = "#b91c1c";
  }

  function lockAndTagContentControl(cc) {
    cc.tag = CC_TAG;
    cc.title = CC_TITLE;
    cc.cannotEdit = true;
    cc.cannotDelete = true;
    cc.appearance = "Hidden";
    cc.color = "#b91c1c";
  }

  async function getHeaderBodies(context) {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    if (!sections.items || sections.items.length === 0) {
      return [];
    }

    return sections.items.map((section) => section.getHeader(HEADER_TYPE));
  }

  async function upsertClassificationIntoAllPrimaryHeaders(context, label) {
    const headerBodies = await getHeaderBodies(context);

    if (headerBodies.length === 0) {
      const p = context.document.body.insertParagraph(label, "Start");
      const cc = p.insertContentControl();
      lockAndTagContentControl(cc);

      const rng = cc.getRange("Content");
      rng.insertText(label, Word.InsertLocation.replace);
      applyClassificationStylingToRange(rng);

      await context.sync();
      return;
    }

    for (const headerBody of headerBodies) {
      const headerCCs = headerBody.contentControls.getByTag(CC_TAG);
      headerCCs.load("items");
      await context.sync();

      if (headerCCs.items && headerCCs.items.length > 0) {
        for (const cc of headerCCs.items) {
          cc.cannotEdit = false;
          cc.cannotDelete = false;
          await context.sync();

          const rng = cc.getRange("Content");
          rng.insertText(label, Word.InsertLocation.replace);
          applyClassificationStylingToRange(rng);

          lockAndTagContentControl(cc);
          await context.sync();
        }
      } else {
        const p = headerBody.insertParagraph(label, "Start");
        const cc = p.insertContentControl();
        lockAndTagContentControl(cc);

        const rng = cc.getRange("Content");
        rng.insertText(label, Word.InsertLocation.replace);
        applyClassificationStylingToRange(rng);

        await context.sync();
      }
    }

    const bodyCCs = context.document.body.contentControls.getByTag(CC_TAG);
    bodyCCs.load("items");
    await context.sync();

    if (bodyCCs.items && bodyCCs.items.length > 0) {
      for (const cc of bodyCCs.items) {
        cc.delete(false);
      }
      await context.sync();
    }
  }

  async function verifyClassificationSetInternal(context, expectedLabel) {
    const result = { ok: false, foundText: "", reason: "" };
    const headerBodies = await getHeaderBodies(context);
    const bodiesToCheck = [context.document.body, ...headerBodies];

    const collections = bodiesToCheck.map((body) => body.contentControls.getByTag(CC_TAG));
    collections.forEach((coll) => coll.load("items"));
    await context.sync();

    const ranges = [];
    for (const coll of collections) {
      if (!coll.items) continue;
      for (const cc of coll.items) {
        const rng = cc.getRange("Content");
        rng.load("text");
        ranges.push(rng);
      }
    }
    await context.sync();

    const texts = ranges
      .map((range) => normalizeLabel(range.text))
      .filter((text) => text.length > 0);

    if (texts.length === 0) {
      result.reason = "No classification content control found (body or primary headers).";
      return result;
    }

    const mismatches = texts.filter((text) => text !== expectedLabel);
    if (mismatches.length > 0) {
      result.foundText = mismatches[0];
      result.reason = `Found ${texts.length} CC(s); ${mismatches.length} mismatch(es).`;
      return result;
    }

    result.ok = true;
    result.foundText = expectedLabel;
    return result;
  }

  async function verifyClassificationSet(expectedLabel) {
    let result = { ok: false, foundText: "", reason: "Verification did not run." };

    await Word.run(async (context) => {
      result = await verifyClassificationSetInternal(context, normalizeLabel(expectedLabel));
    });

    return result;
  }

  async function hasClassification() {
    let exists = false;

    await Word.run(async (context) => {
      const headerBodies = await getHeaderBodies(context);
      const bodiesToCheck = [context.document.body, ...headerBodies];

      const collections = bodiesToCheck.map((body) => body.contentControls.getByTag(CC_TAG));
      collections.forEach((coll) => coll.load("items"));
      await context.sync();

      exists = collections.some((coll) => coll.items && coll.items.length > 0);
    });

    return exists;
  }

  async function apply(label) {
    const normalizedLabel = normalizeLabel(label);

    if (!normalizedLabel) {
      throw new Error("Classification label is empty.");
    }

    await Word.run(async (context) => {
      await cleanupOrphanLabels(context);
      await upsertClassificationIntoAllPrimaryHeaders(context, normalizedLabel);

      const verify = await verifyClassificationSetInternal(context, normalizedLabel);
      if (!verify.ok) {
        throw new Error(
          `Could not verify classification "${normalizedLabel}".` +
            (verify.foundText ? ` Found: "${verify.foundText}".` : "") +
            (verify.reason ? ` Reason: ${verify.reason}` : "")
        );
      }
    });

    return true;
  }

  return {
    apply,
    hasClassification,
    verifyClassificationSet,
    constants: {
      CC_TAG,
      CC_TITLE,
      HEADER_TYPE
    }
  };
})();