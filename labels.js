// ===== labels.js – Word content controls + UI rendering (header-first) =====
//
// Changes in this version:
// - Removes the right-side descriptive column (your request).
// - Tooltip behavior:
//   - Hover/focus on the ℹ️ icon shows the tooltip (preview).
//   - Click/tap on the ℹ️ icon toggles a "locked open" tooltip.
//   - Escape or click outside closes the locked tooltip.
// - Classification insertion is now HEADER-first (Primary header of every section).
//   - Also cleans up any old body-inserted classification CC and orphan paragraphs.

window.Labels = window.Labels || {};

(function () {
  const CC_TAG = "LABELMATE_CLASSIFICATION";
  const CC_TITLE = "Document Classification";
  const HEADER_TYPE = "Primary"; // Word.HeaderFooterType.primary also works

  // ---------------------------
  // Tooltip state (one-at-a-time)
  // ---------------------------
  let openTooltip = null; // { wrap: HTMLElement, btn: HTMLButtonElement } | null
  let tooltipGlobalsBound = false;

  function bindTooltipGlobalsOnce() {
    if (tooltipGlobalsBound) return;
    tooltipGlobalsBound = true;

    // Close when clicking outside the open tooltip's wrap
    document.addEventListener(
      "click",
      (ev) => {
        if (!openTooltip) return;
        const t = ev.target;
        if (t && openTooltip.wrap.contains(t)) return;
        closeTooltip();
      },
      true
    );

    // Close on Escape
    document.addEventListener("keydown", (ev) => {
      if (!openTooltip) return;
      if (ev.key === "Escape" || ev.key === "Esc") {
        ev.preventDefault();
        closeTooltip();
      }
    });
  }

  function closeTooltip() {
    if (!openTooltip) return;
    try {
      openTooltip.wrap.classList.remove("tooltip-open");
      openTooltip.btn.setAttribute("aria-expanded", "false");
    } catch (_) {}
    openTooltip = null;
  }

  function toggleTooltip(wrap, btn) {
    bindTooltipGlobalsOnce();

    const isOpen = wrap.classList.contains("tooltip-open");
    if (isOpen) {
      closeTooltip();
      return;
    }

    // Close any previously open tooltip
    closeTooltip();

    wrap.classList.add("tooltip-open");
    btn.setAttribute("aria-expanded", "true");
    openTooltip = { wrap, btn };
  }

  // ---------------------------
  // Helpers: label text set
  // ---------------------------
  function getAllLabelTexts() {
    const all = new Set();
    Object.values(window.LM_I18N).forEach((lang) => {
      lang.labels.forEach((l) => all.add(l.text));
    });
    return Array.from(all);
  }

  // Remove old orphan paragraphs that exactly match one of our label strings but
  // are not wrapped in a content control. This mainly targets legacy behavior
  // where the label was inserted into document body as plain text.
  async function cleanupOrphanLabels(context) {
    const paras = context.document.body.paragraphs;
    paras.load("items");
    await context.sync();

    // Safer than the old 25: real docs are less likely to start with label text by coincidence.
    const limit = Math.min(paras.items.length, 8);
    const ALL = getAllLabelTexts();

    const checks = [];
    for (let i = 0; i < limit; i++) {
      const p = paras.items[i];
      p.load("text");
      const firstCC = p.contentControls.getFirstOrNullObject();
      firstCC.load("isNullObject");
      checks.push({ p, firstCC });
    }
    await context.sync();

    for (const item of checks) {
      const txt = (item.p.text || "").trim();
      const hasAnyCC = !item.firstCC.isNullObject;
      if (!hasAnyCC && ALL.includes(txt)) {
        item.p.delete();
      }
    }
    await context.sync();
  }

  // ---------------------------
  // Word insertion: header-first
  // ---------------------------
  function applyClassificationStylingToRange(range) {
    // Keep it visible but not obnoxious.
    range.font.bold = true;
    range.font.size = 14;
    range.font.color = "#b91c1c"; // dark red
  }

  function lockAndTagContentControl(cc) {
    cc.tag = CC_TAG;
    cc.title = CC_TITLE;
    cc.cannotEdit = true;
    cc.cannotDelete = true;

    // Cleaner than BoundingBox in a header; keeps content control functional/locked.
    // Appearance values can be "BoundingBox", "Tags", or "Hidden".
    cc.appearance = "Hidden";
    cc.color = "#b91c1c";
  }

  async function upsertClassificationIntoAllPrimaryHeaders(context, label) {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    // If sections aren't available for some reason, fall back to body.
    if (!sections.items || sections.items.length === 0) {
      const p = context.document.body.insertParagraph(label, "Start");
      const cc = p.insertContentControl();
      lockAndTagContentControl(cc);

      const rng = cc.getRange("Content");
      rng.insertText(label, Word.InsertLocation.replace);
      applyClassificationStylingToRange(rng);

      await context.sync();
      return;
    }

    // Insert/update in each section's Primary header.
    for (const section of sections.items) {
      const headerBody = section.getHeader(HEADER_TYPE);

      const headerCCs = headerBody.contentControls.getByTag(CC_TAG);
      headerCCs.load("items");
      await context.sync();

      if (headerCCs.items && headerCCs.items.length > 0) {
        // Update all existing header CCs (dedupe-friendly)
        for (let i = 0; i < headerCCs.items.length; i++) {
          const cc = headerCCs.items[i];

          // Temporarily unlock to edit, then relock.
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
        // Create a new paragraph at the start of the header and wrap it.
        const p = headerBody.insertParagraph(label, "Start");
        const cc = p.insertContentControl();
        lockAndTagContentControl(cc);

        const rng = cc.getRange("Content");
        rng.load("text"); // just to keep the object warm/loaded for subsequent ops
        applyClassificationStylingToRange(rng);

        await context.sync();
      }
    }

    // Remove any old BODY content controls with the same tag so the doc doesn't show two labels.
    // Word.ContentControl.delete(keepContent):
    // - keepContent=true  -> remove control but keep text
    // - keepContent=false -> remove both control and text
    // We'll remove both from body.
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

  // ---------------------------
  // Verification (header-aware)
  // ---------------------------
  async function verifyClassificationSet(expectedLabel) {
    const result = { ok: false, foundText: "", reason: "" };

    await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      const bodiesToCheck = [context.document.body];
      if (sections.items && sections.items.length > 0) {
        for (const s of sections.items) bodiesToCheck.push(s.getHeader(HEADER_TYPE));
      }

      const collections = bodiesToCheck.map((b) => b.contentControls.getByTag(CC_TAG));
      collections.forEach((c) => c.load("items"));
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

      const texts = ranges.map((r) => (r.text || "").trim()).filter((t) => t.length > 0);
      if (texts.length === 0) {
        result.reason = "No classification content control found (body or primary headers).";
        return;
      }

      // If any differs, call it a mismatch.
      const mismatches = texts.filter((t) => t !== expectedLabel);
      if (mismatches.length > 0) {
        result.foundText = mismatches[0];
        result.reason = `Found ${texts.length} CC(s); ${mismatches.length} mismatch(es).`;
        result.ok = false;
        return;
      }

      result.ok = true;
      result.foundText = expectedLabel;
    });

    return result;
  }

  // ---------------------------
  // Public action: apply classification
  // ---------------------------
  let running = false;

  async function applyClassification(label) {
    if (running) return;
    running = true;

    closeTooltip();
    window.LM.ui.clearStatus();
    window.LM.ui.setBusy(true);

    let caughtError = null;

    try {
      await Word.run(async (context) => {
        // Clean legacy body labels first.
        await cleanupOrphanLabels(context);

        // Header-first insert/update.
        await upsertClassificationIntoAllPrimaryHeaders(context, label);
      });
    } catch (err) {
      console.error(err);
      caughtError = err;
    } finally {
      try {
        const verify = await verifyClassificationSet(label);
        const L = window.LM.i18n.T();
        if (verify.ok) {
          window.LM.ui.setStatusOk(L.statusOk(label));
        } else {
          window.LM.ui.setStatusError(
            L.statusErrVerify(label, verify.foundText, verify.reason) +
              (caughtError?.message ? `\n${caughtError.message}` : "")
          );
        }
      } catch (postErr) {
        const L = window.LM.i18n.T();
        window.LM.ui.setStatusError(
          L.statusErrVerify(label, "", "Verification failed") + `\n${postErr?.message || postErr}`
        );
      } finally {
        window.Banner.refresh().catch(() => {});
        window.LM.ui.setBusy(false);
        running = false;
      }
    }
  }

  // ---------------------------
  // UI rendering (no right column)
  // ---------------------------
  function renderButtons() {
    // If re-render happens while a tooltip is open, close it.
    closeTooltip();

    const container = document.getElementById("labelsContainer");
    if (!container) return;
    container.innerHTML = "";

    const L = window.LM?.i18n?.T?.() || { labels: [], docLinkText: "More in documentation" };

    // Safe HTML escaping for plain-text fields.
    const esc = (s) =>
      String(s ?? "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));

    L.labels.forEach((item, idx) => {
      const row = document.createElement("div");
      row.className = "label-row";

      const wrap = document.createElement("div");
      wrap.className = "classify-wrap";

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn primary pill classify-btn";
      btn.textContent = item.text;
      btn.addEventListener("click", () => {
        closeTooltip();
        window.Labels.applyClassification(item.text);
      });

      // Info icon
      const infoWrap = document.createElement("div");
      infoWrap.className = "info-in-btn";

      const infoBtn = document.createElement("button");
      infoBtn.type = "button";
      infoBtn.className = "btn icon info-btn";
      
      infoBtn.setAttribute("aria-label", "More information");
      infoBtn.setAttribute("aria-expanded", "false");

      const tooltipId = `lm-tooltip-${idx}`;
      infoBtn.setAttribute("aria-controls", tooltipId);

      infoBtn.innerHTML = `<span aria-hidden="true">ℹ️</span>`;
      infoBtn.tabIndex = 0;

      // Tooltip bubble
      const bubble = document.createElement("div");
      bubble.id = tooltipId;
      bubble.className = "tooltip-bubble";
      bubble.setAttribute("role", "tooltip");

      let inner;
      if (item.helpHtml) {
        inner = item.helpHtml;
      } else {
        const tip = esc(item.tip || "");
        const link = item.docUrl
          ? ` <a href="${item.docUrl}" target="_blank" rel="noopener">${esc(L.docLinkText)}</a>.`
          : "";
        inner = `${tip}${link}`;
      }
      bubble.innerHTML = `<strong>${esc(item.text)}</strong><br>${inner}`;

      // Click-to-lock
      infoBtn.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        toggleTooltip(wrap, infoBtn);
      });

      // Hover/focus preview behaviour is handled purely via CSS.
      // If the tooltip is "locked open" and user tabs away, click-outside/Escape closes it.

      infoWrap.appendChild(infoBtn);

      // DOM order matters for CSS sibling selectors:
      // infoWrap THEN bubble so we can use .info-in-btn:hover + .tooltip-bubble
      wrap.appendChild(btn);
      wrap.appendChild(infoWrap);
      wrap.appendChild(bubble);

      row.appendChild(wrap);
      container.appendChild(row);
    });

    // Ensure global handlers exist after first render.
    bindTooltipGlobalsOnce();
  }

  window.Labels = { renderButtons, applyClassification };
})();
