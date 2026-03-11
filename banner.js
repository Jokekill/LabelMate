// ===== banner.js – banner logic: "document has no classification" =====
//
// Change in this version:
// - Header-aware detection: checks both document body and Primary headers
//   for the classification content control.

window.Banner = window.Banner || {};

(function () {
  const CC_TAG = "LABELMATE_CLASSIFICATION";
  const HEADER_TYPE = "Primary";

  function show() {
    const el = document.getElementById("missingBanner");
    if (el) el.classList.remove("hidden");
    const tr = window.LM.i18n.T();
    const t = document.getElementById("bnrTitle");
    const d = document.getElementById("bnrDesc");
    if (t && d && tr) {
      t.textContent = tr.banner.title;
      d.textContent = tr.banner.desc;
    }
  }

  function hide() {
    const el = document.getElementById("missingBanner");
    if (el) el.classList.add("hidden");
  }

  async function hasClassificationCC() {
    let exists = false;

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

      exists = ranges.some((r) => ((r.text || "").trim().length > 0));
    });

    return exists;
  }

  async function refresh() {
    try {
      (await hasClassificationCC()) ? hide() : show();
    } catch {
      // If we can't verify (rare runtime glitches), fail safe by showing the banner.
      show();
    }
  }

  let started = false;
  function startHeartbeat() {
    if (started) return;
    started = true;

    // Keep as interval for now; could later be replaced by Word events (preview APIs).
    setInterval(() => {
      refresh().catch(() => {});
    }, 4000);
  }

  window.Banner = { refresh, startHeartbeat, show, hide };
})();
