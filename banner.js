// ===== banner.js – logika banneru „dokument nemá klasifikaci“ =====

window.Banner = window.Banner || {};

(function () {
  const CC_TAG = "LABELMATE_CLASSIFICATION";

  function show() {
    const el = document.getElementById("missingBanner");
    if (el) el.classList.remove("hidden");
    const tr = window.LM.i18n.T();
    const t = document.getElementById("bnrTitle");
    const d = document.getElementById("bnrDesc");
    if (t && d && tr) { t.textContent = tr.banner.title; d.textContent = tr.banner.desc; }
  }
  function hide() {
    const el = document.getElementById("missingBanner");
    if (el) el.classList.add("hidden");
  }

  async function hasClassificationCC() {
    let exists = false;
    await Word.run(async (context) => {
      const found = context.document.contentControls.getByTag(CC_TAG);
      found.load("items");
      await context.sync();
      if (found.items && found.items.length > 0 && !found.items[0].isNullObject) {
        const rng = found.items[0].getRange("Content");
        rng.load("text");
        await context.sync();
        if ((rng.text || "").trim().length > 0) exists = true;
      }
    });
    return exists;
  }

  async function refresh() {
    try { (await hasClassificationCC()) ? hide() : show(); }
    catch { show(); }
  }

  let started = false;
  function startHeartbeat() {
    if (started) return;
    started = true;
    setInterval(() => { refresh().catch(()=>{}); }, 4000);
  }

  window.Banner = { refresh, startHeartbeat, show, hide };
})();
