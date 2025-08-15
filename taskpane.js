// ===== LabelMate – client-only (bez internetu) =====

const CC_TAG = "LABELMATE_CLASSIFICATION";
const CC_TITLE = "Document Classification";

// Přednastavené štítky pro 3 jazyky
const LABEL_SETS = {
  EN: ["TLP:Internal", "TLP:Protected", "TLP:StrictlyProtected"],
  CZ: ["TLP:Interní", "TLP:Chráněný", "TLP:PřísněChráněný"],
  SK: ["TLP:Interné", "TLP:Chránené", "TLP:PrísneChránené"],
};

// všechny možné texty štítků (kvůli úklidu starých odstavců)
const ALL_LABEL_TEXTS = [...LABEL_SETS.EN, ...LABEL_SETS.CZ, ...LABEL_SETS.SK];

// LocalStorage klíč pro ruční volbu jazyka
const LS_KEY = "labelmate_lang_override";
function getSavedLangOverride() {
  try { return localStorage.getItem(LS_KEY) || "AUTO"; } catch { return "AUTO"; }
}
function saveLangOverride(v) {
  try { localStorage.setItem(LS_KEY, v); } catch {}
}

// ===== UI helpery (status) =====
function setStatusOk(msg) {
  const el = document.getElementById("status");
  if (!el) return;
  el.classList.remove("error");
  el.classList.add("ok");
  el.textContent = msg || "";
}
function setStatusError(msg) {
  const el = document.getElementById("status");
  if (!el) return;
  el.classList.remove("ok");
  el.classList.add("error");
  el.textContent = msg || "";
}
function clearStatusProcessing() {
  const el = document.getElementById("status");
  if (!el) return;
  el.classList.remove("ok", "error");
  el.textContent = "";
}
function setBusy(isBusy) {
  document.querySelectorAll("#labelsContainer button").forEach(b => b.disabled = isBusy);
}

// ===== Banner helpery =====
function showMissingBanner() {
  const el = document.getElementById("missingBanner");
  if (el) el.classList.remove("lm-hidden");
}
function hideMissingBanner() {
  const el = document.getElementById("missingBanner");
  if (el) el.classList.add("lm-hidden");
}

/** Zjistí, zda existuje CC s tagem klasifikace */
async function hasClassificationCC() {
  let exists = false;
  await Word.run(async (context) => {
    try {
      const found = context.document.contentControls.getByTag(CC_TAG);
      found.load("items");
      await context.sync();

      if (found.items && found.items.length > 0 && !found.items[0].isNullObject) {
        const rng = found.items[0].getRange("Content");
        rng.load("text");
        await context.sync();
        if ((rng.text || "").trim().length > 0) {
          exists = true;
        }
      }
    } catch (err) {
      console.error("Error in hasClassificationCC:", err);
    }
  });
  console.log("hasClassificationCC result:", exists);
  return exists;
}


/** Zkontroluje dokument a zobrazí/skrýje banner */
async function updateMissingBanner() {
  try {
    const exists = await hasClassificationCC();
    console.log("updateMissingBanner - CC exists:", exists);
    if (exists) {
      hideMissingBanner();
    } else {
      showMissingBanner();
    }
  } catch (e) {
    // Fail-safe: když kontrola selže, raději banner ukázat
    showMissingBanner();
    console.warn("Banner check failed:", e);
  }
}

// (volitelně) jemný pravidelný heartbeat pro případ ručního smazání CC
let _bannerIntervalStarted = false;
function startBannerHeartbeat() {
  if (_bannerIntervalStarted) return;
  _bannerIntervalStarted = true;
  setInterval(() => {
    if (!running) updateMissingBanner().catch(() => {});
  }, 4000);
}

// ===== Jazyk a vykreslení tlačítek =====
function langFromOfficeContext() {
  const content = (Office && Office.context && Office.context.contentLanguage) || "";
  const display = (Office && Office.context && Office.context.displayLanguage) || "";
  const combined = (content || display).toLowerCase();
  if (combined.includes("cs")) return "CZ";
  if (combined.includes("sk")) return "SK";
  if (combined.includes("en")) return "EN";
  return "EN";
}

function renderLabels(langCode) {
  const container = document.getElementById("labelsContainer");
  if (!container) return;
  container.innerHTML = "";

  const labels = LABEL_SETS[langCode] || LABEL_SETS.EN;
  labels.forEach(text => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = text;
    btn.addEventListener("click", () => applyClassification(text));
    container.appendChild(btn);
  });
}

// ===== Úklid „sirotků“ – odstavce s textem štítku bez CC =====
async function cleanupOrphanLabels(context) {
  const paras = context.document.body.paragraphs;
  paras.load("items");
  await context.sync();

  const limit = Math.min(paras.items.length, 25);
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
    if (!hasAnyCC && ALL_LABEL_TEXTS.includes(txt)) {
      item.p.delete();
    }
  }
  await context.sync();
}

// ===== Hlavní akce: vloží / nahradí klasifikaci v jediném CC =====
let running = false;
async function applyClassification(label) {
  if (running) return;
  running = true;
  setBusy(true);
  clearStatusProcessing();

  let caughtError = null;

  try {
    await Word.run(async (context) => {
      const found = context.document.contentControls.getByTag(CC_TAG);
      found.load("items");
      await context.sync();

      if (found.items.length > 0) {
        const cc = found.items[0];
        if (!cc.isNullObject) {
          // dočasně odemknout
          cc.cannotEdit = false;
          cc.cannotDelete = false;
          await context.sync();

          // UPRAVUJEME JEN OBSAH CC (ne celý CC)
          const contentRange = cc.getRange("Content");
          contentRange.insertText(label, Word.InsertLocation.replace);
          contentRange.font.bold = true;
          contentRange.font.size = 14;

          // jistota: tag a title
          cc.tag = CC_TAG;
          cc.title = CC_TITLE;

          // znovu zamknout
          cc.cannotEdit = true;
          cc.cannotDelete = true;
          cc.appearance = "BoundingBox";
          cc.color = "#ff0000";
          await context.sync();
        }
      } else {
        // uklidit staré sirotky
        await cleanupOrphanLabels(context);

        // vložit nový odstavec + CC na začátek
        const p = context.document.body.insertParagraph(label, Word.InsertLocation.start);
        const cc = p.insertContentControl();
        cc.tag = CC_TAG;
        cc.title = CC_TITLE;

        // styl textu
        p.font.bold = true;
        p.font.size = 14;

        // zamknout
        cc.cannotEdit = true;
        cc.cannotDelete = true;
        cc.appearance = "BoundingBox";
        cc.color = "#ff0000";
        await context.sync();
      }
    });
  } catch (err) {
    console.error(err);
    caughtError = err;
  } finally {
    try {
      // Ověření výsledku – status nastavujeme výhradně tady
      const verify = await verifyClassificationSet(label);

      if (verify.ok) {
        setStatusOk(`Klasifikace „${label}” byla úspěšně nastavena.`);
      } else {
        const errDetail = caughtError?.message ? `\nChyba operace: ${caughtError.message}` : "";
        const foundInfo = verify.foundText ? ` Nalezený text: „${verify.foundText}”.` : "";
        setStatusError(
          `Nepodařilo se potvrdit nastavení klasifikace na „${label}”.${foundInfo}${errDetail}${
            verify.reason ? `\nDůvod: ${verify.reason}` : ""
          }`
        );
      }
    } catch (postErr) {
      const errDetail = caughtError?.message ? `\nChyba operace: ${caughtError.message}` : "";
      setStatusError(`Nepodařilo se ověřit výsledek klasifikace.${errDetail}\nVerifikační chyba: ${postErr?.message || postErr}`);
    } finally {
      // Po každém pokusu překreslit banner (a uvolnit UI)
      updateMissingBanner().catch(console.warn);
      setBusy(false);
      running = false;
    }
  }
}

// ===== Verifikace obsahu CC =====
async function verifyClassificationSet(expectedLabel) {
  let result = { ok: false, foundText: "", reason: "" };
  await Word.run(async (context) => {
    const found = context.document.contentControls.getByTag(CC_TAG);
    found.load("items");
    await context.sync();

    if (!found.items || found.items.length === 0) {
      result.reason = "Nebyl nalezen žádný Content Control s daným tagem.";
      return;
    }

    const cc = found.items[0];
    const rng = cc.getRange("Content");
    rng.load("text");
    await context.sync();

    const txt = (rng.text || "").trim();
    result.foundText = txt;
    result.ok = (txt === expectedLabel);
    if (!result.ok) {
      result.reason = "Text v CC neodpovídá očekávanému štítku.";
    }
  });
  return result;
}

// ===== Inicializace UI a jazykové logiky =====
function initLanguageUI() {
  const select = document.getElementById("langSelect");
  const status = document.getElementById("langStatus");
  if (!select) return;

  select.value = getSavedLangOverride();
  const effective = (select.value === "AUTO") ? langFromOfficeContext() : select.value;
  if (status) status.textContent = (select.value === "AUTO") ? `Auto: ${effective}` : `Manual: ${effective}`;
  renderLabels(effective);

  select.addEventListener("change", () => {
    const val = select.value;
    saveLangOverride(val);
    const lang = (val === "AUTO") ? langFromOfficeContext() : val;
    if (status) status.textContent = (val === "AUTO") ? `Auto: ${lang}` : `Manual: ${lang}`;
    renderLabels(lang);
  });
}

// ===== Bootstrap – až když je host připraven =====
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initLanguageUI();
    updateMissingBanner();   // počáteční kontrola
    startBannerHeartbeat();  // volitelné: průběžná kontrola
  }
});
