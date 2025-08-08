// ===== LabelMate – client-only (bez internetu) =====

const CC_TAG = "LABELMATE_CLASSIFICATION";
const CC_TITLE = "Document Classification";

// Přednastavené štítky pro 3 jazyky (DŮLEŽITÉ – chybělo)
const LABEL_SETS = {
  EN: ["TLP:Internal", "TLP:Protected", "TLP:StrictlyProtected"],
  CZ: ["TLP:Interní", "TLP:Chráněný", "TLP:PřísněChráněný"],
  SK: ["TLP:Interné", "TLP:Chránené", "TLP:PrísneChránené"],
};

// všechny možné texty štítků (kvůli úklidu starých odstavců)
const ALL_LABEL_TEXTS = [
  ...LABEL_SETS.EN, ...LABEL_SETS.CZ, ...LABEL_SETS.SK
];

// LocalStorage klíč pro ruční volbu jazyka
const LS_KEY = "labelmate_lang_override";
function getSavedLangOverride() {
  try { return localStorage.getItem(LS_KEY) || "AUTO"; } catch { return "AUTO"; }
}
function saveLangOverride(v) {
  try { localStorage.setItem(LS_KEY, v); } catch {}
}

// Detekce jazyka pouze z Office kontextu
function langFromOfficeContext() {
  const content = (Office && Office.context && Office.context.contentLanguage) || "";
  const display = (Office && Office.context && Office.context.displayLanguage) || "";
  const combined = (content || display).toLowerCase();
  if (combined.includes("cs")) return "CZ";
  if (combined.includes("sk")) return "SK";
  if (combined.includes("en")) return "EN";
  return "EN";
}

// Vykreslí 3 tlačítka dle jazyka
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

// Smaže předchozí klasifikační CC (pokud existuje)
async function removeExistingClassificationCC(context) {
  const existing = context.document.contentControls.getByTag(CC_TAG);
  existing.load("items");
  await context.sync();
  if (existing.items.length > 0) {
    existing.items.forEach(cc => cc.delete(true)); // true = smaže i obsah uvnitř
    await context.sync();
  }
}

// smaže staré „sirotčí“ odstavce se štítky (bez CC) z horní části dokumentu
async function cleanupOrphanLabels(context) {
  const paras = context.document.body.paragraphs;
  paras.load("items");
  await context.sync();

  const limit = Math.min(paras.items.length, 25);
  for (let i = 0; i < limit; i++) {
    const p = paras.items[i];
    p.load(["text", "contentControls"]);
  }
  await context.sync();

  for (let i = 0; i < limit; i++) {
    const p = paras.items[i];
    const txt = (p.text || "").trim();
    const hasCC = p.contentControls.items && p.contentControls.items.length > 0;
    if (!hasCC && ALL_LABEL_TEXTS.includes(txt)) {
      p.delete();
    }
  }
  await context.sync();
}

// Vloží / nahradí klasifikaci v jediném Content Controlu
async function applyClassification(label) {
  const statusEl = document.getElementById("status");
  if (statusEl) { statusEl.style.color = "green"; statusEl.textContent = ""; }

  try {
    await Word.run(async (context) => {
      const found = context.document.contentControls.getByTag(CC_TAG);
      found.load("items");
      await context.sync();

      if (found.items.length > 0) {
        const cc = found.items[0];

        // 1) dočasně odemknout
        cc.cannotEdit = false;
        cc.cannotDelete = false;
        await context.sync();

        // 2) přepsat obsah bezpečně přes range
        const range = cc.getRange();
        range.insertText(label, Word.InsertLocation.replace);
        range.font.bold = true;
        range.font.size = 14;

        // 3) znovu zamknout
        cc.cannotEdit = true;
        cc.cannotDelete = true;
        cc.appearance = "BoundingBox";
        cc.color = "#ff0000";
        await context.sync();
      } else {
        // uklidit staré sirotky
        await cleanupOrphanLabels(context);

        // vložit nový CC na začátek
        const p = context.document.body.insertParagraph(label, Word.InsertLocation.start);
        const cc = p.insertContentControl();
        cc.tag = CC_TAG;
        cc.title = CC_TITLE;

        // styl textu
        p.font.bold = true;
        p.font.size = 14;

        // zamknout až NAKONEC (po vložení)
        cc.cannotEdit = true;
        cc.cannotDelete = true;
        cc.appearance = "BoundingBox";
        cc.color = "#ff0000";
        await context.sync();
      }
    });

    if (statusEl) statusEl.textContent = `Klasifikace „${label}” byla úspěšně vložena.`;
  } catch (error) {
    console.error(error);
    if (statusEl) {
      statusEl.style.color = "crimson";
      statusEl.textContent = "Nastala chyba při aplikaci klasifikace.";
    }
  }
}


// Inicializace UI a jazykové logiky
function initLanguageUI() {
  const select = document.getElementById("langSelect");
  const status = document.getElementById("langStatus");
  if (!select) return; // bezpečnost

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

// Bootstrap – až když je host připraven
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initLanguageUI();
  }
});
