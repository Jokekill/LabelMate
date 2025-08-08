// ===== LabelMate – client-only (bez internetu) =====

const CC_TAG = "LABELMATE_CLASSIFICATION";
const CC_TITLE = "Document Classification";

// všechny možné texty štítků (kvůli úklidu starých odstavců)
const ALL_LABEL_TEXTS = [
  "TLP:Internal","TLP:Protected","TLP:StrictlyProtected",
  "TLP:Interní","TLP:Chráněný","TLP:PřísněChráněný",
  "TLP:Interné","TLP:Chránené","TLP:PrísneChránené"
];

// LocalStorage klíč pro ruční volbu jazyka
const LS_KEY = "labelmate_lang_override";

function getSavedLangOverride() {
  try { return localStorage.getItem(LS_KEY) || "AUTO"; } catch { return "AUTO"; }
}
function saveLangOverride(v) {
  try { localStorage.setItem(LS_KEY, v); } catch {}
}

// Detekce jazyka pouze z Office kontextu (contentLanguage / displayLanguage)
function langFromOfficeContext() {
  // Např. "cs-CZ;en-US" nebo "en-US"
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

  const limit = Math.min(paras.items.length, 25); // koukneme na prvních pár odstavců
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
      p.delete(); // smazat starý volný odstavec se štítkem
    }
  }
  await context.sync();
}

async function applyClassification(label) {
  const statusEl = document.getElementById("status");
  statusEl.style.color = "green";
  statusEl.textContent = "";

  try {
    await Word.run(async (context) => {
      // 1) je tam už náš CC? -> jen vyměň text uvnitř
      const found = context.document.contentControls.getByTag(CC_TAG);
      found.load("items");
      await context.sync();

      if (found.items.length > 0) {
        const cc = found.items[0];
        // replace obsahu, žádné nové odstavce
        cc.insertText(label, Word.InsertLocation.replace);
        cc.cannotEdit = true;
        cc.cannotDelete = true;
        cc.appearance = "BoundingBox";
        cc.color = "#ff0000";

        // pro jistotu i zvětšit/tučně – přes rozsah CC
        const range = cc.getRange();
        range.font.bold = true;
        range.font.size = 14;

        await context.sync();
      } else {
        // 2) uklidit staré sirotčí odstavce z dřívějška
        await cleanupOrphanLabels(context);

        // 3) vložit odstavec úplně na začátek a obalit CC
        const p = context.document.body.insertParagraph(label, Word.InsertLocation.start);

        const cc = p.insertContentControl();
        cc.tag = CC_TAG;
        cc.title = CC_TITLE;
        cc.cannotEdit = true;
        cc.cannotDelete = true;
        cc.appearance = "BoundingBox";
        cc.color = "#ff0000";

        // styling textu uvnitř CC
        p.font.bold = true;
        p.font.size = 14;

        await context.sync();
      }
    });

    statusEl.textContent = `Klasifikace „${label}” byla úspěšně vložena.`;
  } catch (error) {
    console.error(error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", JSON.stringify(error.debugInfo));
    }
    statusEl.style.color = "crimson";
    statusEl.textContent = "Nastala chyba při aplikaci klasifikace.";
  }
}

// Inicializace UI a jazykové logiky
function initLanguageUI() {
  const select = document.getElementById("langSelect");
  const status = document.getElementById("langStatus");

  // výchozí hodnota (AUTO/EN/CZ/SK)
  select.value = getSavedLangOverride();

  const effective = (select.value === "AUTO") ? langFromOfficeContext() : select.value;
  status.textContent = (select.value === "AUTO") ? `Auto: ${effective}` : `Manual: ${effective}`;
  renderLabels(effective);

  // změna ruční volby
  select.addEventListener("change", () => {
    const val = select.value;
    saveLangOverride(val);

    const lang = (val === "AUTO") ? langFromOfficeContext() : val;
    status.textContent = (val === "AUTO") ? `Auto: ${lang}` : `Manual: ${lang}`;
    renderLabels(lang);
  });
}

// Bootstrap – až když je host připraven
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initLanguageUI();
  }
});
