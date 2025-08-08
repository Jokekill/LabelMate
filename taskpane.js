// ===== LabelMate – client-only (bez internetu) =====

// Přednastavené štítky pro 3 jazyky
const LABEL_SETS = {
  EN: ["TLP:Internal", "TLP:Protected", "TLP:StrictlyProtected"],
  CZ: ["TLP:Interní", "TLP:Chráněný", "TLP:PřísněChráněný"],
  SK: ["TLP:Interné", "TLP:Chránené", "TLP:PrísneChránené"],
};

// Tag a název našeho content controlu s klasifikací
const CC_TAG = "LABELMATE_CLASSIFICATION";
const CC_TITLE = "Document Classification";

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

// Aplikace klasifikace do needitovatelného Content Control na začátek dokumentu
async function applyClassification(label) {
  const statusEl = document.getElementById("status");
  statusEl.style.color = "green";
  statusEl.textContent = "";

  try {
    await Word.run(async (context) => {
      // 1) Zruš starou klasifikaci (ať je vždy jen jedna)
      await removeExistingClassificationCC(context);

      // 2) Vlož odstavec s textem na samý začátek dokumentu
      const p = context.document.body.insertParagraph(label, Word.InsertLocation.start);

      // 3) Obal odstavec do Content Control
      const cc = p.insertContentControl();
      cc.tag = CC_TAG;
      cc.title = CC_TITLE;
      cc.color = "#ff0000";               // volitelné zvýraznění rámečku
      cc.appearance = "BoundingBox";      // "BoundingBox" | "Tags" | "Hidden"
      cc.cannotEdit = true;               // zamkne editaci obsahu
      cc.removeWhenEdited = false;        // pokus o editaci nespustí odstranění

      // (Volitelná úprava vzhledu textu)
      p.font.bold = true;
      p.font.size = 14;

      await context.sync();
    });

    statusEl.textContent = `Klasifikace „${label}” byla úspěšně vložena.`;
  } catch (error) {
    console.error("Error:", error);
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
