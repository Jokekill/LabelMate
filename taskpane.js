// ===== LabelMate – client-only (žádný internet) =====

// Přednastavené štítky
const LABEL_SETS = {
  EN: ["TLP:Internal", "TLP:Protected", "TLP:StrictlyProtected"],
  CZ: ["TLP:Interní", "TLP:Chráněný", "TLP:PřísněChráněný"],
  SK: ["TLP:Interné", "TLP:Chránené", "TLP:PrísneChránené"],
};

// Persist ruční volby (lokálně)
const LS_KEY = "labelmate_lang_override";

function getSavedLangOverride() {
  try { return localStorage.getItem(LS_KEY) || "AUTO"; } catch { return "AUTO"; }
}
function saveLangOverride(v) {
  try { localStorage.setItem(LS_KEY, v); } catch {}
}

// Detekce jazyka POUZE z Office (bez knihoven, bez síťových volání)
function langFromOfficeContext() {
  // Např. "cs-CZ;en-US" nebo "en-US"
  const content = (Office && Office.context && Office.context.contentLanguage) || "";
  const display = (Office && Office.context && Office.context.displayLanguage) || "";
  const combined = (content ? content : display).toLowerCase();

  if (combined.includes("cs")) return "CZ";
  if (combined.includes("sk")) return "SK";
  if (combined.includes("en")) return "EN";

  // Když Office nic nedá, padáme na EN
  return "EN";
}

// Vyrenderuj 3 tlačítka podle jazyka
function renderLabels(langCode) {
  const container = document.getElementById("labelsContainer");
  container.innerHTML = "";

  const labels = LABEL_SETS[langCode] || LABEL_SETS.EN;
  labels.forEach(text => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = text;
    btn.addEventListener("click", () => applyLabelToDocument(text));
    container.appendChild(btn);
  });
}

// Jednoduché vložení štítku do dokumentu (uprav podle své logiky)
async function applyLabelToDocument(labelText) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph(labelText, Word.InsertLocation.start);
    await context.sync();
  });
}

// Inicializace UI
async function initLanguageUI() {
  const select = document.getElementById("langSelect");
  const status = document.getElementById("langStatus");

  // Načti ruční volbu
  select.value = getSavedLangOverride();

  const effective = (select.value === "AUTO")
    ? langFromOfficeContext()
    : select.value;

  status.textContent = (select.value === "AUTO")
    ? `Auto: ${effective}`
    : `Manual: ${effective}`;

  renderLabels(effective);

  // Handler změny ruční volby
  select.addEventListener("change", () => {
    const val = select.value;
    saveLangOverride(val);

    const lang = (val === "AUTO") ? langFromOfficeContext() : val;
    status.textContent = (val === "AUTO") ? `Auto: ${lang}` : `Manual: ${lang}`;
    renderLabels(lang);
  });
}

// Office bootstrap
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // UI až když je host připravený
    initLanguageUI();
  }
});
