// ===== LabelMate – modern UI (Tailwind) + i18n + dark mode =====

const CC_TAG = "LABELMATE_CLASSIFICATION";
const CC_TITLE = "Document Classification";

const LS_THEME = "labelmate_theme";
const LS_LANG  = "labelmate_lang_override";

// --- i18n helpers ---
function getSavedLangOverride() {
  try { return localStorage.getItem(LS_LANG) || "AUTO"; } catch { return "AUTO"; }
}
function saveLangOverride(v) {
  try { localStorage.setItem(LS_LANG, v); } catch {}
}
function langFromOfficeContext() {
  const content = (Office && Office.context && Office.context.contentLanguage) || "";
  const display = (Office && Office.context && Office.context.displayLanguage) || "";
  const combined = (content || display).toLowerCase();
  if (combined.includes("cs")) return "CZ";
  if (combined.includes("sk")) return "SK";
  if (combined.includes("en")) return "EN";
  return window.LM_DEFAULT_LANG || "EN";
}
function getCurrentLangCode() {
  const sel = document.getElementById("langSelect");
  const v = sel?.value || "AUTO";
  return (v === "AUTO") ? langFromOfficeContext() : v;
}
function T() {
  const code = getCurrentLangCode();
  return window.LM_I18N[code] || window.LM_I18N[window.LM_DEFAULT_LANG];
}

// --- status helpers (Tailwind colors via classes) ---
function setStatusOk(msg) {
  const el = document.getElementById("status");
  if (!el) return;
  el.classList.remove("text-rose-600","dark:text-rose-400");
  el.classList.add("text-emerald-600","dark:text-emerald-400");
  el.textContent = msg || "";
}
function setStatusError(msg) {
  const el = document.getElementById("status");
  if (!el) return;
  el.classList.remove("text-emerald-600","dark:text-emerald-400");
  el.classList.add("text-rose-600","dark:text-rose-400");
  el.textContent = msg || "";
}
function clearStatusProcessing() {
  const el = document.getElementById("status");
  if (!el) return;
  el.textContent = "";
}
function setBusy(isBusy) {
  document.querySelectorAll("#labelsContainer button").forEach(b => b.disabled = isBusy);
}

// --- banner helpers ---
function showMissingBanner() {
  const el = document.getElementById("missingBanner");
  if (el) el.classList.remove("hidden");
  const t = document.getElementById("bnrTitle");
  const d = document.getElementById("bnrDesc");
  const tr = T();
  if (t && d && tr) {
    t.textContent = tr.banner.title;
    d.textContent = tr.banner.desc;
  }
}
function hideMissingBanner() {
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
async function updateMissingBanner() {
  try {
    const exists = await hasClassificationCC();
    if (exists) hideMissingBanner(); else showMissingBanner();
  } catch {
    showMissingBanner();
  }
}
let running = false;
let _bannerIntervalStarted = false;
function startBannerHeartbeat() {
  if (_bannerIntervalStarted) return;
  _bannerIntervalStarted = true;
  setInterval(() => {
    if (!running) updateMissingBanner().catch(()=>{});
  }, 4000);
}

// --- render classification buttons (with tooltip + doc link) ---
function renderLabels() {
  const container = document.getElementById("labelsContainer");
  if (!container) return;
  container.innerHTML = "";

  const L = T();
  L.labels.forEach(item => {
    const row = document.createElement("div");
    row.className = "flex items-stretch gap-2";

    // Button
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = item.text;
    btn.title = item.tip; // nativní tooltip
    btn.className =
      "flex-1 rounded-md bg-sky-600 hover:bg-sky-700 disabled:opacity-60 " +
      "text-white font-semibold px-4 py-3 text-sm transition";
    btn.addEventListener("click", () => applyClassification(item.text));

    // Doc link (i)
    const a = document.createElement("a");
    a.href = item.docUrl;
    a.target = "_blank";
    a.rel = "noopener";
    a.title = L.choosePrompt + " " + item.text;
    a.className =
      "shrink-0 rounded-md border border-slate-300 dark:border-slate-700 " +
      "bg-white dark:bg-slate-800 px-3 py-3 text-sm hover:bg-slate-50 dark:hover:bg-slate-700 " +
      "flex items-center justify-center";
    a.innerHTML = `<span aria-hidden="true">ℹ️</span>`;

    row.appendChild(btn);
    row.appendChild(a);
    container.appendChild(row);
  });
}

// --- orphan cleanup (stejné jako dřív) ---
async function cleanupOrphanLabels(context) {
  const paras = context.document.body.paragraphs;
  paras.load("items");
  await context.sync();

  const limit = Math.min(paras.items.length, 25);
  const checks = [];
  const ALL = getAllLabelTexts();
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
function getAllLabelTexts() {
  // z i18n posbíráme všechny texty napříč jazyky
  const all = new Set();
  Object.values(window.LM_I18N).forEach(lang => {
    lang.labels.forEach(l => all.add(l.text));
  });
  return Array.from(all);
}

// --- main action: insert/replace classification CC ---
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
          cc.cannotEdit = false;
          cc.cannotDelete = false;
          await context.sync();

          const contentRange = cc.getRange("Content");
          contentRange.insertText(label, Word.InsertLocation.replace);
          contentRange.font.bold = true;
          contentRange.font.size = 14;

          cc.tag = CC_TAG;
          cc.title = CC_TITLE;

          cc.cannotEdit = true;
          cc.cannotDelete = true;
          cc.appearance = "BoundingBox";
          cc.color = "#ff0000";
          await context.sync();
        }
      } else {
        await cleanupOrphanLabels(context);
        const p = context.document.body.insertParagraph(label, Word.InsertLocation.start);
        const cc = p.insertContentControl();
        cc.tag = CC_TAG;
        cc.title = CC_TITLE;
        p.font.bold = true;
        p.font.size = 14;
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
      const verify = await verifyClassificationSet(label);
      const L = T();
      if (verify.ok) {
        setStatusOk(L.statusOk(label));
      } else {
        setStatusError(L.statusErrVerify(label, verify.foundText, verify.reason) +
          (caughtError?.message ? `\n${caughtError.message}` : ""));
      }
    } catch (postErr) {
      const L = T();
      setStatusError((L.statusErrVerify(label, "", "Verification failed")) +
        `\n${postErr?.message || postErr}`);
    } finally {
      updateMissingBanner().catch(console.warn);
      setBusy(false);
      running = false;
    }
  }
}
async function verifyClassificationSet(expectedLabel) {
  let result = { ok: false, foundText: "", reason: "" };
  await Word.run(async (context) => {
    const found = context.document.contentControls.getByTag(CC_TAG);
    found.load("items");
    await context.sync();

    if (!found.items || found.items.length === 0) {
      result.reason = "No CC with the tag.";
      return;
    }
    const cc = found.items[0];
    const rng = cc.getRange("Content");
    rng.load("text");
    await context.sync();

    const txt = (rng.text || "").trim();
    result.foundText = txt;
    result.ok = (txt === expectedLabel);
    if (!result.ok) result.reason = "Content doesn’t match expected label.";
  });
  return result;
}

// --- language + theme UI init ---
function initLanguageUI() {
  const select = document.getElementById("langSelect");
  const status = document.getElementById("langStatus");
  const themeSel = document.getElementById("themeSelect");

  // i18n texty v UI
  function applyStaticTexts() {
    const L = T();
    document.getElementById("appTitle").textContent = L.appTitle;
    document.getElementById("themeLabel").textContent = L.themeLabel;
    document.getElementById("langLabel").textContent = L.langLabel;
    document.getElementById("choosePrompt").textContent = L.choosePrompt;
    // banner texty doplní showMissingBanner()
  }

  // init jazyka
  if (select) {
    select.value = getSavedLangOverride();
    const effective = (select.value === "AUTO") ? langFromOfficeContext() : select.value;
    if (status) status.textContent = (select.value === "AUTO") ? `Auto: ${effective}` : `Manual: ${effective}`;
    applyStaticTexts();
    renderLabels();

    select.addEventListener("change", () => {
      const val = select.value;
      saveLangOverride(val);
      const lang = (val === "AUTO") ? langFromOfficeContext() : val;
      if (status) status.textContent = (val === "AUTO") ? `Auto: ${lang}` : `Manual: ${lang}`;
      applyStaticTexts();
      renderLabels();
      updateMissingBanner().catch(()=>{});
    });
  }

  // init theme
  if (themeSel) {
    const stored = localStorage.getItem(LS_THEME) || 'auto';
    themeSel.value = stored;
    applyTheme(stored);
    themeSel.addEventListener('change', () => {
      const v = themeSel.value;
      localStorage.setItem(LS_THEME, v);
      applyTheme(v);
    });
  }
}
function applyTheme(mode) {
  const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  const root = document.documentElement;
  if (mode === 'dark' || (mode === 'auto' && prefersDark)) {
    root.classList.add('dark');
  } else {
    root.classList.remove('dark');
  }
  root.dataset.theme = mode;
}

// --- bootstrap ---
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initLanguageUI();
    updateMissingBanner().catch(()=>{});
    startBannerHeartbeat();
  }
});
