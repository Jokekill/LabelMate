// ===== labels.js – práce s Word CC, render tlačítek, ověření =====

window.Labels = window.Labels || {};

(function () {
  const CC_TAG = "LABELMATE_CLASSIFICATION";
  const CC_TITLE = "Document Classification";

  function getAllLabelTexts() {
    const all = new Set();
    Object.values(window.LM_I18N).forEach(lang => {
      lang.labels.forEach(l => all.add(l.text));
    });
    return Array.from(all);
  }

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

  let running = false;

  async function applyClassification(label) {
    if (running) return;
    running = true;
    window.LM.ui.clearStatus();
    window.LM.ui.setBusy(true);

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
          (L.statusErrVerify(label, "", "Verification failed")) +
          `\n${postErr?.message || postErr}`
        );
      } finally {
        window.Banner.refresh().catch(()=>{});
        window.LM.ui.setBusy(false);
        running = false;
      }
    }
  }

function renderButtons() {
  const container = document.getElementById("labelsContainer");
  if (!container) return;
  container.innerHTML = "";

  const L = window.LM?.i18n?.T?.() || { labels: [], docLinkText: "More in documentation" };

  // drobný helper pro bezpečné vložení plain textu
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, m =>
    ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[m])
  );

  L.labels.forEach(item => {
    const row = document.createElement("div");
    row.className = "row label-row";
    row.style.alignItems = "center";
    row.style.gap = "12px";

    // levý blok: tlačítko + „i“ v rohu
    const wrap = document.createElement("div");
    wrap.className = "classify-wrap";

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn primary pill classify-btn";
    btn.textContent = item.text;
    btn.addEventListener("click", () => window.Labels.applyClassification(item.text));

    // info ikonka (v rohu tlačítka) + tooltip NAD středem
    const infoWrap = document.createElement("div");
    infoWrap.className = "has-tooltip info-in-btn";

    const infoBtn = document.createElement("button");
    infoBtn.type = "button";
    infoBtn.className = "btn icon";
    infoBtn.setAttribute("aria-label", "Info");
    infoBtn.innerHTML = `<span aria-hidden="true">ℹ️</span>`;
    infoBtn.tabIndex = 0;

    const bubble = document.createElement("div");
    bubble.className = "tooltip-bubble";
    bubble.setAttribute("role", "tooltip");

    // 1) Preferuj item.helpHtml (lokalizované, může obsahovat <a>)
    // 2) Fallback: tip + odkaz s lokalizovaným textem
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

    infoWrap.appendChild(infoBtn);
    infoWrap.appendChild(bubble);

    wrap.appendChild(btn);
    wrap.appendChild(infoWrap);

    // pravý popisek (volitelné)
    const right = document.createElement("div");
    right.style.flex = "1 1 auto";
    right.innerHTML = `
      <div style="font-weight:700">${esc(item.text)}</div>
      <div class="muted">${esc(item.tip || "")}</div>
    `;

    row.appendChild(wrap);
    row.appendChild(right);
    container.appendChild(row);
  });
}

  window.Labels = { renderButtons, applyClassification };
})();
