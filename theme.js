// ===== theme.js – ThemeManager (Auto/Light/Dark) s pojistkami proti „přilepenému“ dark =====

window.Theme = window.Theme || {};

(function () {
  const LS_THEME = "labelmate_theme"; // 'auto' | 'light' | 'dark'
  let mql = null;
  let observer = null;

  function getMode() {
    try { return localStorage.getItem(LS_THEME) || "auto"; } catch { return "auto"; }
  }
  function setMode(v) {
    try { localStorage.setItem(LS_THEME, v); } catch {}
  }

  function hardLightGuard() {
    // Některé knihovny (nebo samotný host) umí přilepit 'dark' zpět – po přepnutí na Light
    // na pár sekund hlídáme <html> a <body> a třídu případně odstraníme.
    const html = document.documentElement;
    const body = document.body;

    const applyLight = () => {
      html.classList.remove('dark');
      html.setAttribute('data-theme', 'light');
      if (body && body.classList) body.classList.remove('dark');
      if (body && body.removeAttribute) body.removeAttribute('data-theme');
    };

    // okamžitě + krátké časové okno
    applyLight();
    let ticks = 0;
    const iv = setInterval(() => {
      ticks++;
      applyLight();
      if (ticks > 20) clearInterval(iv); // ~2s
    }, 100);

    // a ještě sledování změn class na <html> po delší dobu
    try {
      if (observer) observer.disconnect();
      observer = new MutationObserver(() => {
        if (getMode() === 'light' && html.classList.contains('dark')) {
          html.classList.remove('dark');
          html.setAttribute('data-theme', 'light');
        }
      });
      observer.observe(html, { attributes: true, attributeFilter: ['class'] });
      // za 10s to vypneme (už to obvykle nikdo nepřepíše)
      setTimeout(() => { try { observer.disconnect(); } catch {} }, 10000);
    } catch {}
  }

  function resolveAndApply(mode) {
    const html = document.documentElement;

    // odpoj starý posluchač systému
    if (mql) {
      try { mql.removeEventListener('change', onSystemChange); } catch {
        try { mql.removeListener(onSystemChange); } catch {}
      }
      mql = null;
    }

    let isDark = false;

    if (mode === 'dark') {
      isDark = true;
    } else if (mode === 'light') {
      isDark = false;
    } else {
      mql = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)');
      isDark = !!(mql && mql.matches);
      if (mql) {
        try { mql.addEventListener('change', onSystemChange); }
        catch { try { mql.addListener(onSystemChange); } catch {} }
      }
    }

    // aplikace
    html.classList.toggle('dark', isDark);
    html.setAttribute('data-theme', isDark ? 'dark' : 'light');

    // jistota: body odtemnit, pokud nejsme v dark
    if (!isDark && document.body) {
      document.body.classList && document.body.classList.remove('dark');
      document.body.removeAttribute && document.body.removeAttribute('data-theme');
      // „pojistka“ proti opětovnému přilepení
      hardLightGuard();
    }
  }

  function onSystemChange() {
    if (getMode() === 'auto') resolveAndApply('auto');
  }

  function initUI() {
    const sel = document.getElementById('themeSelect');
    const stored = getMode();
    if (sel) sel.value = stored;
    resolveAndApply(stored);

    if (sel) {
      sel.addEventListener('change', () => {
        const v = sel.value;
        setMode(v);
        resolveAndApply(v);
      });
    }
  }

  window.Theme = { getMode, setMode, resolveAndApply, initUI };
})();
