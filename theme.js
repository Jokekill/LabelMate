// ===== theme.js – Light/Dark only, kompletně self-contained =====
(function () {
  const LS_THEME = "labelmate_theme"; // 'light' | 'dark'
  let observer = null;

  // --- Pre-apply: okamžitě po načtení skriptu, ještě před Tailwindem ---
  (function preapply() {
    let v = 'light';
    try {
      const s = localStorage.getItem(LS_THEME);
      if (s === 'dark' || s === 'light') v = s;
    } catch {}
    const html = document.documentElement;
    html.classList.toggle('dark', v === 'dark');
    html.setAttribute('data-theme', v);
  })();

  // --- Helpers ---
  function getMode() {
    try {
      const v = localStorage.getItem(LS_THEME);
      return (v === 'dark' || v === 'light') ? v : 'light';
    } catch { return 'light'; }
  }
  function setMode(v) {
    try { localStorage.setItem(LS_THEME, (v === 'dark' ? 'dark' : 'light')); } catch {}
  }

  function hardLightGuard() {
    const html = document.documentElement;
    const body = document.body;
    const applyLight = () => {
      html.classList.remove('dark');
      html.setAttribute('data-theme', 'light');
      if (body) {
        body.classList && body.classList.remove('dark');
        body.removeAttribute && body.removeAttribute('data-theme');
      }
    };
    applyLight();
    let ticks = 0;
    const iv = setInterval(() => { ticks++; applyLight(); if (ticks > 20) clearInterval(iv); }, 100);
    try {
      if (observer) observer.disconnect();
      observer = new MutationObserver(() => {
        if (getMode() === 'light' && html.classList.contains('dark')) {
          html.classList.remove('dark'); html.setAttribute('data-theme', 'light');
        }
      });
      observer.observe(html, { attributes: true, attributeFilter: ['class'] });
      setTimeout(() => { try { observer.disconnect(); } catch {} }, 10000);
    } catch {}
  }

  function apply(mode) {
    const html = document.documentElement;
    const isDark = (mode === 'dark');
    html.classList.toggle('dark', isDark);
    html.setAttribute('data-theme', isDark ? 'dark' : 'light');
    if (!isDark && document.body) {
      document.body.classList && document.body.classList.remove('dark');
      document.body.removeAttribute && document.body.removeAttribute('data-theme');
      hardLightGuard();
    }
  }

  // --- UI wiring (dropdown) ---
  function initUI() {
    const sel = document.getElementById('themeSelect');
    const stored = getMode();
    if (sel) {
      sel.value = stored;
      sel.addEventListener('change', () => {
        const v = sel.value === 'dark' ? 'dark' : 'light';
        setMode(v);
        apply(v);
      });
    }
    apply(stored); // pro jistotu po startu
  }

  // --- Boot: v Office i mimo Office ---
  function boot() {
    const start = () => initUI();
    if (typeof Office !== 'undefined' && Office.onReady) {
      Office.onReady(() => start());
    } else if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', start);
    } else {
      start();
    }
  }
  boot();

  // nic nevystavujeme globálně – vše je řízeno zde
})();
