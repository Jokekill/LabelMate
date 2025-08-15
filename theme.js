// ===== theme.js – Light/Dark only (žádné Auto), s pojistkou proti „přilepenému“ dark =====

window.Theme = window.Theme || {};

(function () {
  const LS_THEME = "labelmate_theme"; // 'light' | 'dark'
  let observer = null;

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

    // okamžitě + krátké okno (cca 2s), kdyby to jiné skripty zkusily znovu přilepit
    applyLight();
    let ticks = 0;
    const iv = setInterval(() => {
      ticks++; applyLight();
      if (ticks > 20) clearInterval(iv);
    }, 100);

    // krátce sleduj class na <html> (10s)
    try {
      if (observer) observer.disconnect();
      observer = new MutationObserver(() => {
        if (getMode() === 'light' && html.classList.contains('dark')) {
          html.classList.remove('dark');
          html.setAttribute('data-theme', 'light');
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

    apply(stored); // aplikuj při startu
  }

  window.Theme = { getMode, setMode, apply, initUI };
})();
