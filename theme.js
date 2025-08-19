(function () {
  const STORAGE_KEY = 'lm-theme'; // 'light' | 'dark'
  const root = document.documentElement;
  const listeners = new Set();

  function defaultTheme() {
    try {
      return (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches)
        ? 'dark' : 'light';
    } catch (_) {
      return 'light';
    }
  }

  function current() {
    try {
      return localStorage.getItem(STORAGE_KEY) || defaultTheme();
    } catch (_) {
      return defaultTheme();
    }
  }

  function syncColorSchemeMeta(theme) {
    let meta = document.querySelector('meta[name="color-scheme"]');
    if (!meta) {
      meta = document.createElement('meta');
      meta.setAttribute('name', 'color-scheme');
      document.head.appendChild(meta);
    }
    // pouze aktivní téma (žádný 'system')
    meta.setAttribute('content', theme === 'dark' ? 'dark' : 'light');
  }

  function apply(theme) {
    // theme je vždy 'light' nebo 'dark'
    root.setAttribute('data-theme', theme);
    syncColorSchemeMeta(theme);
    syncUI(theme);
    listeners.forEach(fn => { try { fn(theme); } catch (_) {} });
  }

  function set(theme) {
    const t = theme === 'dark' ? 'dark' : 'light';
    try { localStorage.setItem(STORAGE_KEY, t); } catch (_) {}
    apply(t);
  }

  function cycle() {
    set(current() === 'dark' ? 'light' : 'dark');
  }

  function syncUI(theme) {
    const btn = document.getElementById('theme-toggle');
    if (btn) {
      // ikonka = aktuální režim; title/aria zůstává anglicky, aby nebyla závislost na jazyku appky
      const icon = theme === 'dark' ? '🌙' : '☀️';
      const icEl = btn.querySelector('.icon');
      if (icEl) icEl.textContent = icon;
      btn.title = 'Toggle theme';
      btn.setAttribute('aria-label', 'Toggle theme');
    }
  }

  function wireUI() {
    const btn = document.getElementById('theme-toggle');
    if (btn && !btn.__lm_bound) {
      btn.addEventListener('click', cycle);
      btn.__lm_bound = true;
    }
  }

  function init() {
    apply(current());
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => { wireUI(); syncUI(current()); });
    } else {
      wireUI(); syncUI(current());
    }
  }

  // veřejné API (zachování kompatibility)
  window.Theme = {
    apply, current, set, cycle, init,
    onChange(fn) { listeners.add(fn); return () => listeners.delete(fn); }
  };

  init();
})();
