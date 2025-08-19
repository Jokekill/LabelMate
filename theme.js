(function () {
  const STORAGE_KEY = 'lm-theme'; // 'light' | 'dark' | 'system'
  const root = document.documentElement;
  const listeners = new Set();

  function syncColorSchemeMeta(theme) {
    let meta = document.querySelector('meta[name="color-scheme"]');
    if (!meta) {
      meta = document.createElement('meta');
      meta.setAttribute('name', 'color-scheme');
      document.head.appendChild(meta);
    }
    if (theme === 'light') meta.setAttribute('content', 'light');
    else if (theme === 'dark') meta.setAttribute('content', 'dark');
    else meta.setAttribute('content', 'light dark');
  }

  function prefersDarkOS() {
    try { return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches; }
    catch (_) { return false; }
  }

  function apply(theme) {
    if (theme === 'system' || !theme) {
      root.removeAttribute('data-theme');
    } else {
      root.setAttribute('data-theme', theme);
    }
    syncColorSchemeMeta(theme);
    syncUI(theme);
    // notifikuj posluchače
    listeners.forEach(fn => { try { fn(theme); } catch (_) {} });
  }

  function current() {
    try {
      return localStorage.getItem(STORAGE_KEY) || 'system';
    } catch (_) {
      return 'system';
    }
  }

  function set(theme) {
    try { localStorage.setItem(STORAGE_KEY, theme); } catch (_) {}
    apply(theme);
  }

  function cycle() {
    const c = current();
    const next = c === 'light' ? 'dark' : c === 'dark' ? 'system' : 'light';
    set(next);
  }

  function syncUI(theme) {
    // Tlačítko
    const btn = document.getElementById('theme-toggle');
    if (btn) {
      const state = theme === 'light' ? 'Světlý' : theme === 'dark' ? 'Tmavý' : 'Systém';
      const icon = theme === 'light' ? '☀️' : theme === 'dark' ? '🌙' : '🖥️';
      const stEl = btn.querySelector('.state');
      const icEl = btn.querySelector('.icon');
      if (stEl) stEl.textContent = state;
      if (icEl) icEl.textContent = icon;
      btn.title = `Motiv: ${state}`;
    }

    // (Volitelný) select, pokud ho máš někde jinde v UI
    const sel = document.getElementById('themeSelect');
    if (sel) {
      const hasSystem = Array.from(sel.options).some(o => o.value === 'system');
      if (hasSystem) {
        sel.value = theme;
      } else {
        // Pokud select umí jen 'light' a 'dark', namapuj 'system' na preferenci OS
        sel.value = theme === 'system' ? (prefersDarkOS() ? 'dark' : 'light') : theme;
      }
    }
  }

  function wireUI() {
    const btn = document.getElementById('theme-toggle');
    if (btn && !btn.__lm_bound) {
      btn.addEventListener('click', cycle);
      btn.__lm_bound = true;
    }
    const sel = document.getElementById('themeSelect');
    if (sel && !sel.__lm_bound) {
      sel.addEventListener('change', e => set(e.target.value));
      sel.__lm_bound = true;
    }
  }

  function watchOSChanges() {
    try {
      const mq = window.matchMedia('(prefers-color-scheme: dark)');
      const handler = () => { if (current() === 'system') apply('system'); };
      if (mq.addEventListener) mq.addEventListener('change', handler);
      else if (mq.addListener) mq.addListener(handler);
    } catch (_) {}
  }

  function init() {
    apply(current());
    watchOSChanges();
    // UI prvky se mohou objevit později — naváž je po DOMContentLoaded
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => { wireUI(); syncUI(current()); });
    } else {
      wireUI(); syncUI(current());
    }
  }

  // Veřejné API + zpětná kompatibilita (některé části appky očekávaly `Theme`)
  window.Theme = {
    apply, current, set, cycle, init,
    onChange(fn) { listeners.add(fn); return () => listeners.delete(fn); }
  };

  // Auto-init
  init();
})();
