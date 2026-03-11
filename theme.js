(function () {
  const STORAGE_KEY = "lm-theme"; // 'light' | 'dark'
  const root = document.documentElement;
  const listeners = new Set();

  function officePrefersDarkTheme() {
    try {
      if (typeof Office === "undefined" || !Office.context || !Office.context.officeTheme) return null;
      const t = Office.context.officeTheme;
      if (typeof t.isDarkTheme === "boolean") return t.isDarkTheme;
      return null;
    } catch (_) {
      return null;
    }
  }

  function defaultTheme() {
    // Prefer Office theme if available.
    const officeDark = officePrefersDarkTheme();
    if (officeDark === true) return "dark";
    if (officeDark === false) return "light";

    // Otherwise prefer OS theme.
    try {
      return (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) ? "dark" : "light";
    } catch (_) {
      return "light";
    }
  }

  function hasUserThemeOverride() {
    try {
      return !!localStorage.getItem(STORAGE_KEY);
    } catch (_) {
      return false;
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
      meta = document.createElement("meta");
      meta.setAttribute("name", "color-scheme");
      document.head.appendChild(meta);
    }
    meta.setAttribute("content", theme === "dark" ? "dark" : "light");
  }

  function syncUI(theme) {
    const btn = document.getElementById("theme-toggle");
    if (btn) {
      const icon = theme === "dark" ? "🌙" : "☀️";
      const icEl = btn.querySelector(".icon");
      if (icEl) icEl.textContent = icon;
      btn.title = "Toggle theme";
      btn.setAttribute("aria-label", "Toggle theme");
    }
  }

  function apply(theme) {
    const t = theme === "dark" ? "dark" : "light";
    root.setAttribute("data-theme", t);
    syncColorSchemeMeta(t);
    syncUI(t);
    listeners.forEach((fn) => {
      try { fn(t); } catch (_) {}
    });
  }

  function set(theme) {
    const t = theme === "dark" ? "dark" : "light";
    try { localStorage.setItem(STORAGE_KEY, t); } catch (_) {}
    apply(t);
  }

  function cycle() {
    set(current() === "dark" ? "light" : "dark");
  }

  function wireUI() {
    const btn = document.getElementById("theme-toggle");
    if (btn && !btn.__lm_bound) {
      btn.addEventListener("click", cycle);
      btn.__lm_bound = true;
    }
  }

  function init() {
    apply(current());

    // If the user hasn't explicitly chosen a theme, try to align with Office theme
    // once Office is ready. Office.onReady can be called from multiple places and
    // resolves immediately if Office is already ready. citeturn8search4turn8search2
    try {
      if (!hasUserThemeOverride() && typeof Office !== "undefined" && Office.onReady) {
        Office.onReady(() => {
          if (!hasUserThemeOverride()) apply(defaultTheme());
        });
      }
    } catch (_) {}

    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", () => {
        wireUI();
        syncUI(current());
      });
    } else {
      wireUI();
      syncUI(current());
    }
  }

  // Public API (backward compatible)
  window.Theme = {
    apply,
    current,
    set,
    cycle,
    init,
    onChange(fn) {
      listeners.add(fn);
      return () => listeners.delete(fn);
    },
  };

  init();
})();
