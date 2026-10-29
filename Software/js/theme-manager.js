// ========================================================
// VerbFlow - Gestor de Tema Visual Dual (Claro / Oscuro)
// Modo Claro: Blanco puro y tonos de azules campus
// Modo Oscuro: Slate Navy confortable (no negro deprimente)
// I.E. Promoción Social • Palermo, Huila
// ========================================================

const ThemeManager = (() => {
  const THEME_KEY = "vf_theme_preference";

  function getStoredTheme() {
    try {
      const stored = localStorage.getItem(THEME_KEY);
      if (stored === "light" || stored === "dark") return stored;
    } catch (e) {
      // Ignorar si localStorage está deshabilitado
    }
    // Por defecto institucional: modo claro (blanco y azul campus)
    return "light";
  }

  function applyTheme(theme) {
    const isDark = theme === "dark";
    document.documentElement.setAttribute("data-theme", theme);
    if (isDark) {
      document.documentElement.classList.add("dark");
    } else {
      document.documentElement.classList.remove("dark");
    }

    try {
      localStorage.setItem(THEME_KEY, theme);
    } catch (e) {}

    // Actualizar todos los botones e indicadores de tema en pantalla
    document.querySelectorAll(".vf-theme-toggle, [data-action='toggle-theme']").forEach(btn => {
      const iconSpan = btn.querySelector(".material-symbols-outlined");
      const labelSpan = btn.querySelector(".vf-theme-label") || btn.querySelector("span:not(.material-symbols-outlined)");
      
      if (isDark) {
        if (iconSpan) {
          iconSpan.textContent = "light_mode";
          iconSpan.className = "material-symbols-outlined text-amber-400";
        }
        if (labelSpan) labelSpan.textContent = "Modo Claro";
        btn.setAttribute("aria-label", "Cambiar a modo claro");
      } else {
        if (iconSpan) {
          iconSpan.textContent = "dark_mode";
          iconSpan.className = "material-symbols-outlined text-sky-600";
        }
        if (labelSpan) labelSpan.textContent = "Modo Oscuro";
        btn.setAttribute("aria-label", "Cambiar a modo oscuro");
      }
    });
  }

  function toggleTheme() {
    const current = document.documentElement.getAttribute("data-theme") || "light";
    const next = current === "dark" ? "light" : "dark";
    applyTheme(next);
    return next;
  }

  function init() {
    const initial = getStoredTheme();
    applyTheme(initial);

    // Delegación de eventos para clicks en botones de tema
    document.addEventListener("click", (e) => {
      const targetBtn = e.target.closest(".vf-theme-toggle, [data-action='toggle-theme']");
      if (targetBtn) {
        e.preventDefault();
        toggleTheme();
      }
    });
  }

  // Ejecutar inmediatamente para evitar parpadeos de carga (FOUC)
  if (typeof document !== "undefined") {
    const preTheme = getStoredTheme();
    document.documentElement.setAttribute("data-theme", preTheme);
    if (preTheme === "dark") {
      document.documentElement.classList.add("dark");
    }
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", init);
    } else {
      init();
    }
  }

  return {
    getTheme: () => document.documentElement.getAttribute("data-theme") || "light",
    applyTheme,
    toggleTheme,
    init
  };
})();

if (typeof window !== "undefined") {
  window.ThemeManager = ThemeManager;
}
