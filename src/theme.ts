import { useCallback, useEffect, useState } from "react";

export type Theme = "light" | "dark";

const themeStorageKey = "dct-contract-assistant.theme";

function storedTheme(): Theme | null {
  try {
    const value = localStorage.getItem(themeStorageKey);
    return value === "light" || value === "dark" ? value : null;
  } catch {
    return null;
  }
}

// An explicit choice wins; otherwise use the Teams theme applied at startup, then the OS.
function initialTheme(): Theme {
  const saved = storedTheme();
  if (saved) return saved;
  const fromHost = document.documentElement.dataset.theme;
  if (fromHost === "light" || fromHost === "dark") return fromHost;
  return window.matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark";
}

export function useTheme() {
  const [theme, setTheme] = useState<Theme>(initialTheme);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    document.querySelector('meta[name="theme-color"]')?.setAttribute("content", theme === "dark" ? "#0a0c10" : "#f4f6fa");
  }, [theme]);

  const toggleTheme = useCallback(() => {
    setTheme((current) => {
      const next = current === "dark" ? "light" : "dark";
      try {
        localStorage.setItem(themeStorageKey, next);
      } catch {
        // Storage can be unavailable in private windows; the toggle still works for this visit.
      }
      return next;
    });
  }, []);

  return { theme, toggleTheme };
}
