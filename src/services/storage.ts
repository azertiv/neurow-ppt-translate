import type { Settings } from "../utils/types";

const STORAGE_KEY = "slideTranslate.settings.v1";

export function defaultSettings(): Settings {
  return {
    apiKey: "",
    fromLang: "auto",
    toLang: "en",
    scope: "current",
    mode: "apply",
    keepLineBreaks: true,
    fitToLength: false,
    fitStrength: 60,
    glossary: {},
    ignoreRegex: "",
    applyUnderline: true
  };
}

function safeParseJSON<T>(raw: string | null): T | null {
  if (!raw) return null;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

export async function loadSettings(): Promise<Settings> {
  const base = defaultSettings();

  try {
    const rs = Office.context?.roamingSettings;
    if (rs) {
      const val = rs.get(STORAGE_KEY);
      if (val && typeof val === "object") {
        return { ...base, ...(val as Partial<Settings>) };
      }
    }
  } catch {
    // ignore
  }

  const local = safeParseJSON<Partial<Settings>>(localStorage.getItem(STORAGE_KEY));
  return { ...base, ...(local ?? {}) };
}

export async function saveSettings(next: Settings): Promise<void> {
  try {
    const rs = Office.context?.roamingSettings;
    if (rs) {
      rs.set(STORAGE_KEY, next);
      await new Promise<void>((resolve) => rs.saveAsync(() => resolve()));
    }
  } catch {
    // ignore
  }

  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
}
