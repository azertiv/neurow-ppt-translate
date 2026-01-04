import { LANGUAGES } from "../utils/language";
import type { Settings, Scope, Mode } from "../utils/types";
import { parseGlossaryText } from "../utils/text";
import { Logger } from "../services/logger";
import { analyzeScope, translateScope } from "../services/ppt";
import { loadSettings, saveSettings, defaultSettings } from "../services/storage";
import { testOpenAIKey } from "../services/openai";

function $(id: string): HTMLElement {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element: ${id}`);
  return el;
}

let settings: Settings = defaultSettings();
let logger: Logger;
let abortController: AbortController | null = null;

function setStatus(text: string, kind: "ready" | "busy" | "error" = "ready") {
  const pill = $("statusPill");
  pill.textContent = text;
  pill.classList.remove("busy", "error");
  if (kind !== "ready") pill.classList.add(kind);
}

function setProgress(done: number, total: number) {
  const bar = $("progressBar") as HTMLDivElement;
  const pct = total <= 0 ? 0 : Math.min(100, Math.round((done / total) * 100));
  bar.style.width = `${pct}%`;
}

function updateFitLabel() {
  const v = Number(($("fitStrength") as HTMLInputElement).value);
  $("fitStrengthLabel").textContent = `${v}%`;
}

function readUI(): Settings {
  const next: Settings = {
    ...settings,
    apiKey: ( $("apiKey") as HTMLInputElement).value.trim(),
    model: ( $("model") as HTMLSelectElement).value,
    temperature: Number(( $("temperature") as HTMLInputElement).value),
    fromLang: ( $("fromLang") as HTMLSelectElement).value,
    toLang: ( $("toLang") as HTMLSelectElement).value,
    keepLineBreaks: ( $("keepLineBreaks") as HTMLInputElement).checked,
    fitToLength: ( $("fitToLength") as HTMLInputElement).checked,
    fitStrength: Number(( $("fitStrength") as HTMLInputElement).value),
    glossary: parseGlossaryText(( $("glossary") as HTMLTextAreaElement).value),
    ignoreRegex: ( $("ignoreRegex") as HTMLInputElement).value,
    applyUnderline: ( $("applyUnderline") as HTMLInputElement).checked
  };
  return next;
}

async function writeUI(s: Settings) {
  ( $("apiKey") as HTMLInputElement).value = s.apiKey;
  ( $("model") as HTMLSelectElement).value = s.model;
  ( $("temperature") as HTMLInputElement).value = String(s.temperature);
  ( $("fromLang") as HTMLSelectElement).value = s.fromLang;
  ( $("toLang") as HTMLSelectElement).value = s.toLang;
  ( $("keepLineBreaks") as HTMLInputElement).checked = s.keepLineBreaks;
  ( $("fitToLength") as HTMLInputElement).checked = s.fitToLength;
  ( $("fitStrength") as HTMLInputElement).value = String(s.fitStrength);
  ( $("ignoreRegex") as HTMLInputElement).value = s.ignoreRegex;
  ( $("applyUnderline") as HTMLInputElement).checked = s.applyUnderline;

  // glossary back to text
  const g = Object.keys(s.glossary)
    .map((k) => `${k}=${s.glossary[k]}`)
    .join("\n");
  ( $("glossary") as HTMLTextAreaElement).value = g;

  // scope/mode segmented
  setScopeUI(s.scope);
  setModeUI(s.mode);

  updateFitLabel();
}

function setScopeUI(scope: Scope) {
  const cur = $("scopeCurrent");
  const all = $("scopeAll");
  cur.classList.toggle("active", scope === "current");
  all.classList.toggle("active", scope === "all");
  cur.setAttribute("aria-selected", scope === "current" ? "true" : "false");
  all.setAttribute("aria-selected", scope === "all" ? "true" : "false");
}

function setModeUI(mode: Mode) {
  const a = $("modeApply");
  const p = $("modePreview");
  a.classList.toggle("active", mode === "apply");
  p.classList.toggle("active", mode === "preview");
  a.setAttribute("aria-selected", mode === "apply" ? "true" : "false");
  p.setAttribute("aria-selected", mode === "preview" ? "true" : "false");
}

async function persistFromUI() {
  settings = readUI();
  await saveSettings(settings);
}

function initLanguageSelects() {
  const from = $("fromLang") as HTMLSelectElement;
  const to = $("toLang") as HTMLSelectElement;

  from.innerHTML = "";
  to.innerHTML = "";

  for (const lang of LANGUAGES) {
    const o1 = document.createElement("option");
    o1.value = lang.code;
    o1.textContent = lang.label;
    from.appendChild(o1);

    if (lang.code !== "auto") {
      const o2 = document.createElement("option");
      o2.value = lang.code;
      o2.textContent = lang.label;
      to.appendChild(o2);
    }
  }
}

function bindEvents() {
  // advanced accordion
  const toggle = $("advancedToggle");
  const body = $("advancedBody");
  toggle.addEventListener("click", () => {
    const isOpen = toggle.getAttribute("aria-expanded") === "true";
    toggle.setAttribute("aria-expanded", isOpen ? "false" : "true");
    body.toggleAttribute("hidden", isOpen);
  });

  $("scopeCurrent").addEventListener("click", async () => {
    settings.scope = "current";
    setScopeUI(settings.scope);
    await persistFromUI();
  });
  $("scopeAll").addEventListener("click", async () => {
    settings.scope = "all";
    setScopeUI(settings.scope);
    await persistFromUI();
  });

  $("modeApply").addEventListener("click", async () => {
    settings.mode = "apply";
    setModeUI(settings.mode);
    await persistFromUI();
  });
  $("modePreview").addEventListener("click", async () => {
    settings.mode = "preview";
    setModeUI(settings.mode);
    await persistFromUI();
  });

  $("fitStrength").addEventListener("input", () => {
    updateFitLabel();
  });

  // auto-save on changes
  const autosaveIds = [
    "apiKey",
    "model",
    "temperature",
    "fromLang",
    "toLang",
    "keepLineBreaks",
    "fitToLength",
    "fitStrength",
    "glossary",
    "ignoreRegex",
    "applyUnderline"
  ];
  for (const id of autosaveIds) {
    $(id).addEventListener("change", () => void persistFromUI());
  }

  $("clearLogsBtn").addEventListener("click", () => logger.clear());

  $("testKeyBtn").addEventListener("click", async () => {
    try {
      setStatus("Test clé…", "busy");
      await persistFromUI();
      await testOpenAIKey(settings);
      logger.log("Clé OpenAI OK ✅");
      setStatus("Prêt", "ready");
    } catch (e: any) {
      logger.log(`Test clé échoué: ${e?.message ?? e}`, "error");
      setStatus("Erreur", "error");
    }
  });

  $("analyzeBtn").addEventListener("click", async () => {
    try {
      setStatus("Analyse…", "busy");
      await persistFromUI();
      logger.log("Analyse de la portée…", "dim");
      const analyses = await analyzeScope(settings, logger);
      const totals = analyses.reduce(
        (acc, s) => {
          acc.textBoxes += s.textBoxes;
          acc.tables += s.tables;
          acc.paragraphs += s.paragraphs;
          acc.characters += s.characters;
          return acc;
        },
        { textBoxes: 0, tables: 0, paragraphs: 0, characters: 0 }
      );
      $("metrics").textContent = `${analyses.length} slide(s) · ${totals.textBoxes} zone(s) texte · ${totals.tables} table(s) · ${totals.paragraphs} paragraphes · ${totals.characters} caractères`;
      logger.log("Analyse terminée.", "dim");
      setStatus("Prêt", "ready");
    } catch (e: any) {
      logger.log(`Analyse échouée: ${e?.message ?? e}`, "error");
      setStatus("Erreur", "error");
    }
  });

  $("translateBtn").addEventListener("click", async () => {
    if (abortController) return;

    abortController = new AbortController();
    ( $("cancelBtn") as HTMLButtonElement).disabled = false;

    try {
      setStatus("Traduction…", "busy");
      await persistFromUI();

      const res = await translateScope(
        settings,
        logger,
        (done, total, label) => {
          setProgress(done, total);
          $("metrics").textContent = label;
        },
        abortController.signal
      );

      if (settings.mode === "preview" && res.preview) {
        logger.log("Prévisualisation (extrait) :", "dim");
        logger.log(res.preview, "info");
      }

      logger.log(`Terminé. Paragraphes traduits: ${res.translated}.`, "info");
      setStatus(abortController.signal.aborted ? "Annulé" : "Prêt", "ready");
    } catch (e: any) {
      logger.log(`Traduction échouée: ${e?.message ?? e}`, "error");
      setStatus("Erreur", "error");
    } finally {
      abortController = null;
      ( $("cancelBtn") as HTMLButtonElement).disabled = true;
    }
  });

  $("cancelBtn").addEventListener("click", () => {
    abortController?.abort();
    logger.log("Annulation demandée…", "warn");
  });
}

Office.onReady(async () => {
  logger = new Logger($("logBox"));
  initLanguageSelects();
  bindEvents();

  settings = await loadSettings();
  await writeUI(settings);

  logger.log("Add-in chargé. Sélectionne une slide puis clique sur Analyser ou Traduire.", "dim");
  setStatus("Prêt", "ready");
});
