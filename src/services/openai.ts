import type { Settings, TranslationResult } from "../utils/types";
import { labelFor } from "../utils/language";

export class OpenAIError extends Error {
  constructor(
    message: string,
    public status?: number,
    public detail?: unknown
  ) {
    super(message);
  }
}

function extractOutputText(resp: any): string {
  const chunks: string[] = [];
  const outputs = resp?.output;
  if (Array.isArray(outputs)) {
    for (const item of outputs) {
      if (item?.type !== "message") continue;
      const contents = item?.content;
      if (!Array.isArray(contents)) continue;
      for (const c of contents) {
        if (c?.type === "output_text" && typeof c?.text === "string") {
          chunks.push(c.text);
        }
      }
    }
  }
  return chunks.join("\n").trim();
}

function parseGlossary(g: Record<string, string>): string {
  const keys = Object.keys(g);
  if (keys.length === 0) return "";
  return `\nGlossaire (prioritaire) :\n${keys.map((k) => `- ${k} => ${g[k]}`).join("\n")}`;
}

export async function testOpenAIKey(settings: Settings): Promise<void> {
  if (!settings.apiKey) throw new OpenAIError("Aucune clé API.");

  const res = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${settings.apiKey}`
    },
    body: JSON.stringify({
      model: settings.model || "gpt-5-mini",
      input: "Reply with the single word OK.",
      temperature: 0,
      store: false
    })
  });

  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new OpenAIError((json?.error?.message as string) || "Erreur OpenAI", res.status, json);
  }

  const out = extractOutputText(json);
  if (!out.toLowerCase().includes("ok")) {
    throw new OpenAIError("La clé a répondu, mais la réponse n'est pas attendue.", res.status, json);
  }
}

export interface TranslateBatchItem {
  paragraphId: string;
  originalChars: number;
  runs: { index: number; text: string }[];
}

export async function translateBatch(items: TranslateBatchItem[], settings: Settings): Promise<TranslationResult[]> {
  if (!settings.apiKey) throw new OpenAIError("Aucune clé API.");
  if (!items.length) return [];

  const fromLabel = settings.fromLang === "auto" ? "auto" : labelFor(settings.fromLang);
  const toLabel = labelFor(settings.toLang);
  const fit = settings.fitToLength ? settings.fitStrength : 0;

  const schema = {
    type: "object",
    additionalProperties: false,
    properties: {
      items: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            paragraphId: { type: "string" },
            translatedRuns: {
              type: "array",
              items: {
                type: "object",
                additionalProperties: false,
                properties: {
                  index: { type: "integer" },
                  text: { type: "string" }
                },
                required: ["index", "text"]
              }
            }
          },
          required: ["paragraphId", "translatedRuns"]
        }
      }
    },
    required: ["items"]
  } as const;

  const instructions = [
    "You are a high-precision translation engine for PowerPoint.",
    `Translate from ${fromLabel} to ${toLabel}.`,
    "CRITICAL: Keep the number of runs exactly the same for each paragraph and keep them in the same order.",
    "Return translated text per run index. Do not reorder runs.",
    "Preserve leading/trailing whitespace of each run exactly.",
    "Do NOT translate protected tokens like {0}, {{name}}, %s, URLs, email addresses, or product codes; keep them unchanged.",
    settings.fitToLength
      ? `Try to keep total paragraph length close to original. Strength: ${fit}% (higher=closer).`
      : "Length fitting is disabled; prioritize best translation.",
    parseGlossary(settings.glossary)
  ].join("\n");

  const payload = {
    task: "translate_powerpoint_text",
    from: fromLabel,
    to: toLabel,
    keepSegmentation: true,
    keepWhitespace: true,
    preserveNumbersAndSymbols: true,
    fitToOriginalLengthPercent: fit,
    glossary: settings.glossary,
    items
  };

  const res = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${settings.apiKey}`
    },
    body: JSON.stringify({
      model: settings.model,
      temperature: settings.temperature,
      store: false,
      input: [
        { role: "system", content: instructions },
        { role: "user", content: JSON.stringify(payload) }
      ],
      text: {
        format: {
          type: "json_schema",
          name: "ppt_translation",
          strict: true,
          schema
        }
      }
    })
  });

  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new OpenAIError((json?.error?.message as string) || "Erreur OpenAI", res.status, json);
  }

  const out = extractOutputText(json);
  if (!out) throw new OpenAIError("Réponse vide du modèle.", res.status, json);

  let parsed: any;
  try {
    parsed = JSON.parse(out);
  } catch {
    throw new OpenAIError("Impossible de parser la sortie JSON du modèle.", res.status, { out, raw: json });
  }

  if (!parsed?.items || !Array.isArray(parsed.items)) {
    throw new OpenAIError("Sortie JSON inattendue.", res.status, parsed);
  }

  return parsed.items as TranslationResult[];
}
