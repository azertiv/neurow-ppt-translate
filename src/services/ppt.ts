import type { Paragraph, Settings, SlideAnalysis, TranslationResult } from "../utils/types";
import type { TranslateBatchItem } from "./openai";
import { Logger } from "./logger";
import { translateBatch } from "./openai";
import {
  applyRunTranslations,
  extractShapeTextParagraphs,
  queueFontRuns,
  queueParagraphFormats
} from "./formatting";

export interface ShapeTextTarget {
  kind: "shapeText";
  shapeId: string;
  shapeName: string;
  paragraphs: Paragraph[];
}

export interface TableRunSnapshot {
  text: string;
  font?: any; // PowerPoint.FontProperties
}

export interface TableCellTarget {
  kind: "tableCell";
  shapeId: string;
  shapeName: string;
  row: number;
  col: number;
  paragraphId: string;
  originalChars: number;
  runs: TableRunSnapshot[];
}

export interface SlideTargets {
  slideIndex: number;
  shapeTextTargets: ShapeTextTarget[];
  tableCellTargets: TableCellTarget[];
}

export async function getSlideIndices(scope: Settings["scope"]): Promise<number[]> {
  return PowerPoint.run(async (context) => {
    if (scope === "all") {
      const count = context.presentation.slides.getCount();
      await context.sync();
      return Array.from({ length: count.value }, (_, i) => i);
    }

    const selected = context.presentation.getSelectedSlides();
    const selectedCount = selected.getCount();
    await context.sync();

    if (selectedCount.value > 0) {
      selected.load("items/index");
      await context.sync();
      return selected.items.map((s) => s.index);
    }

    const active = context.presentation.getActiveSlideOrNullObject();
    active.load("isNullObject,index");
    await context.sync();

    if (!active.isNullObject) return [active.index];
    return [0];
  });
}

function compileIgnoreRegex(pattern: string): RegExp | null {
  const p = pattern?.trim();
  if (!p) return null;
  try {
    return new RegExp(p);
  } catch {
    return null;
  }
}

export async function analyzeScope(settings: Settings, logger?: Logger): Promise<SlideAnalysis[]> {
  const indices = await getSlideIndices(settings.scope);
  const ignore = compileIgnoreRegex(settings.ignoreRegex);

  const out: SlideAnalysis[] = [];

  for (const slideIndex of indices) {
    const analysis = await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(slideIndex);
      slide.load("shapes/items/id,shapes/items/name,shapes/items/type");
      await context.sync();

      let textBoxes = 0;
      let tables = 0;
      let paragraphs = 0;
      let characters = 0;

      for (const shape of slide.shapes.items) {
        const name = shape.name ?? "";
        if (ignore && ignore.test(name)) continue;

        if ((shape.type as any) === PowerPoint.ShapeType.table) {
          tables++;
          // Approximate: count later
          continue;
        }
        if ((shape.type as any) === PowerPoint.ShapeType.group) continue;

        const tf = shape.getTextFrameOrNullObject();
        tf.load("isNullObject,hasText");
        await context.sync();
        if (tf.isNullObject || !tf.hasText) continue;

        tf.textRange.load("text");
        await context.sync();
        const text = tf.textRange.text ?? "";
        if (!text.trim()) continue;
        textBoxes++;
        const ps = text.split("\n");
        paragraphs += ps.length;
        characters += text.length;
      }

      return { slideIndex, textBoxes, tables, paragraphs, characters } as SlideAnalysis;
    });

    out.push(analysis);
    logger?.log(
      `Slide ${slideIndex + 1} — ${analysis.textBoxes} zone(s) texte, ${analysis.tables} table(s), ${analysis.paragraphs} paragraphe(s)`,
      "dim"
    );
  }

  return out;
}

export async function extractSlideTargets(
  slideIndex: number,
  settings: Settings,
  logger?: Logger
): Promise<SlideTargets> {
  const ignore = compileIgnoreRegex(settings.ignoreRegex);

  return PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(slideIndex);
    slide.load("shapes/items/id,shapes/items/name,shapes/items/type");
    await context.sync();

    const shapeTextTargets: ShapeTextTarget[] = [];
    const tableCellTargets: TableCellTarget[] = [];

    for (const shape of slide.shapes.items) {
      const shapeId = shape.id;
      const shapeName = shape.name ?? "";

      if (ignore && ignore.test(shapeName)) {
        logger?.log(`Ignoré: ${shapeName || shapeId}`, "dim");
        continue;
      }

      if ((shape.type as any) === PowerPoint.ShapeType.group) {
        logger?.log(`Group ignoré (limitation Office.js): ${shapeName || shapeId}`, "dim");
        continue;
      }

      if ((shape.type as any) === PowerPoint.ShapeType.table) {
        const table = shape.getTable();
        table.load("rowCount,columnCount");
        await context.sync();

        const cellObjs: PowerPoint.TableCell[] = [];
        const coords: Array<{ r: number; c: number }> = [];

        for (let r = 0; r < table.rowCount; r++) {
          for (let c = 0; c < table.columnCount; c++) {
            const cell = table.getCellOrNullObject(r, c);
            cell.load("isNullObject,textRuns,text");
            cellObjs.push(cell);
            coords.push({ r, c });
          }
        }
        await context.sync();

        for (let i = 0; i < cellObjs.length; i++) {
          const cell = cellObjs[i];
          if ((cell as any).isNullObject) continue;

          const runs = (cell as any).textRuns as any[];
          const text = (cell as any).text as string;
          if (!text || !text.trim()) continue;

          const runSnapshots: TableRunSnapshot[] = Array.isArray(runs) && runs.length
            ? runs.map((tr) => ({ text: tr.text ?? "", font: tr.font }))
            : [{ text, font: undefined }];

          const id = `s${slideIndex}_shape${shapeId}_cell${coords[i].r}_${coords[i].c}`;
          tableCellTargets.push({
            kind: "tableCell",
            shapeId,
            shapeName,
            row: coords[i].r,
            col: coords[i].c,
            paragraphId: id,
            originalChars: text.length,
            runs: runSnapshots
          });
        }
        continue;
      }

      const tf = shape.getTextFrameOrNullObject();
      tf.load("isNullObject,hasText");
      await context.sync();
      if (tf.isNullObject || !tf.hasText) continue;

      const paragraphs = await extractShapeTextParagraphs(
        context,
        tf.textRange,
        slideIndex,
        shapeId,
        shapeName
      );

      const useful = paragraphs.some((p) => p.runs.some((r) => r.text.trim().length > 0));
      if (!useful) continue;

      shapeTextTargets.push({ kind: "shapeText", shapeId, shapeName, paragraphs });
    }

    logger?.log(
      `Slide ${slideIndex + 1}: ${shapeTextTargets.length} shape(s) texte, ${tableCellTargets.length} cellule(s) de table`,
      "dim"
    );

    return { slideIndex, shapeTextTargets, tableCellTargets };
  });
}

function buildTranslateItems(targets: SlideTargets): TranslateBatchItem[] {
  const items: TranslateBatchItem[] = [];
  for (const t of targets.shapeTextTargets) {
    for (const p of t.paragraphs) {
      items.push({
        paragraphId: p.id,
        originalChars: p.originalCharCount,
        runs: p.runs.map((r, idx) => ({ index: idx, text: r.text }))
      });
    }
  }
  for (const t of targets.tableCellTargets) {
    items.push({
      paragraphId: t.paragraphId,
      originalChars: t.originalChars,
      runs: t.runs.map((r, idx) => ({ index: idx, text: r.text }))
    });
  }
  return items;
}

function chunkByChars<T extends { originalChars: number }>(
  items: T[],
  maxChars = 6000,
  maxItems = 35
): T[][] {
  const chunks: T[][] = [];
  let current: T[] = [];
  let sum = 0;
  for (const item of items) {
    if (current.length >= maxItems || sum + item.originalChars > maxChars) {
      if (current.length) chunks.push(current);
      current = [];
      sum = 0;
    }
    current.push(item);
    sum += item.originalChars;
  }
  if (current.length) chunks.push(current);
  return chunks;
}

export async function translateAndMaybeApplySlide(
  targets: SlideTargets,
  settings: Settings,
  logger: Logger,
  abortSignal?: AbortSignal
): Promise<{ translated: number; preview: string }>
{
  const items = buildTranslateItems(targets);
  if (!items.length) return { translated: 0, preview: "" };

  const chunks = chunkByChars(items);
  const allResults: TranslationResult[] = [];

  for (let i = 0; i < chunks.length; i++) {
    if (abortSignal?.aborted) break;
    logger.log(`OpenAI: lot ${i + 1}/${chunks.length}…`, "dim");
    const res = await translateBatch(chunks[i], settings);
    allResults.push(...res);
  }

  const map = new Map<string, TranslationResult>();
  for (const r of allResults) map.set(r.paragraphId, r);

  // Build a small preview (first 3 paragraphs)
  const previewParts: string[] = [];
  let previewCount = 0;
  for (const st of targets.shapeTextTargets) {
    for (const p of st.paragraphs) {
      const tr = map.get(p.id);
      if (!tr) continue;
      const applied = applyRunTranslations(p, tr.translatedRuns);
      const text = applied.runs.map((r) => r.text).join("");
      if (text.trim()) {
        previewParts.push(text);
        previewCount++;
      }
      if (previewCount >= 3) break;
    }
    if (previewCount >= 3) break;
  }

  if (settings.mode === "preview") {
    return { translated: map.size, preview: previewParts.join("\n\n") };
  }

  // Apply
  await applySlideTranslations(targets.slideIndex, targets, map, settings, logger);

  return { translated: map.size, preview: previewParts.join("\n\n") };
}

export async function applySlideTranslations(
  slideIndex: number,
  targets: SlideTargets,
  translationMap: Map<string, TranslationResult>,
  settings: Settings,
  logger?: Logger
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(slideIndex);

    // Shapes (text frames)
    for (const st of targets.shapeTextTargets) {
      const shape = slide.shapes.getItem(st.shapeId);
      const tf = shape.getTextFrameOrNullObject();
      tf.load("isNullObject,hasText");
      await context.sync();
      if (tf.isNullObject || !tf.hasText) continue;

      const updatedParagraphs: Paragraph[] = st.paragraphs.map((p) => {
        const r = translationMap.get(p.id);
        return r ? applyRunTranslations(p, r.translatedRuns) : p;
      });

      const composed = composeText(updatedParagraphs, settings);

      tf.textRange.text = composed.fullText;
      // Queue formats on the new text.
      queueParagraphFormats(tf.textRange, composed.paragraphSpans);
      queueFontRuns(tf.textRange, composed.runSpans, settings);

      logger?.log(`Appliqué: ${st.shapeName || st.shapeId}`, "dim");
    }

    // Tables
    for (const tc of targets.tableCellTargets) {
      const shape = slide.shapes.getItem(tc.shapeId);
      const table = shape.getTable();
      const cell = table.getCellOrNullObject(tc.row, tc.col);

      const tr = translationMap.get(tc.paragraphId);
      if (!tr) continue;

      const runMap = new Map<number, string>(tr.translatedRuns.map((r) => [r.index, r.text]));
      const newTextRuns = tc.runs.map((r, idx) => ({
        text: runMap.get(idx) ?? r.text,
        font: r.font
      }));

      (cell as any).set({ textRuns: newTextRuns });
    }

    await context.sync();
  });
}

function composeText(paragraphs: Paragraph[], settings: Settings): {
  fullText: string;
  paragraphSpans: { start: number; length: number; format?: any }[];
  runSpans: { start: number; length: number; font: any }[];
} {
  let fullText = "";
  const paragraphSpans: { start: number; length: number; format?: any }[] = [];
  const runSpans: { start: number; length: number; font: any }[] = [];

  let offset = 0;
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const pStart = offset;

    for (const run of p.runs) {
      fullText += run.text;
      runSpans.push({ start: offset, length: run.text.length, font: run.font });
      offset += run.text.length;
    }

    const pLen = offset - pStart;
    paragraphSpans.push({ start: pStart, length: pLen, format: p.paragraphFormat });

    if (settings.keepLineBreaks && i < paragraphs.length - 1) {
      fullText += "\n";
      offset += 1;
    } else if (!settings.keepLineBreaks && i < paragraphs.length - 1) {
      fullText += " ";
      offset += 1;
    }
  }

  return { fullText, paragraphSpans, runSpans };
}

export async function translateScope(
  settings: Settings,
  logger: Logger,
  onProgress: (done: number, total: number, label: string) => void,
  abortSignal?: AbortSignal
): Promise<{ translated: number; preview?: string }>
{
  const indices = await getSlideIndices(settings.scope);
  const total = indices.length;
  let translatedTotal = 0;
  let preview = "";

  for (let i = 0; i < indices.length; i++) {
    if (abortSignal?.aborted) break;
    const slideIndex = indices[i];

    onProgress(i, total, `Extraction slide ${slideIndex + 1}/${total}`);
    logger.log(`Slide ${slideIndex + 1} — extraction…`);

    const targets = await extractSlideTargets(slideIndex, settings, logger);

    const count =
      targets.shapeTextTargets.reduce((a, t) => a + t.paragraphs.length, 0) +
      targets.tableCellTargets.length;

    if (count === 0) {
      logger.log(`Slide ${slideIndex + 1} — rien à traduire.`, "dim");
      onProgress(i + 1, total, `Slide ${slideIndex + 1} terminé (0)`);
      continue;
    }

    onProgress(i, total, `Traduction slide ${slideIndex + 1}/${total}`);
    logger.log(`Slide ${slideIndex + 1} — traduction (${count} bloc(s))…`);

    const res = await translateAndMaybeApplySlide(targets, settings, logger, abortSignal);
    translatedTotal += res.translated;
    if (!preview && res.preview) preview = res.preview;

    onProgress(i + 1, total, `Slide ${slideIndex + 1} terminé`);
  }

  onProgress(total, total, abortSignal?.aborted ? "Annulé" : "Terminé");
  return { translated: translatedTotal, preview };
}
