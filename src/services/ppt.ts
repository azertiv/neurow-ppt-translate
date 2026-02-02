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
import { isNonTranslatable } from "../utils/text";

export interface ShapeTextTarget {
  kind: "shapeText";
  shapeId: string;
  shapeName: string;
  groupPath: string[];
  shapePath?: string;
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
  groupPath: string[];
  shapePath?: string;
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

function shapeKey(shapeId: string, groupPath: string[] = []): string {
  return groupPath.length ? `${groupPath.join("::")}::${shapeId}` : shapeId;
}

function shapeLabel(shapeName: string, shapeId: string, groupNamePath: string[] = []): string {
  const base = shapeName || shapeId;
  if (!groupNamePath.length) return base;
  return `${groupNamePath.join(" / ")} / ${base}`;
}

async function loadShapesCollection(
  context: PowerPoint.RequestContext,
  collection: any
): Promise<PowerPoint.Shape[]> {
  if (!collection) return [];
  try {
    collection.load("items/id,items/name,items/type");
    await context.sync();
    return (collection as any).items ?? [];
  } catch {
    return [];
  }
}

async function getGroupChildShapes(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  logger?: Logger,
  label?: string
): Promise<PowerPoint.Shape[]> {
  const candidates = [
    () => (shape as any).group?.shapes,
    () => (shape as any).shapes,
    () => (shape as any).groupItems
  ];

  for (const candidate of candidates) {
    try {
      const collection = candidate();
      const items = await loadShapesCollection(context, collection);
      if (items.length) return items;
    } catch {
      // try next candidate
    }
  }

  if (logger && label) {
    logger.log(`Group non accessible: ${label}`, "dim");
  }
  return [];
}

async function walkShapeItems(
  context: PowerPoint.RequestContext,
  items: PowerPoint.Shape[],
  groupPath: string[],
  groupNamePath: string[],
  ignore: RegExp | null,
  logger: Logger | undefined,
  visitor: (shape: PowerPoint.Shape, groupPath: string[], groupNamePath: string[]) => Promise<void>
): Promise<void> {
  for (const shape of items) {
    const shapeId = shape.id;
    const shapeName = shape.name ?? "";

    if (ignore && ignore.test(shapeName)) {
      logger?.log(`Ignoré: ${shapeLabel(shapeName, shapeId, groupNamePath)}`, "dim");
      continue;
    }

    if ((shape.type as any) === PowerPoint.ShapeType.group) {
      const label = shapeLabel(shapeName, shapeId, groupNamePath);
      const nextGroupPath = [...groupPath, shapeId];
      const nextGroupNamePath = [...groupNamePath, shapeName || shapeId];
      const children = await getGroupChildShapes(context, shape, logger, label);
      if (!children.length) continue;
      await walkShapeItems(context, children, nextGroupPath, nextGroupNamePath, ignore, logger, visitor);
      continue;
    }

    await visitor(shape, groupPath, groupNamePath);
  }
}

async function walkShapes(
  context: PowerPoint.RequestContext,
  collection: any,
  ignore: RegExp | null,
  logger: Logger | undefined,
  visitor: (shape: PowerPoint.Shape, groupPath: string[], groupNamePath: string[]) => Promise<void>
): Promise<void> {
  const items = await loadShapesCollection(context, collection);
  if (!items.length) return;
  await walkShapeItems(context, items, [], [], ignore, logger, visitor);
}

async function buildShapeIndex(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  logger?: Logger
): Promise<Map<string, PowerPoint.Shape>> {
  const map = new Map<string, PowerPoint.Shape>();
  await walkShapes(context, slide.shapes, null, logger, async (shape, groupPath) => {
    map.set(shapeKey(shape.id, groupPath), shape);
  });
  return map;
}

export async function analyzeScope(settings: Settings, logger?: Logger): Promise<SlideAnalysis[]> {
  const indices = await getSlideIndices(settings.scope);
  const ignore = compileIgnoreRegex(settings.ignoreRegex);

  const out: SlideAnalysis[] = [];

  for (const slideIndex of indices) {
    const analysis = await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(slideIndex);

      let textBoxes = 0;
      let tables = 0;
      let paragraphs = 0;
      let characters = 0;

      await walkShapes(context, slide.shapes, ignore, undefined, async (shape) => {
        if ((shape.type as any) === PowerPoint.ShapeType.table) {
          tables++;
          // Approximate: count later
          return;
        }

        const tf = shape.getTextFrameOrNullObject();
        tf.load("isNullObject,hasText");
        await context.sync();
        if (tf.isNullObject || !tf.hasText) return;

        tf.textRange.load("text");
        await context.sync();
        const text = tf.textRange.text ?? "";
        if (!text.trim()) return;
        textBoxes++;
        const ps = text.split("\n");
        paragraphs += ps.length;
        characters += text.length;
      });

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

    const shapeTextTargets: ShapeTextTarget[] = [];
    const tableCellTargets: TableCellTarget[] = [];

    await walkShapes(context, slide.shapes, ignore, logger, async (shape, groupPath, groupNamePath) => {
      const shapeId = shape.id;
      const shapeName = shape.name ?? "";
      const shapePath = shapeLabel(shapeName, shapeId, groupNamePath);
      const shapeRef = shapeKey(shapeId, groupPath);

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

          const id = `s${slideIndex}_shape${shapeRef}_cell${coords[i].r}_${coords[i].c}`;
          tableCellTargets.push({
            kind: "tableCell",
            shapeId,
            shapeName,
            groupPath: [...groupPath],
            shapePath,
            row: coords[i].r,
            col: coords[i].c,
            paragraphId: id,
            originalChars: text.length,
            runs: runSnapshots
          });
        }
        return;
      }

      const tf = shape.getTextFrameOrNullObject();
      tf.load("isNullObject,hasText");
      await context.sync();
      if (tf.isNullObject || !tf.hasText) return;

      const paragraphs = await extractShapeTextParagraphs(
        context,
        tf.textRange,
        slideIndex,
        shapeRef,
        shapeName
      );

      const useful = paragraphs.some((p) => p.runs.some((r) => r.text.trim().length > 0));
      if (!useful) return;

      shapeTextTargets.push({
        kind: "shapeText",
        shapeId,
        shapeName,
        groupPath: [...groupPath],
        shapePath,
        paragraphs
      });
    });

    logger?.log(
      `Slide ${slideIndex + 1}: ${shapeTextTargets.length} shape(s) texte, ${tableCellTargets.length} cellule(s) de table`,
      "dim"
    );

    return { slideIndex, shapeTextTargets, tableCellTargets };
  });
}

function buildTranslateItems(targets: SlideTargets): TranslateBatchItem[] {
  const items: TranslateBatchItem[] = [];
  const isSkippable = (runs: { text: string }[]) => runs.every((r) => isNonTranslatable(r.text));
  for (const t of targets.shapeTextTargets) {
    for (const p of t.paragraphs) {
      const runs = p.runs.map((r, idx) => ({ index: idx, text: r.text }));
      if (isSkippable(runs)) continue;
      items.push({
        paragraphId: p.id,
        originalChars: p.originalCharCount,
        runs
      });
    }
  }
  for (const t of targets.tableCellTargets) {
    const runs = t.runs.map((r, idx) => ({ index: idx, text: r.text }));
    if (isSkippable(runs)) continue;
    items.push({
      paragraphId: t.paragraphId,
      originalChars: t.originalChars,
      runs
    });
  }
  return items;
}

function chunkByChars<T extends { originalChars: number }>(
  items: T[],
  maxChars = 12000,
  maxItems = 60
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

function translateKey(item: TranslateBatchItem): string {
  return JSON.stringify(item.runs.map((r) => r.text));
}

async function translateChunks(
  chunks: TranslateBatchItem[][],
  settings: Settings,
  logger: Logger,
  abortSignal?: AbortSignal,
  concurrency = 3
): Promise<TranslationResult[]> {
  if (!chunks.length) return [];

  const results: TranslationResult[] = [];
  let cursor = 0;
  const workers = Array.from({ length: Math.min(concurrency, chunks.length) }, async () => {
    while (true) {
      if (abortSignal?.aborted) break;
      const i = cursor++;
      if (i >= chunks.length) break;
      logger.log(`OpenAI: lot ${i + 1}/${chunks.length}…`, "dim");
      const res = await translateBatch(chunks[i], settings);
      results.push(...res);
    }
  });

  await Promise.all(workers);
  return results;
}

export async function translateAndMaybeApplySlide(
  targets: SlideTargets,
  settings: Settings,
  logger: Logger,
  abortSignal?: AbortSignal,
  translationCache?: Map<string, { translatedRuns: { index: number; text: string }[] }>
): Promise<{ translated: number; preview: string }>
{
  const items = buildTranslateItems(targets);
  if (!items.length) return { translated: 0, preview: "" };

  const cache = translationCache ?? new Map<string, { translatedRuns: { index: number; text: string }[] }>();
  const pending: TranslateBatchItem[] = [];
  const keyToIds = new Map<string, string[]>();
  const idToKey = new Map<string, string>();
  const allResults: TranslationResult[] = [];

  for (const item of items) {
    const key = translateKey(item);
    const cached = cache.get(key);
    if (cached) {
      allResults.push({ paragraphId: item.paragraphId, translatedRuns: cached.translatedRuns });
      continue;
    }
    const ids = keyToIds.get(key);
    if (ids) {
      ids.push(item.paragraphId);
    } else {
      keyToIds.set(key, [item.paragraphId]);
      idToKey.set(item.paragraphId, key);
      pending.push(item);
    }
  }

  const chunks = chunkByChars(pending);
  const translated = await translateChunks(chunks, settings, logger, abortSignal);
  for (const r of translated) {
    const key = idToKey.get(r.paragraphId);
    if (!key) continue;
    cache.set(key, { translatedRuns: r.translatedRuns });
    const ids = keyToIds.get(key) ?? [];
    for (const id of ids) {
      allResults.push({ paragraphId: id, translatedRuns: r.translatedRuns });
    }
  }

  const map = new Map<string, TranslationResult>();
  for (const r of allResults) map.set(r.paragraphId, r);

  if (settings.mode === "preview") {
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

    return { translated: map.size, preview: previewParts.join("\n\n") };
  }

  // Apply
  await applySlideTranslations(targets.slideIndex, targets, map, settings, logger);

  return { translated: map.size, preview: "" };
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
    const shapeIndex = await buildShapeIndex(context, slide, logger);

    // Shapes (text frames)
    for (const st of targets.shapeTextTargets) {
      const key = shapeKey(st.shapeId, st.groupPath);
      const shape = shapeIndex.get(key) ?? slide.shapes.getItem(st.shapeId);
      if (!shape) {
        logger?.log(`Forme introuvable: ${st.shapePath || st.shapeName || st.shapeId}`, "dim");
        continue;
      }
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

      logger?.log(`Appliqué: ${st.shapePath || st.shapeName || st.shapeId}`, "dim");
    }

    // Tables
    for (const tc of targets.tableCellTargets) {
      const key = shapeKey(tc.shapeId, tc.groupPath);
      const shape = shapeIndex.get(key) ?? slide.shapes.getItem(tc.shapeId);
      if (!shape) {
        logger?.log(`Table introuvable: ${tc.shapePath || tc.shapeName || tc.shapeId}`, "dim");
        continue;
      }
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
  const translationCache = new Map<string, { translatedRuns: { index: number; text: string }[] }>();

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

    const res = await translateAndMaybeApplySlide(targets, settings, logger, abortSignal, translationCache);
    translatedTotal += res.translated;
    if (!preview && res.preview) preview = res.preview;

    onProgress(i + 1, total, `Slide ${slideIndex + 1} terminé`);
  }

  onProgress(total, total, abortSignal?.aborted ? "Annulé" : "Terminé");
  return { translated: translatedTotal, preview };
}
