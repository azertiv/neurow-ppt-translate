import type { FontSnapshot, Paragraph, ParagraphFormatSnapshot, Run, Settings } from "../utils/types";
import { isNonTranslatable, preserveWhitespace } from "../utils/text";

function fontKey(f: FontSnapshot): string {
  return [
    f.allCaps,
    f.bold,
    f.color,
    f.doubleStrikethrough,
    f.italic,
    f.name,
    f.size,
    f.smallCaps,
    f.strikethrough,
    f.subscript,
    f.superscript,
    f.underline
  ].join("|");
}

export function isApiSupported(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", version);
  } catch {
    return false;
  }
}

export function safeFontSnapshot(font: PowerPoint.ShapeFont): FontSnapshot {
  return {
    allCaps: Boolean(font.allCaps ?? false),
    bold: Boolean(font.bold ?? false),
    color: (font.color as any) ?? "#FFFFFF",
    doubleStrikethrough: Boolean(font.doubleStrikethrough ?? false),
    italic: Boolean(font.italic ?? false),
    name: (font.name as any) ?? "",
    size: Number(font.size ?? 12),
    smallCaps: Boolean(font.smallCaps ?? false),
    strikethrough: Boolean(font.strikethrough ?? false),
    subscript: Boolean(font.subscript ?? false),
    superscript: Boolean(font.superscript ?? false),
    underline: (font.underline as any) ?? "None"
  };
}

function defaultFontSnapshot(): FontSnapshot {
  return {
    allCaps: false,
    bold: false,
    color: "#FFFFFF",
    doubleStrikethrough: false,
    italic: false,
    name: "",
    size: 12,
    smallCaps: false,
    strikethrough: false,
    subscript: false,
    superscript: false,
    underline: "None"
  };
}

export async function extractParagraphFormat(
  context: PowerPoint.RequestContext,
  range: PowerPoint.TextRange
): Promise<ParagraphFormatSnapshot | undefined> {
  const out: ParagraphFormatSnapshot = {};
  const wantsBulletDetails = isApiSupported("1.10");

  try {
    const props = [
      "paragraphFormat/indentLevel",
      "paragraphFormat/bulletFormat/visible",
      "paragraphFormat/horizontalAlignment"
    ];
    if (wantsBulletDetails) {
      props.push("paragraphFormat/bulletFormat/type", "paragraphFormat/bulletFormat/style");
    }
    range.load(props.join(","));
    await context.sync();
  } catch {
    try {
      range.load("paragraphFormat/indentLevel,paragraphFormat/bulletFormat/visible");
      await context.sync();
    } catch {
      // ignore
    }
    try {
      range.load("paragraphFormat/horizontalAlignment");
      await context.sync();
    } catch {
      // ignore
    }
    if (wantsBulletDetails) {
      try {
        range.load("paragraphFormat/bulletFormat/type,paragraphFormat/bulletFormat/style");
        await context.sync();
      } catch {
        // ignore
      }
    }
  }

  try {
    out.indentLevel = range.paragraphFormat.indentLevel ?? undefined;
  } catch {
    // ignore
  }
  try {
    out.bulletVisible = range.paragraphFormat.bulletFormat.visible ?? undefined;
  } catch {
    // ignore
  }
  try {
    out.horizontalAlignment = range.paragraphFormat.horizontalAlignment as any;
  } catch {
    // ignore
  }
  if (wantsBulletDetails) {
    try {
      out.bulletType = range.paragraphFormat.bulletFormat.type as any;
    } catch {
      // ignore
    }
    try {
      out.bulletStyle = range.paragraphFormat.bulletFormat.style as any;
    } catch {
      // ignore
    }
  }

  return Object.keys(out).length ? out : undefined;
}

export async function extractShapeTextParagraphs(
  context: PowerPoint.RequestContext,
  textRange: PowerPoint.TextRange,
  slideIndex: number,
  shapeId: string,
  knownText?: string
): Promise<Paragraph[]> {
  let text = knownText;
  if (text === undefined) {
    textRange.load("text");
    await context.sync();
    text = textRange.text ?? "";
  }
  if (!text) return [];

  const paras = text.split("\n");
  const out: Paragraph[] = [];
  let cursor = 0;
  for (let p = 0; p < paras.length; p++) {
    const pText = paras[p];
    const pStart = cursor;
    const pRange = textRange.getSubstring(pStart, pText.length);

    const pFormat = await extractParagraphFormat(context, pRange);
    const runs = await extractRunsByFont(context, textRange, pStart, pText);

    out.push({
      id: `s${slideIndex}_shape${shapeId}_p${p}`,
      originalCharCount: pText.length,
      runs,
      paragraphFormat: pFormat
    });

    cursor += pText.length;
    if (p < paras.length - 1) cursor += 1;
  }

  return out;
}

async function extractRunsByFont(
  context: PowerPoint.RequestContext,
  fullRange: PowerPoint.TextRange,
  paragraphStart: number,
  paragraphText: string
): Promise<Run[]> {
  const n = paragraphText.length;
  if (n === 0) return [{ text: "", font: defaultFontSnapshot() }];

  const MAX_CHAR_SCAN = 1500;
  if (n <= MAX_CHAR_SCAN) {
    return extractRunsByCharScan(context, fullRange, paragraphStart, paragraphText);
  }
  return extractRunsByBinarySplit(context, fullRange, paragraphStart, paragraphText);
}

async function extractRunsByCharScan(
  context: PowerPoint.RequestContext,
  fullRange: PowerPoint.TextRange,
  paragraphStart: number,
  paragraphText: string
): Promise<Run[]> {
  const n = paragraphText.length;
  const charRanges: PowerPoint.TextRange[] = [];
  for (let i = 0; i < n; i++) {
    const r = fullRange.getSubstring(paragraphStart + i, 1);
    r.load(
      "font/allCaps,font/bold,font/color,font/doubleStrikethrough,font/italic,font/name,font/size,font/smallCaps,font/strikethrough,font/subscript,font/superscript,font/underline"
    );
    charRanges.push(r);
  }
  await context.sync();

  const runs: Run[] = [];
  let curFont: FontSnapshot | null = null;
  let curText = "";

  for (let i = 0; i < n; i++) {
    const f = safeFontSnapshot(charRanges[i].font);
    if (!curFont) {
      curFont = f;
      curText = paragraphText[i];
      continue;
    }
    if (fontKey(curFont) === fontKey(f)) {
      curText += paragraphText[i];
    } else {
      runs.push({ text: curText, font: curFont });
      curFont = f;
      curText = paragraphText[i];
    }
  }
  if (curFont) runs.push({ text: curText, font: curFont });
  return runs;
}

async function extractRunsByBinarySplit(
  context: PowerPoint.RequestContext,
  fullRange: PowerPoint.TextRange,
  paragraphStart: number,
  paragraphText: string
): Promise<Run[]> {
  const spans: { start: number; length: number; font: FontSnapshot }[] = [];

  async function split(start: number, length: number): Promise<void> {
    const r = fullRange.getSubstring(start, length);
    r.load(
      "font/allCaps,font/bold,font/color,font/doubleStrikethrough,font/italic,font/name,font/size,font/smallCaps,font/strikethrough,font/subscript,font/superscript,font/underline"
    );
    await context.sync();

    const raw: any = {
      allCaps: r.font.allCaps,
      bold: r.font.bold,
      color: r.font.color,
      doubleStrikethrough: r.font.doubleStrikethrough,
      italic: r.font.italic,
      name: r.font.name,
      size: r.font.size,
      smallCaps: r.font.smallCaps,
      strikethrough: r.font.strikethrough,
      subscript: r.font.subscript,
      superscript: r.font.superscript,
      underline: r.font.underline
    };

    const mixed = Object.values(raw).some((v) => v === null || v === undefined);
    if (!mixed || length <= 1) {
      spans.push({ start, length, font: safeFontSnapshot(r.font) });
      return;
    }

    const mid = Math.floor(length / 2);
    await split(start, mid);
    await split(start + mid, length - mid);
  }

  await split(paragraphStart, paragraphText.length);

  spans.sort((a, b) => a.start - b.start);

  const out: Run[] = [];
  let last: Run | null = null;
  for (const s of spans) {
    const localStart = s.start - paragraphStart;
    const piece = paragraphText.slice(localStart, localStart + s.length);
    if (!last) {
      last = { text: piece, font: s.font };
      continue;
    }
    if (fontKey(last.font) === fontKey(s.font)) {
      last.text += piece;
    } else {
      out.push(last);
      last = { text: piece, font: s.font };
    }
  }
  if (last) out.push(last);
  return out;
}

export function applyRunTranslations(
  paragraph: Paragraph,
  translatedRuns: { index: number; text: string }[]
): Paragraph {
  const map = new Map<number, string>(translatedRuns.map((r) => [r.index, r.text]));
  const nextRuns: Run[] = paragraph.runs.map((r, idx) => {
    if (isNonTranslatable(r.text)) return r;
    const t = map.get(idx);
    if (typeof t !== "string") return r;
    return { ...r, text: preserveWhitespace(r.text, t) };
  });
  return { ...paragraph, runs: nextRuns };
}

export function queueParagraphFormats(
  fullRange: PowerPoint.TextRange,
  paragraphs: { start: number; length: number; format?: ParagraphFormatSnapshot }[]
) {
  for (const p of paragraphs) {
    if (!p.format) continue;
    const r = fullRange.getSubstring(p.start, p.length);

    try {
      if (p.format.horizontalAlignment !== undefined) {
        (r.paragraphFormat as any).horizontalAlignment = p.format.horizontalAlignment;
      }
    } catch {
      // ignore
    }

    try {
      if (p.format.indentLevel !== undefined) {
        (r.paragraphFormat as any).indentLevel = p.format.indentLevel;
      }
    } catch {
      // ignore
    }

    try {
      if (p.format.bulletVisible !== undefined) {
        (r.paragraphFormat.bulletFormat as any).visible = p.format.bulletVisible;
      }
    } catch {
      // ignore
    }

    if (isApiSupported("1.10")) {
      try {
        if (p.format.bulletType !== undefined && p.format.bulletType !== null) {
          (r.paragraphFormat.bulletFormat as any).type = p.format.bulletType;
        }
      } catch {
        // ignore
      }
      try {
        if (p.format.bulletStyle !== undefined && p.format.bulletStyle !== null) {
          (r.paragraphFormat.bulletFormat as any).style = p.format.bulletStyle;
        }
      } catch {
        // ignore
      }
    }
  }
}

export function queueFontRuns(
  fullRange: PowerPoint.TextRange,
  runs: { start: number; length: number; font: FontSnapshot }[],
  settings: Settings
) {
  for (const run of runs) {
    if (run.length <= 0) continue;
    const r = fullRange.getSubstring(run.start, run.length);

    r.font.allCaps = run.font.allCaps;
    r.font.bold = run.font.bold;
    r.font.color = run.font.color;
    r.font.doubleStrikethrough = run.font.doubleStrikethrough;
    r.font.italic = run.font.italic;
    if (run.font.name) r.font.name = run.font.name;
    if (run.font.size) r.font.size = run.font.size;
    r.font.smallCaps = run.font.smallCaps;
    r.font.strikethrough = run.font.strikethrough;
    r.font.subscript = run.font.subscript;
    r.font.superscript = run.font.superscript;

    if (settings.applyUnderline) {
      try {
        r.font.underline = run.font.underline as any;
      } catch {
        // ignore
      }
    }
  }
}
