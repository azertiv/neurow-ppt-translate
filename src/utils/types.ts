export type Scope = "current" | "all";
export type Mode = "apply" | "preview";

export interface Settings {
  apiKey: string;
  fromLang: string; // code or "auto"
  toLang: string; // code
  scope: Scope;
  mode: Mode;
  keepLineBreaks: boolean;
  fitToLength: boolean;
  fitStrength: number; // 0-100
  glossary: Record<string, string>;
  ignoreRegex: string;
  applyUnderline: boolean;
}

export interface FontSnapshot {
  allCaps: boolean;
  bold: boolean;
  color: string;
  doubleStrikethrough: boolean;
  italic: boolean;
  name: string;
  size: number;
  smallCaps: boolean;
  strikethrough: boolean;
  subscript: boolean;
  superscript: boolean;
  underline: string;
}

export interface Run {
  text: string;
  font: FontSnapshot;
}

export interface ParagraphFormatSnapshot {
  horizontalAlignment?: string;
  indentLevel?: number;
  bulletVisible?: boolean;
  bulletType?: string | null;
  bulletStyle?: string | null;
}

export interface Paragraph {
  id: string;
  originalCharCount: number;
  runs: Run[];
  paragraphFormat?: ParagraphFormatSnapshot;
}

export interface TranslationResult {
  paragraphId: string;
  translatedRuns: { index: number; text: string }[];
}

export interface SlideAnalysis {
  slideIndex: number;
  textBoxes: number;
  tables: number;
  paragraphs: number;
  characters: number;
}
