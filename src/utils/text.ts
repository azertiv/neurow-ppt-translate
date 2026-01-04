export function parseGlossaryText(input: string): Record<string, string> {
  const out: Record<string, string> = {};
  const lines = input
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  for (const line of lines) {
    const idx = line.indexOf("=");
    if (idx <= 0) continue;
    const k = line.slice(0, idx).trim();
    const v = line.slice(idx + 1).trim();
    if (!k || !v) continue;
    out[k] = v;
  }
  return out;
}

export function preserveWhitespace(original: string, translated: string): string {
  const pre = original.match(/^\s+/)?.[0] ?? "";
  const suf = original.match(/\s+$/)?.[0] ?? "";
  return `${pre}${translated.trim()}${suf}`;
}

export function isNonTranslatable(text: string): boolean {
  if (!text) return true;
  if (/^\s+$/.test(text)) return true;
  if (/^[\s\d.,%+\-–—()\[\]{}<>:;!?/\\|@#^&*=~`'"€$£¥]+$/.test(text)) return true;
  return false;
}
