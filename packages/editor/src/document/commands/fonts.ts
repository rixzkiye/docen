/**
 * Font substitution mapping and missing font analyzer.
 * Provides fallback chains for Microsoft Office standard fonts when running across Linux, macOS, and Windows.
 */

export const FONT_SUBSTITUTIONS: Record<string, string[]> = {
  // Office 365 default fonts
  Aptos: ["Calibri", "Segoe UI", "Arial", "sans-serif"],
  "Aptos Display": ["Calibri Light", "Segoe UI Semibold", "Arial", "sans-serif"],
  // Classic Office fonts
  Calibri: ["Carlito", "Arial", "Helvetica", "Liberation Sans", "sans-serif"],
  "Calibri Light": ["Arial Light", "Segoe UI Light", "sans-serif"],
  Cambria: ["Caladea", "Georgia", "Times New Roman", "serif"],
  "Times New Roman": ["Tinos", "Times", "Nimbus Roman", "serif"],
  Arial: ["Liberation Sans", "Helvetica", "sans-serif"],
  Georgia: ["Garamond", "serif"],
  Consolas: ["Cousine", "Courier New", "monospace"],
  "Courier New": ["Courier", "monospace"],
  // CJK Fonts
  SimSun: ["Songti SC", "SimSun", "STSong", "serif"],
  SimHei: ["PingFang SC", "Heiti SC", "Microsoft YaHei", "SimHei", "sans-serif"],
  "Microsoft YaHei": ["PingFang SC", "Microsoft YaHei", "Hiragino Sans GB", "sans-serif"],
  FangSong: ["STFangsong", "FangSong", "serif"],
  KaiTi: ["STKaiti", "KaiTi", "serif"],
  // Symbolic fonts
  Symbol: ["Symbol", "Segoe UI Symbol"],
  Wingdings: ["Wingdings", "Segoe UI Emoji"],
  Webdings: ["Webdings", "Segoe UI Emoji"],
};

/**
 * Return an array of fallback font names for a given font family.
 */
export function resolveFontSubstitution(fontFamily: string): string[] {
  if (!fontFamily) return ["sans-serif"];
  const clean = fontFamily.trim().replace(/^["']|["']$/g, "");
  const found = FONT_SUBSTITUTIONS[clean];
  if (found) {
    return [clean, ...found];
  }
  return [clean, "sans-serif"];
}

/**
 * Return a CSS font-family fallback string (e.g. '"Aptos", "Calibri", "Segoe UI", sans-serif').
 */
export function getFallbackFontChain(fontFamily: string): string {
  const list = resolveFontSubstitution(fontFamily);
  return list
    .map((f) => (f === "sans-serif" || f === "serif" || f === "monospace" ? f : `"${f}"`))
    .join(", ");
}

/**
 * Detect fonts used in a document that are not present in a given list of available system fonts.
 */
export function detectMissingFonts(docFonts: string[], availableFonts?: string[]): string[] {
  if (!availableFonts || availableFonts.length === 0) return [];
  const availSet = new Set(availableFonts.map((f) => f.toLowerCase().trim()));
  const missing: string[] = [];

  for (const font of docFonts) {
    const clean = font.trim().replace(/^["']|["']$/g, "");
    if (!clean) continue;
    if (!availSet.has(clean.toLowerCase())) {
      missing.push(clean);
    }
  }

  return [...new Set(missing)];
}
