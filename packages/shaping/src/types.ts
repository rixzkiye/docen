/**
 * A single shaped glyph resulting from OpenType shaping.
 */
export interface ShapedGlyph {
  readonly glyphId: number;
  readonly cluster: number;
  readonly xAdvance: number;
  readonly yAdvance: number;
  readonly xOffset: number;
  readonly yOffset: number;
}

/**
 * Result of shaping a text run.
 */
export interface ShapingResult {
  readonly glyphs: readonly ShapedGlyph[];
  readonly totalAdvance: number;
}

/**
 * Font horizontal and vertical metrics in font design units.
 */
export interface FontMetrics {
  readonly unitsPerEm: number;
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
  readonly capHeight: number;
  readonly xHeight: number;
}

/**
 * Normalized 2D vector path commands for glyph outlines.
 */
export type PathCommand =
  | { readonly type: "M"; readonly x: number; readonly y: number }
  | { readonly type: "L"; readonly x: number; readonly y: number }
  | {
      readonly type: "Q";
      readonly cx: number;
      readonly cy: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly type: "C";
      readonly cx1: number;
      readonly cy1: number;
      readonly cx2: number;
      readonly cy2: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly type: "Z" };

/**
 * Options configuring text shaping behavior.
 */
export interface ShapingOptions {
  readonly direction?: "ltr" | "rtl";
  readonly script?: string; // 4-char ISO 15924 tag, e.g. "Latn", "Arab"
  readonly language?: string; // BCP 47 or OpenType language tag
  readonly features?: readonly { readonly tag: string; readonly value: number }[];
  readonly size?: number;
}

/**
 * Font identity metadata extracted from font name and maxp tables.
 */
export interface FontIdentity {
  readonly familyName?: string;
  readonly postscriptName?: string;
  readonly fullName?: string;
  readonly subfamilyName?: string;
  readonly unitsPerEm: number;
  readonly glyphCount: number;
}

/**
 * Common abstraction for text shaping engines (rustybuzz WASM, HarfBuzz oracle, Canvas fallback).
 */
export interface ShapingBackend {
  readonly id: string;
  registerFont(fontData: Uint8Array, fontIndex?: number): number;
  dropFont(fontId: number): void;
  getFontMetrics(fontId: number): FontMetrics;
  shape(fontId: number, text: string, options?: ShapingOptions): ShapingResult;
  getGlyphOutline?(fontId: number, glyphId: number): readonly PathCommand[];
  getFontName?(fontId: number, nameId: number): string | undefined;
  getGlyphCount?(fontId: number): number;
  subset?(fontId: number, glyphIds: readonly number[]): Uint8Array;
}
