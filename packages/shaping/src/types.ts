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
 * Font vertical header metrics (vhea table) in font design units.
 */
export interface VerticalFontMetrics {
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
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
  readonly vertical?: VerticalFontMetrics;
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
 * OpenType variation axis metadata.
 */
export interface FontAxis {
  readonly tag: string;
  readonly min: number;
  readonly max: number;
  readonly default: number;
}

/**
 * Single font variation axis coordinate setting (e.g. wght: 700).
 */
export interface FontVariationSetting {
  readonly tag: string;
  readonly value: number;
}

/**
 * Single OpenType feature setting (e.g. kern: 1, smcp: 1, liga: 0).
 */
export interface FontFeatureSetting {
  readonly tag: string;
  readonly value?: number;
}

/**
 * Options configuring text shaping behavior.
 */
export interface ShapingOptions {
  readonly direction?: "ltr" | "rtl" | "ttb" | "btt" | "auto";
  readonly script?: string; // 4-char ISO 15924 tag, e.g. "Latn", "Arab", "Hebr", "Thai", "Deva", "Khmr", "Hani"
  readonly language?: string; // BCP 47 or OpenType language tag
  readonly features?: readonly FontFeatureSetting[];
  readonly variations?: readonly FontVariationSetting[];
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
  readonly fsType?: number;
  readonly isEmbeddingAllowed?: boolean;
  readonly isSubsettingAllowed?: boolean;
  readonly axes?: readonly FontAxis[];
}

export const FS_TYPE_RESTRICTED = 0x0002;
export const FS_TYPE_PREVIEW_PRINT = 0x0004;
export const FS_TYPE_EDITABLE = 0x0008;
export const FS_TYPE_NO_SUBSETTING = 0x0100;
export const FS_TYPE_BITMAP_ONLY = 0x0200;

export function isFontEmbeddingAllowed(fsType: number): boolean {
  return (fsType & FS_TYPE_RESTRICTED) === 0;
}

export function isFontSubsettingAllowed(fsType: number): boolean {
  return isFontEmbeddingAllowed(fsType) && (fsType & FS_TYPE_NO_SUBSETTING) === 0;
}

/**
 * Common abstraction for text shaping engines (rustybuzz WASM, HarfBuzz oracle, Canvas fallback).
 */
export interface ShapingBackend {
  readonly id: string;
  registerFont(fontData: Uint8Array, fontIndex?: number): number;
  dropFont(fontId: number): void;
  getFontMetrics(fontId: number, variations?: readonly FontVariationSetting[]): FontMetrics;
  shape(fontId: number, text: string, options?: ShapingOptions): ShapingResult;
  getGlyphOutline?(
    fontId: number,
    glyphId: number,
    variations?: readonly FontVariationSetting[],
  ): readonly PathCommand[];
  getFontAxes?(fontId: number): readonly FontAxis[];
  getFontName?(fontId: number, nameId: number): string | undefined;
  getGlyphCount?(fontId: number): number;
  getFontFsType?(fontId: number): number;
}
