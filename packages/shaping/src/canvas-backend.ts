import type {
  FontMetrics,
  ShapedGlyph,
  ShapingBackend,
  ShapingOptions,
  ShapingResult,
} from "./types.js";

/**
 * Fallback backend simulating simple unshaped text advance measurement.
 */
export class CanvasBackend implements ShapingBackend {
  readonly id = "canvas";
  private nextFontId = 1;

  registerFont(_fontData: Uint8Array): number {
    return this.nextFontId++;
  }

  dropFont(_fontId: number): void {}

  getFontMetrics(_fontId: number): FontMetrics {
    return {
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      lineGap: 0,
      capHeight: 700,
      xHeight: 500,
    };
  }

  shape(_fontId: number, text: string, _options?: ShapingOptions): ShapingResult {
    if (!text) {
      return { glyphs: [], totalAdvance: 0 };
    }

    const glyphs: ShapedGlyph[] = [];
    let totalAdvance = 0;

    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      // Rough Latin advance approximation for fallback
      const adv = code < 128 ? 600 : 1000;
      glyphs.push({
        glyphId: code,
        cluster: i,
        xAdvance: adv,
        yAdvance: 0,
        xOffset: 0,
        yOffset: 0,
      });
      totalAdvance += adv;
    }

    return { glyphs, totalAdvance };
  }
}
