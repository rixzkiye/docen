import * as hb from "harfbuzzjs";

import type {
  FontMetrics,
  ShapedGlyph,
  ShapingBackend,
  ShapingOptions,
  ShapingResult,
} from "./types.js";

interface HarfBuzzFontEntry {
  readonly blob: any;
  readonly face: any;
  readonly font: any;
}

export class HarfbuzzBackend implements ShapingBackend {
  readonly id = "harfbuzz";
  private nextFontId = 1;
  private readonly fonts = new Map<number, HarfBuzzFontEntry>();

  registerFont(fontData: Uint8Array): number {
    const id = this.nextFontId++;
    const blob = new hb.Blob(fontData);
    const face = new hb.Face(blob, 0);
    const font = new hb.Font(face);
    this.fonts.set(id, { blob, face, font });
    return id;
  }

  dropFont(fontId: number): void {
    const entry = this.fonts.get(fontId);
    if (entry) {
      if (typeof entry.font?.destroy === "function") entry.font.destroy();
      if (typeof entry.face?.destroy === "function") entry.face.destroy();
      if (typeof entry.blob?.destroy === "function") entry.blob.destroy();
      this.fonts.delete(fontId);
    }
  }

  getFontMetrics(fontId: number): FontMetrics {
    const entry = this.fonts.get(fontId);
    if (!entry) {
      throw new Error(`Font ${fontId} not found in HarfbuzzBackend`);
    }

    const upem = entry.face.getUpem();
    // Default metrics extraction from font if available
    return {
      unitsPerEm: upem,
      ascender: upem * 0.8,
      descender: -upem * 0.2,
      lineGap: 0,
      capHeight: upem * 0.7,
      xHeight: upem * 0.5,
    };
  }

  shape(fontId: number, text: string, options?: ShapingOptions): ShapingResult {
    const entry = this.fonts.get(fontId);
    if (!entry) {
      throw new Error(`Font ${fontId} not found in HarfbuzzBackend`);
    }
    if (!text) {
      return { glyphs: [], totalAdvance: 0 };
    }

    const buffer = new hb.Buffer();
    buffer.addText(text);

    if (options?.direction === "rtl") {
      buffer.setDirection(hb.Direction.RTL);
    } else {
      buffer.setDirection(hb.Direction.LTR);
    }

    if (options?.script && typeof buffer.setScript === "function") {
      buffer.setScript(options.script);
    }
    if (options?.language && typeof buffer.setLanguage === "function") {
      buffer.setLanguage(options.language);
    }

    buffer.guessSegmentProperties();
    hb.shape(entry.font, buffer);

    const hbGlyphs = buffer.getGlyphInfosAndPositions();
    const glyphs: ShapedGlyph[] = [];
    let totalAdvance = 0;

    for (const g of hbGlyphs) {
      const xAdv = g.xAdvance ?? 0;
      glyphs.push({
        glyphId: g.codepoint,
        cluster: g.cluster,
        xAdvance: xAdv,
        yAdvance: g.yAdvance ?? 0,
        xOffset: g.xOffset ?? 0,
        yOffset: g.yOffset ?? 0,
      });
      totalAdvance += xAdv;
    }

    return { glyphs, totalAdvance };
  }
}
