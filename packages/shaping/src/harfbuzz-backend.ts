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
    } else if (options?.direction === "ttb") {
      buffer.setDirection(hb.Direction.TTB);
    } else if (options?.direction === "btt") {
      buffer.setDirection(hb.Direction.BTT);
    } else if (options?.direction === "ltr") {
      buffer.setDirection(hb.Direction.LTR);
    }

    if (options?.script && typeof buffer.setScript === "function") {
      buffer.setScript(options.script);
    }
    if (options?.language && typeof buffer.setLanguage === "function") {
      buffer.setLanguage(options.language);
    }

    if (options?.variations && options.variations.length > 0) {
      entry.font.setVariations(options.variations.map((v) => new hb.Variation(v.tag, v.value)));
    } else {
      entry.font.setVariations([]);
    }

    let features: any[] | undefined;
    if (options?.features && options.features.length > 0) {
      features = options.features.map((f) => new hb.Feature(f.tag, f.value ?? 1));
    }

    buffer.guessSegmentProperties();
    hb.shape(entry.font, buffer, features);

    const hbGlyphs = buffer.getGlyphInfosAndPositions();
    const glyphs: ShapedGlyph[] = [];
    const isVertical = options?.direction === "ttb" || options?.direction === "btt";
    let totalAdvance = 0;

    for (const g of hbGlyphs) {
      const xAdv = g.xAdvance ?? 0;
      const yAdv = g.yAdvance ?? 0;
      glyphs.push({
        glyphId: g.codepoint,
        cluster: g.cluster,
        xAdvance: xAdv,
        yAdvance: yAdv,
        xOffset: g.xOffset ?? 0,
        yOffset: g.yOffset ?? 0,
      });
      totalAdvance += isVertical ? Math.abs(yAdv) : xAdv;
    }

    return { glyphs, totalAdvance };
  }
}
