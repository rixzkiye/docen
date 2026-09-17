import type {
  FontMetrics,
  PathCommand,
  ShapedGlyph,
  ShapingBackend,
  ShapingOptions,
  ShapingResult,
} from "./types.js";
import { getShapingWasm } from "./wasm-loader.js";

const textEncoder = new TextEncoder();

export class RustybuzzBackend implements ShapingBackend {
  readonly id = "rustybuzz";

  registerFont(fontData: Uint8Array): number {
    const wasm = getShapingWasm();
    const ptr = wasm.alloc(fontData.length);
    new Uint8Array(wasm.memory.buffer, ptr, fontData.length).set(fontData);

    const fontId = wasm.register_font(ptr, fontData.length);
    wasm.dealloc(ptr, fontData.length);

    if (fontId < 0) {
      throw new Error(`Failed to register font with error code ${fontId}`);
    }
    return fontId;
  }

  dropFont(fontId: number): void {
    const wasm = getShapingWasm();
    wasm.drop_font(fontId);
  }

  getFontMetrics(fontId: number): FontMetrics {
    const wasm = getShapingWasm();
    const code = wasm.get_font_metrics(fontId);
    if (code !== 0) {
      throw new Error(`Failed to get font metrics for font ${fontId}, error: ${code}`);
    }

    const ptr = wasm.get_metrics_buffer_ptr();
    const metricsView = new Float32Array(wasm.memory.buffer, ptr, 6);
    return {
      unitsPerEm: metricsView[0],
      ascender: metricsView[1],
      descender: metricsView[2],
      lineGap: metricsView[3],
      capHeight: metricsView[4],
      xHeight: metricsView[5],
    };
  }

  shape(fontId: number, text: string, options?: ShapingOptions): ShapingResult {
    if (!text) {
      return { glyphs: [], totalAdvance: 0 };
    }

    const wasm = getShapingWasm();
    const encodedText = textEncoder.encode(text);
    const textPtr = wasm.alloc(encodedText.length);
    new Uint8Array(wasm.memory.buffer, textPtr, encodedText.length).set(encodedText);

    let direction = 0;
    if (options?.direction === "rtl") {
      direction = 1;
    }

    let scriptTag = 0;
    if (options?.script && options.script.length === 4) {
      for (let i = 0; i < 4; i++) {
        scriptTag = (scriptTag << 8) | options.script.charCodeAt(i);
      }
    }

    let langPtr = 0;
    let langLen = 0;
    if (options?.language) {
      const encodedLang = textEncoder.encode(options.language);
      langPtr = wasm.alloc(encodedLang.length);
      langLen = encodedLang.length;
      new Uint8Array(wasm.memory.buffer, langPtr, langLen).set(encodedLang);
    }

    const glyphCount = wasm.shape_text(
      fontId,
      textPtr,
      encodedText.length,
      direction,
      scriptTag,
      langPtr,
      langLen,
    );

    wasm.dealloc(textPtr, encodedText.length);
    if (langPtr > 0) {
      wasm.dealloc(langPtr, langLen);
    }

    if (glyphCount < 0) {
      throw new Error(`Shaping failed with error code ${glyphCount}`);
    }

    const outPtr = wasm.get_shape_buffer_ptr();
    const floats = new Float32Array(wasm.memory.buffer, outPtr, glyphCount * 6);

    const glyphs: ShapedGlyph[] = [];
    let totalAdvance = 0;

    for (let i = 0; i < glyphCount; i++) {
      const offset = i * 6;
      const xAdvance = floats[offset + 2];
      glyphs.push({
        glyphId: floats[offset],
        cluster: floats[offset + 1],
        xAdvance,
        yAdvance: floats[offset + 3],
        xOffset: floats[offset + 4],
        yOffset: floats[offset + 5],
      });
      totalAdvance += xAdvance;
    }

    return { glyphs, totalAdvance };
  }

  getGlyphOutline(fontId: number, glyphId: number): readonly PathCommand[] {
    const wasm = getShapingWasm();
    const count = wasm.get_glyph_outline(fontId, glyphId);
    if (count <= 0) {
      return [];
    }

    const ptr = wasm.get_outline_buffer_ptr();
    const floats = new Float32Array(wasm.memory.buffer, ptr, count);

    const commands: PathCommand[] = [];
    let idx = 0;
    while (idx < count) {
      const tag = floats[idx++];
      if (tag === 0) {
        commands.push({ type: "M", x: floats[idx++], y: floats[idx++] });
      } else if (tag === 1) {
        commands.push({ type: "L", x: floats[idx++], y: floats[idx++] });
      } else if (tag === 2) {
        commands.push({
          type: "Q",
          cx: floats[idx++],
          cy: floats[idx++],
          x: floats[idx++],
          y: floats[idx++],
        });
      } else if (tag === 3) {
        commands.push({
          type: "C",
          cx1: floats[idx++],
          cy1: floats[idx++],
          cx2: floats[idx++],
          cy2: floats[idx++],
          x: floats[idx++],
          y: floats[idx++],
        });
      } else if (tag === 4) {
        commands.push({ type: "Z" });
      }
    }

    return commands;
  }

  getFontName(fontId: number, nameId: number): string | undefined {
    const wasm = getShapingWasm();
    const len = wasm.get_font_name(fontId, nameId);
    if (len <= 0) return undefined;
    const ptr = wasm.get_string_buffer_ptr();
    const bytes = new Uint8Array(wasm.memory.buffer, ptr, len);
    return new TextDecoder().decode(bytes);
  }

  getGlyphCount(fontId: number): number {
    const wasm = getShapingWasm();
    const count = wasm.get_glyph_count(fontId);
    return count >= 0 ? count : 0;
  }

  subset(_fontId: number, _glyphIds: readonly number[]): Uint8Array {
    return new Uint8Array(0);
  }
}
