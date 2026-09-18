import type {
  FontAxis,
  FontMetrics,
  FontVariationSetting,
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

  registerFont(fontData: Uint8Array, fontIndex = 0): number {
    const wasm = getShapingWasm();
    const ptr = wasm.alloc(fontData.length);
    new Uint8Array(wasm.memory.buffer, ptr, fontData.length).set(fontData);

    // fontIndex selects a face inside a .ttc/.otc collection (0 = plain sfnt).
    const fontId = wasm.register_font_index(ptr, fontData.length, fontIndex);
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

  getFontMetrics(fontId: number, variations?: readonly FontVariationSetting[]): FontMetrics {
    const wasm = getShapingWasm();
    let code: number;
    let varPtr = 0;
    let varBytes = 0;
    if (variations && variations.length > 0) {
      varBytes = variations.length * 8;
      varPtr = wasm.alloc(varBytes);
      const view = new DataView(wasm.memory.buffer, varPtr, varBytes);
      for (let i = 0; i < variations.length; i++) {
        const v = variations[i]!;
        const off = i * 8;
        for (let c = 0; c < 4; c++) {
          view.setUint8(off + c, v.tag.charCodeAt(c) || 32);
        }
        view.setFloat32(off + 4, v.value, true);
      }
      code = wasm.get_font_metrics_var(fontId, varPtr, variations.length);
      wasm.dealloc(varPtr, varBytes);
    } else {
      code = wasm.get_font_metrics(fontId);
    }

    if (code !== 0) {
      throw new Error(`Failed to get font metrics for font ${fontId}, error: ${code}`);
    }

    const ptr = wasm.get_metrics_buffer_ptr();
    const metricsView = new Float32Array(wasm.memory.buffer, ptr, 12);
    const hasVertical = metricsView[6] === 1.0;
    return {
      unitsPerEm: metricsView[0]!,
      ascender: metricsView[1]!,
      descender: metricsView[2]!,
      lineGap: metricsView[3]!,
      capHeight: metricsView[4]!,
      xHeight: metricsView[5]!,
      // OS/2 usWinAscent/usWinDescent — Word's line-height formula inputs.
      winAscent: metricsView[10]!,
      winDescent: metricsView[11]!,
      ...(hasVertical
        ? {
            vertical: {
              ascender: metricsView[7]!,
              descender: metricsView[8]!,
              lineGap: metricsView[9]!,
            },
          }
        : {}),
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

    let direction = 4; // Auto / guess by default
    if (options?.direction === "ltr") {
      direction = 0;
    } else if (options?.direction === "rtl") {
      direction = 1;
    } else if (options?.direction === "ttb") {
      direction = 2;
    } else if (options?.direction === "btt") {
      direction = 3;
    } else if (options?.direction === "auto") {
      direction = 4;
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

    let featPtr = 0;
    let featCount = 0;
    let featBytes = 0;
    if (options?.features && options.features.length > 0) {
      featCount = options.features.length;
      featBytes = featCount * 8;
      featPtr = wasm.alloc(featBytes);
      const view = new DataView(wasm.memory.buffer, featPtr, featBytes);
      for (let i = 0; i < featCount; i++) {
        const f = options.features[i]!;
        const off = i * 8;
        for (let c = 0; c < 4; c++) {
          view.setUint8(off + c, f.tag.charCodeAt(c) || 32);
        }
        view.setUint32(off + 4, f.value ?? 1, true);
      }
    }

    let varPtr = 0;
    let varCount = 0;
    let varBytes = 0;
    if (options?.variations && options.variations.length > 0) {
      varCount = options.variations.length;
      varBytes = varCount * 8;
      varPtr = wasm.alloc(varBytes);
      const view = new DataView(wasm.memory.buffer, varPtr, varBytes);
      for (let i = 0; i < varCount; i++) {
        const v = options.variations[i]!;
        const off = i * 8;
        for (let c = 0; c < 4; c++) {
          view.setUint8(off + c, v.tag.charCodeAt(c) || 32);
        }
        view.setFloat32(off + 4, v.value, true);
      }
    }

    const glyphCount = wasm.shape_text(
      fontId,
      textPtr,
      encodedText.length,
      direction,
      scriptTag,
      langPtr,
      langLen,
      featPtr,
      featCount,
      varPtr,
      varCount,
    );

    wasm.dealloc(textPtr, encodedText.length);
    if (langPtr > 0) wasm.dealloc(langPtr, langLen);
    if (featPtr > 0) wasm.dealloc(featPtr, featBytes);
    if (varPtr > 0) wasm.dealloc(varPtr, varBytes);

    if (glyphCount < 0) {
      throw new Error(`Shaping failed with error code ${glyphCount}`);
    }

    const outPtr = wasm.get_shape_buffer_ptr();
    const floats = new Float32Array(wasm.memory.buffer, outPtr, glyphCount * 6);

    const glyphs: ShapedGlyph[] = [];
    const isVertical = options?.direction === "ttb" || options?.direction === "btt";
    let totalAdvance = 0;

    for (let i = 0; i < glyphCount; i++) {
      const offset = i * 6;
      const xAdvance = floats[offset + 2]!;
      const yAdvance = floats[offset + 3]!;
      glyphs.push({
        glyphId: floats[offset]!,
        cluster: floats[offset + 1]!,
        xAdvance,
        yAdvance,
        xOffset: floats[offset + 4]!,
        yOffset: floats[offset + 5]!,
      });
      totalAdvance += isVertical ? Math.abs(yAdvance) : xAdvance;
    }

    return { glyphs, totalAdvance };
  }

  getGlyphOutline(
    fontId: number,
    glyphId: number,
    variations?: readonly FontVariationSetting[],
  ): readonly PathCommand[] {
    const wasm = getShapingWasm();
    let count: number;
    if (variations && variations.length > 0) {
      const varBytes = variations.length * 8;
      const varPtr = wasm.alloc(varBytes);
      const view = new DataView(wasm.memory.buffer, varPtr, varBytes);
      for (let i = 0; i < variations.length; i++) {
        const v = variations[i]!;
        const off = i * 8;
        for (let c = 0; c < 4; c++) {
          view.setUint8(off + c, v.tag.charCodeAt(c) || 32);
        }
        view.setFloat32(off + 4, v.value, true);
      }
      count = wasm.get_glyph_outline_var(fontId, glyphId, varPtr, variations.length);
      wasm.dealloc(varPtr, varBytes);
    } else {
      count = wasm.get_glyph_outline(fontId, glyphId);
    }

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

  getFontFsType(fontId: number): number {
    const wasm = getShapingWasm();
    const ret = wasm.get_font_fs_type(fontId);
    return ret >= 0 ? ret : 0;
  }

  getFontAxes(fontId: number): readonly FontAxis[] {
    const wasm = getShapingWasm();
    const count = wasm.get_font_axes(fontId);
    if (count <= 0) return [];

    const ptr = wasm.get_font_axes_ptr();
    const view = new DataView(wasm.memory.buffer, ptr, count * 16);
    const axes: FontAxis[] = [];
    for (let i = 0; i < count; i++) {
      const off = i * 16;
      const tag = String.fromCharCode(
        view.getUint8(off),
        view.getUint8(off + 1),
        view.getUint8(off + 2),
        view.getUint8(off + 3),
      );
      const min = view.getFloat32(off + 4, true);
      const max = view.getFloat32(off + 8, true);
      const defaultValue = view.getFloat32(off + 12, true);
      axes.push({ tag, min, max, default: defaultValue });
    }
    return axes;
  }
}
