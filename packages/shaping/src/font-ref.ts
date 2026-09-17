import { getShapingBackend } from "./backend.js";
import { subsetFont } from "./subsetter.js";
import {
  isFontEmbeddingAllowed,
  isFontSubsettingAllowed,
  type FontAxis,
  type FontIdentity,
  type FontMetrics,
  type FontVariationSetting,
  type PathCommand,
  type ShapingBackend,
  type ShapingOptions,
  type ShapingResult,
  type VerticalFontMetrics,
} from "./types.js";
import { initShapingWasm, isShapingWasmInitialized } from "./wasm-loader.js";

/**
 * Handle representing an instantiated font face in the shaping engine.
 * Provides high-level methods for shaping text, extracting metrics, querying vector outlines,
 * and generating subsets.
 */
export class FontRef {
  readonly id: number;
  readonly backend: ShapingBackend;
  readonly metrics: FontMetrics;
  readonly identity: FontIdentity;
  private readonly fontData?: Uint8Array;
  private isDisposed = false;

  constructor(id: number, backend: ShapingBackend, fontData?: Uint8Array) {
    this.id = id;
    this.backend = backend;
    this.fontData = fontData;
    this.metrics = backend.getFontMetrics(id);

    const familyName = backend.getFontName?.(id, 1);
    const subfamilyName = backend.getFontName?.(id, 2);
    const fullName = backend.getFontName?.(id, 4);
    const postscriptName = backend.getFontName?.(id, 6);
    const glyphCount = backend.getGlyphCount?.(id) ?? 0;

    const fsType = backend.getFontFsType?.(id);
    const isEmbeddingAllowed = fsType !== undefined ? isFontEmbeddingAllowed(fsType) : true;
    const isSubsettingAllowed = fsType !== undefined ? isFontSubsettingAllowed(fsType) : true;
    const axes = backend.getFontAxes?.(id);

    this.identity = {
      familyName,
      subfamilyName,
      fullName,
      postscriptName,
      unitsPerEm: this.metrics.unitsPerEm,
      glyphCount,
      fsType,
      isEmbeddingAllowed,
      isSubsettingAllowed,
      axes,
    };
  }

  get familyName(): string | undefined {
    return this.identity.familyName;
  }

  get postscriptName(): string | undefined {
    return this.identity.postscriptName;
  }

  get unitsPerEm(): number {
    return this.metrics.unitsPerEm;
  }

  get verticalMetrics(): VerticalFontMetrics | undefined {
    return this.metrics.vertical;
  }

  getAxes(): readonly FontAxis[] {
    this.assertNotDisposed();
    return this.identity.axes ?? this.backend.getFontAxes?.(this.id) ?? [];
  }

  getMetrics(variations?: readonly FontVariationSetting[]): FontMetrics {
    this.assertNotDisposed();
    if (!variations || variations.length === 0) {
      return this.metrics;
    }
    return this.backend.getFontMetrics(this.id, variations);
  }

  /**
   * Shape a run of text with OpenType features and return glyphs and advances.
   */
  shape(text: string, options?: ShapingOptions): ShapingResult {
    this.assertNotDisposed();
    return this.backend.shape(this.id, text, options);
  }

  /**
   * Get vector outline path commands (M, L, Q, C, Z) for a given glyph.
   */
  getGlyphOutline(
    glyphId: number,
    variations?: readonly FontVariationSetting[],
  ): readonly PathCommand[] {
    this.assertNotDisposed();
    return this.backend.getGlyphOutline?.(this.id, glyphId, variations) ?? [];
  }

  /**
   * Produce a subset font buffer containing only the requested glyph IDs
   * (plus composite components and `.notdef`). Throws when the FontRef was
   * built without the source bytes; use {@link createFontRef} /
   * {@link createFontRefSync} so subsetting is always available.
   */
  subset(glyphIds: readonly number[]): Uint8Array {
    this.assertNotDisposed();
    if (!this.fontData) {
      throw new Error(
        `FontRef (id=${this.id}) has no source font bytes; construct it via createFontRef().`,
      );
    }
    return subsetFont(this.fontData, glyphIds);
  }

  /**
   * Free native memory associated with this font face.
   */
  dispose(): void {
    if (!this.isDisposed) {
      this.isDisposed = true;
      this.backend.dropFont(this.id);
    }
  }

  private assertNotDisposed(): void {
    if (this.isDisposed) {
      throw new Error(`FontRef (id=${this.id}) has been disposed.`);
    }
  }
}

/**
 * Asynchronously create and initialize a FontRef from font file bytes.
 */
export async function createFontRef(
  fontData: Uint8Array,
  options?: { backend?: string },
): Promise<FontRef> {
  const backendId = options?.backend ?? "rustybuzz";
  if (backendId === "rustybuzz" && !isShapingWasmInitialized()) {
    await initShapingWasm();
  }
  const backend = getShapingBackend(backendId);
  const fontId = backend.registerFont(fontData);
  return new FontRef(fontId, backend, fontData);
}

/**
 * Synchronously create a FontRef assuming WASM is already initialized.
 */
export function createFontRefSync(fontData: Uint8Array, options?: { backend?: string }): FontRef {
  const backend = getShapingBackend(options?.backend);
  const fontId = backend.registerFont(fontData);
  return new FontRef(fontId, backend, fontData);
}
