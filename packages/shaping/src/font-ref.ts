import { getShapingBackend } from "./backend.js";
import {
  isFontEmbeddingAllowed,
  isFontSubsettingAllowed,
  type FontIdentity,
  type FontMetrics,
  type PathCommand,
  type ShapingBackend,
  type ShapingOptions,
  type ShapingResult,
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
  private isDisposed = false;

  constructor(id: number, backend: ShapingBackend) {
    this.id = id;
    this.backend = backend;
    this.metrics = backend.getFontMetrics(id);

    const familyName = backend.getFontName?.(id, 1);
    const subfamilyName = backend.getFontName?.(id, 2);
    const fullName = backend.getFontName?.(id, 4);
    const postscriptName = backend.getFontName?.(id, 6);
    const glyphCount = backend.getGlyphCount?.(id) ?? 0;

    const fsType = backend.getFontFsType?.(id);
    const isEmbeddingAllowed = fsType !== undefined ? isFontEmbeddingAllowed(fsType) : true;
    const isSubsettingAllowed = fsType !== undefined ? isFontSubsettingAllowed(fsType) : true;

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
  getGlyphOutline(glyphId: number): readonly PathCommand[] {
    this.assertNotDisposed();
    return this.backend.getGlyphOutline?.(this.id, glyphId) ?? [];
  }

  /**
   * Produce a subset font buffer containing only the requested glyph IDs.
   */
  subset(glyphIds: readonly number[]): Uint8Array {
    this.assertNotDisposed();
    return this.backend.subset?.(this.id, glyphIds) ?? new Uint8Array(0);
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
  return new FontRef(fontId, backend);
}

/**
 * Synchronously create a FontRef assuming WASM is already initialized.
 */
export function createFontRefSync(fontData: Uint8Array, options?: { backend?: string }): FontRef {
  const backend = getShapingBackend(options?.backend);
  const fontId = backend.registerFont(fontData);
  return new FontRef(fontId, backend);
}
