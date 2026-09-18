import { FallbackChainRegistry, type MissingGlyphReport } from "./fallback-chain.js";
import { FontCache } from "./font-cache.js";
import { computeFontHash, type FontDescriptor } from "./font-identity.js";
import { FontRef, createFontRef, createFontRefSync } from "./font-ref.js";
import { BundledFontSource, MemoryFontSource, type FontSource } from "./font-source.js";

export interface FontManagerOptions {
  readonly maxMemoryEntries?: number;
  readonly enableOpfs?: boolean;
}

export class FontManager {
  readonly cache: FontCache;
  readonly fallbackRegistry = new FallbackChainRegistry();
  readonly memorySource = new MemoryFontSource();
  readonly bundledSource = new BundledFontSource();
  private readonly sources: FontSource[] = [];
  private readonly activeFontRefs = new Map<string, FontRef>();

  constructor(options?: FontManagerOptions) {
    this.cache = new FontCache(options);
    this.sources.push(this.memorySource, this.bundledSource);
  }

  addSource(source: FontSource): void {
    this.sources.unshift(source);
  }

  registerFontBytes(name: string, fontBytes: Uint8Array): void {
    this.memorySource.registerFont(name, fontBytes);
  }

  getActiveFont(nameOrFamily: string): FontRef | undefined {
    return this.activeFontRefs.get(nameOrFamily.toLowerCase());
  }

  registerActiveFont(nameOrFamily: string, fontRef: FontRef): void {
    this.activeFontRefs.set(nameOrFamily.toLowerCase(), fontRef);
  }

  registerActiveFontSync(nameOrFamily: string, fontBytes: Uint8Array): FontRef {
    const fontRef = createFontRefSync(fontBytes);
    this.registerActiveFont(nameOrFamily, fontRef);
    return fontRef;
  }

  async resolveFont(descriptor: FontDescriptor): Promise<FontRef> {
    const key = (descriptor.hash ?? descriptor.name).toLowerCase();
    if (this.activeFontRefs.has(key)) {
      return this.activeFontRefs.get(key)!;
    }

    // 1. Try cache by hash or name
    let bytes = await this.cache.get(key);

    // 2. If not cached, query sources
    if (!bytes) {
      for (const source of this.sources) {
        bytes = await source.loadFont(descriptor);
        if (bytes) {
          const hash = descriptor.hash ?? (await computeFontHash(bytes));
          await this.cache.set(hash, bytes);
          await this.cache.set(descriptor.name.toLowerCase(), bytes);
          break;
        }
      }
    }

    if (!bytes) {
      throw new Error(`Font "${descriptor.name}" could not be resolved from any source.`);
    }

    const fontRef = await createFontRef(bytes);
    this.activeFontRefs.set(key, fontRef);
    return fontRef;
  }

  async resolveFallback(char: string, script?: string): Promise<FontRef | undefined> {
    const codePoint = char.codePointAt(0) ?? 0;
    const detectedScript = script ?? FallbackChainRegistry.detectScript(codePoint);
    const chain = this.fallbackRegistry.getChain(detectedScript);

    for (const descriptor of chain) {
      try {
        const fontRef = await this.resolveFont(descriptor);
        // Test if font shapes the character (not glyph 0)
        const shaped = fontRef.shape(char);
        if (shaped.glyphs.length > 0 && shaped.glyphs[0].glyphId !== 0) {
          return fontRef;
        }
      } catch {
        // Continue chain
      }
    }

    this.fallbackRegistry.recordMissingGlyph({
      char,
      codePoint,
      requestedFont: chain[0]?.name ?? "unknown",
      script: detectedScript,
    });

    return undefined;
  }

  async invalidate(key: string): Promise<void> {
    const normKey = key.toLowerCase();
    const existing = this.activeFontRefs.get(normKey);
    if (existing) {
      existing.dispose();
      this.activeFontRefs.delete(normKey);
    }
    await this.cache.delete(normKey);
  }

  getMissingGlyphs(): readonly MissingGlyphReport[] {
    return this.fallbackRegistry.getMissingGlyphs();
  }

  clearMissingGlyphs(): void {
    this.fallbackRegistry.clearMissingGlyphs();
  }

  dispose(): void {
    for (const fontRef of this.activeFontRefs.values()) {
      fontRef.dispose();
    }
    this.activeFontRefs.clear();
  }
}
