import type { FontDescriptor } from "./font-identity.js";

export interface FontSource {
  readonly id: string;
  loadFont(descriptor: FontDescriptor): Promise<Uint8Array | undefined>;
}

export class MemoryFontSource implements FontSource {
  readonly id = "memory";
  private readonly fonts = new Map<string, Uint8Array>();

  registerFont(name: string, data: Uint8Array): void {
    this.fonts.set(name.toLowerCase(), data);
  }

  unregisterFont(name: string): boolean {
    return this.fonts.delete(name.toLowerCase());
  }

  async loadFont(descriptor: FontDescriptor): Promise<Uint8Array | undefined> {
    const targetName = descriptor.name.toLowerCase();
    if (this.fonts.has(targetName)) {
      return this.fonts.get(targetName);
    }
    if (descriptor.postscriptName) {
      const psName = descriptor.postscriptName.toLowerCase();
      if (this.fonts.has(psName)) {
        return this.fonts.get(psName);
      }
    }
    return undefined;
  }
}

export class BundledFontSource implements FontSource {
  readonly id = "bundled";
  private readonly bundledFonts = new Map<string, Uint8Array>();

  addBundledFont(name: string, data: Uint8Array): void {
    this.bundledFonts.set(name.toLowerCase(), data);
  }

  async loadFont(descriptor: FontDescriptor): Promise<Uint8Array | undefined> {
    const key = descriptor.name.toLowerCase();
    return this.bundledFonts.get(key);
  }
}

export class LocalFontSource implements FontSource {
  readonly id = "local";

  async loadFont(descriptor: FontDescriptor): Promise<Uint8Array | undefined> {
    if (
      typeof window !== "undefined" &&
      "queryLocalFonts" in window &&
      typeof (window as any).queryLocalFonts === "function"
    ) {
      try {
        const localFonts = await (window as any).queryLocalFonts({
          postscriptNames: descriptor.postscriptName ? [descriptor.postscriptName] : undefined,
        });
        for (const fontData of localFonts) {
          if (
            fontData.family?.toLowerCase() === descriptor.name.toLowerCase() ||
            fontData.postscriptName?.toLowerCase() === descriptor.postscriptName?.toLowerCase()
          ) {
            const blob = await fontData.blob();
            const buffer = await blob.arrayBuffer();
            return new Uint8Array(buffer);
          }
        }
      } catch {
        return undefined;
      }
    }
    return undefined;
  }
}
