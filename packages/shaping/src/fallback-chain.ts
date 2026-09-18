import type { FontDescriptor } from "./font-identity.js";

export interface MissingGlyphReport {
  readonly char: string;
  readonly codePoint: number;
  readonly requestedFont: string;
  readonly script: string;
}

export class FallbackChainRegistry {
  private readonly scriptChains = new Map<string, FontDescriptor[]>();
  private readonly missingGlyphs: MissingGlyphReport[] = [];

  constructor() {
    this.registerDefaultChains();
  }

  private registerDefaultChains(): void {
    this.setChain("Latn", [
      { name: "Open Sans" },
      { name: "Noto Sans" },
      { name: "Arial" },
      { name: "DejaVu Sans" },
    ]);

    this.setChain("Arab", [
      { name: "Noto Naskh Arabic" },
      { name: "Noto Kufi Arabic" },
      { name: "Arial" },
    ]);

    this.setChain("Hani", [
      { name: "Noto Sans CJK SC" },
      { name: "Microsoft YaHei" },
      { name: "PingFang SC" },
      { name: "SimSun" },
    ]);

    this.setChain("Hebr", [{ name: "Noto Sans Hebrew" }, { name: "David" }, { name: "Arial" }]);
  }

  setChain(script: string, chain: FontDescriptor[]): void {
    this.scriptChains.set(script, chain);
  }

  getChain(script: string): readonly FontDescriptor[] {
    return this.scriptChains.get(script) ?? this.scriptChains.get("Latn") ?? [];
  }

  recordMissingGlyph(report: MissingGlyphReport): void {
    this.missingGlyphs.push(report);
  }

  getMissingGlyphs(): readonly MissingGlyphReport[] {
    return this.missingGlyphs;
  }

  clearMissingGlyphs(): void {
    this.missingGlyphs.length = 0;
  }

  /**
   * Determine primary script for a Unicode codepoint.
   */
  static detectScript(codePoint: number): string {
    if (codePoint >= 0x0600 && codePoint <= 0x06ff) return "Arab";
    if (codePoint >= 0x0590 && codePoint <= 0x05ff) return "Hebr";
    if (codePoint >= 0x4e00 && codePoint <= 0x9fff) return "Hani";
    if (codePoint >= 0x3400 && codePoint <= 0x4dbf) return "Hani";
    if (codePoint >= 0x0e00 && codePoint <= 0x0e7f) return "Thai";
    if (codePoint >= 0x0900 && codePoint <= 0x097f) return "Deva";
    return "Latn";
  }
}
