import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { HarfbuzzBackend } from "../src/harfbuzz-backend.js";
import { RustybuzzBackend, initShapingWasm } from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.join(__dirname, "fixtures/fonts/OpenSans-Regular.ttf");
const arabicPath = path.join(__dirname, "fixtures/fonts/NotoNaskhArabic-Regular.ttf");

describe("R6.0 Shaping Differential Oracle (Rustybuzz vs HarfBuzz)", () => {
  let rustybuzz: RustybuzzBackend;
  let harfbuzz: HarfbuzzBackend;
  let openSansBytes: Uint8Array;
  let arabicBytes: Uint8Array;
  let rbOpenSansId: number;
  let hbOpenSansId: number;
  let rbArabicId: number;
  let hbArabicId: number;

  beforeAll(async () => {
    await initShapingWasm();
    rustybuzz = new RustybuzzBackend();
    harfbuzz = new HarfbuzzBackend();

    openSansBytes = fs.readFileSync(openSansPath);
    arabicBytes = fs.readFileSync(arabicPath);

    rbOpenSansId = rustybuzz.registerFont(openSansBytes);
    hbOpenSansId = harfbuzz.registerFont(openSansBytes);

    rbArabicId = rustybuzz.registerFont(arabicBytes);
    hbArabicId = harfbuzz.registerFont(arabicBytes);
  });

  it("matches HarfBuzz for Latin kerning pairs", () => {
    const pairs = ["AV", "To", "Wa", "LT", "Te", "Vo", "Yo"];

    for (const pair of pairs) {
      const rbRes = rustybuzz.shape(rbOpenSansId, pair);
      const hbRes = harfbuzz.shape(hbOpenSansId, pair);

      expect(rbRes.glyphs.length).toBe(hbRes.glyphs.length);
      expect(rbRes.totalAdvance).toBeCloseTo(hbRes.totalAdvance, 1);

      for (let i = 0; i < rbRes.glyphs.length; i++) {
        expect(rbRes.glyphs[i].glyphId).toBe(hbRes.glyphs[i].glyphId);
        expect(rbRes.glyphs[i].xAdvance).toBe(hbRes.glyphs[i].xAdvance);
        expect(rbRes.glyphs[i].yAdvance).toBe(hbRes.glyphs[i].yAdvance);
        expect(rbRes.glyphs[i].xOffset).toBe(hbRes.glyphs[i].xOffset);
        expect(rbRes.glyphs[i].yOffset).toBe(hbRes.glyphs[i].yOffset);
      }
    }
  });

  it("matches HarfBuzz for standard OpenType ligatures (fi, fl, ffi, ffl)", () => {
    const testCases = ["office", "flight", "waffle", "difficult", "fi fl ffi ffl"];

    for (const text of testCases) {
      const rbRes = rustybuzz.shape(rbOpenSansId, text);
      const hbRes = harfbuzz.shape(hbOpenSansId, text);

      expect(rbRes.glyphs.length).toBe(hbRes.glyphs.length);
      expect(rbRes.totalAdvance).toBeCloseTo(hbRes.totalAdvance, 1);

      for (let i = 0; i < rbRes.glyphs.length; i++) {
        expect(rbRes.glyphs[i].glyphId).toBe(hbRes.glyphs[i].glyphId);
        expect(rbRes.glyphs[i].xAdvance).toBe(hbRes.glyphs[i].xAdvance);
        expect(rbRes.glyphs[i].yAdvance).toBe(hbRes.glyphs[i].yAdvance);
        expect(rbRes.glyphs[i].xOffset).toBe(hbRes.glyphs[i].xOffset);
        expect(rbRes.glyphs[i].yOffset).toBe(hbRes.glyphs[i].yOffset);
      }
    }
  });

  it("matches HarfBuzz for Arabic cursive joining and mark positioning", () => {
    const arabicWords = ["مرحبا", "العربية", "كتاب", "محمد"];

    for (const text of arabicWords) {
      const rbRes = rustybuzz.shape(rbArabicId, text, { direction: "rtl" });
      const hbRes = harfbuzz.shape(hbArabicId, text, { direction: "rtl" });

      expect(rbRes.glyphs.length).toBe(hbRes.glyphs.length);
      expect(rbRes.totalAdvance).toBeCloseTo(hbRes.totalAdvance, 1);

      for (let i = 0; i < rbRes.glyphs.length; i++) {
        expect(rbRes.glyphs[i].glyphId).toBe(hbRes.glyphs[i].glyphId);
        expect(rbRes.glyphs[i].xAdvance).toBe(hbRes.glyphs[i].xAdvance);
        expect(rbRes.glyphs[i].yAdvance).toBe(hbRes.glyphs[i].yAdvance);
        expect(rbRes.glyphs[i].xOffset).toBe(hbRes.glyphs[i].xOffset);
        expect(rbRes.glyphs[i].yOffset).toBe(hbRes.glyphs[i].yOffset);
      }
    }
  });

  it("extracts font metrics correctly via Skrifa", () => {
    const metrics = rustybuzz.getFontMetrics(rbOpenSansId);
    expect(metrics.unitsPerEm).toBe(2048);
    expect(metrics.ascender).toBeGreaterThan(0);
    expect(metrics.descender).toBeLessThan(0);
    expect(metrics.capHeight).toBeGreaterThan(0);
    expect(metrics.xHeight).toBeGreaterThan(0);
  });

  it("extracts glyph vector outline commands via Skrifa OutlinePen", () => {
    // OpenSans glyph 43 ('H')
    const commands = rustybuzz.getGlyphOutline(rbOpenSansId, 43);
    expect(commands.length).toBeGreaterThan(0);

    const first = commands[0];
    expect(first.type).toBe("M");

    const closes = commands.filter((c) => c.type === "Z");
    expect(closes.length).toBeGreaterThanOrEqual(1);
  });
});
