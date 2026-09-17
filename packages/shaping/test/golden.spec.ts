import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { RustybuzzBackend, initShapingWasm } from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.join(__dirname, "fixtures/fonts/OpenSans-Regular.ttf");
const arabicPath = path.join(__dirname, "fixtures/fonts/NotoNaskhArabic-Regular.ttf");
const goldensPath = path.join(__dirname, "fixtures/golden-records.json");

describe("R6.1 Golden hb-shape Parity Tests", () => {
  let rustybuzz: RustybuzzBackend;
  let openSansId: number;
  let arabicId: number;
  let goldens: any;

  beforeAll(async () => {
    await initShapingWasm();
    rustybuzz = new RustybuzzBackend();
    const openSansBytes = fs.readFileSync(openSansPath);
    const arabicBytes = fs.readFileSync(arabicPath);
    openSansId = rustybuzz.registerFont(openSansBytes);
    arabicId = rustybuzz.registerFont(arabicBytes);
    goldens = JSON.parse(fs.readFileSync(goldensPath, "utf-8"));
  });

  it("reproduces golden glyph sequence for Latin pangram", () => {
    const res = rustybuzz.shape(openSansId, "The quick brown fox jumps over the lazy dog.");
    const expected = goldens.latinAlphabet;

    expect(res.glyphs.length).toBe(expected.length);
    for (let i = 0; i < res.glyphs.length; i++) {
      expect(res.glyphs[i].glyphId).toBe(expected[i].glyphId);
      expect(res.glyphs[i].xAdvance).toBe(expected[i].xAdvance);
    }
  });

  it("reproduces golden glyph sequence for kerning pairs", () => {
    const res = rustybuzz.shape(openSansId, "AV To Wa LT");
    const expected = goldens.kerningPairs;

    expect(res.glyphs.length).toBe(expected.length);
    for (let i = 0; i < res.glyphs.length; i++) {
      expect(res.glyphs[i].glyphId).toBe(expected[i].glyphId);
      expect(res.glyphs[i].xAdvance).toBe(expected[i].xAdvance);
    }
  });

  it("reproduces golden glyph sequence for ligatures", () => {
    const res = rustybuzz.shape(openSansId, "office flight waffle");
    const expected = goldens.ligatures;

    expect(res.glyphs.length).toBe(expected.length);
    for (let i = 0; i < res.glyphs.length; i++) {
      expect(res.glyphs[i].glyphId).toBe(expected[i].glyphId);
      expect(res.glyphs[i].xAdvance).toBe(expected[i].xAdvance);
    }
  });

  it("reproduces golden glyph sequence for Arabic greeting", () => {
    const res = rustybuzz.shape(arabicId, "مرحبا", { direction: "rtl" });
    const expected = goldens.arabicGreeting;

    expect(res.glyphs.length).toBe(expected.length);
    for (let i = 0; i < res.glyphs.length; i++) {
      expect(res.glyphs[i].glyphId).toBe(expected[i].glyphId);
      expect(res.glyphs[i].xAdvance).toBe(expected[i].xAdvance);
      expect(res.glyphs[i].xOffset).toBe(expected[i].xOffset);
      expect(res.glyphs[i].yOffset).toBe(expected[i].yOffset);
    }
  });
});
