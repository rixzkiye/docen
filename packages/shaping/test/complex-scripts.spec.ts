import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { HarfbuzzBackend } from "../src/harfbuzz-backend.js";
import { RustybuzzBackend, initShapingWasm } from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fontsDir = path.join(__dirname, "fixtures/fonts");

describe("R6.5 Complex Scripts & Vertical CJK", () => {
  let rustybuzz: RustybuzzBackend;
  let harfbuzz: HarfbuzzBackend;

  const loadFont = (filename: string) => {
    const bytes = fs.readFileSync(path.join(fontsDir, filename));
    const rbId = rustybuzz.registerFont(bytes);
    const hbId = harfbuzz.registerFont(bytes);
    return { bytes, rbId, hbId };
  };

  beforeAll(async () => {
    await initShapingWasm();
    rustybuzz = new RustybuzzBackend();
    harfbuzz = new HarfbuzzBackend();
  });

  it("matches HarfBuzz for Hebrew RTL shaping and Niqqud mark positioning", () => {
    const { rbId, hbId } = loadFont("NotoSansHebrew-Regular.ttf");
    const testCases = ["שָׁלוֹם", "עֲלֵיכֶם", "בְּרֵאשִׁית", "יְרוּשָׁלַיִם"];

    for (const text of testCases) {
      const rbRes = rustybuzz.shape(rbId, text, { direction: "rtl", script: "Hebr" });
      const hbRes = harfbuzz.shape(hbId, text, { direction: "rtl", script: "Hebr" });

      expect(rbRes.glyphs.length).toBe(hbRes.glyphs.length);
      expect(rbRes.totalAdvance).toBeCloseTo(hbRes.totalAdvance, 1);

      for (let i = 0; i < rbRes.glyphs.length; i++) {
        expect(rbRes.glyphs[i].glyphId).toBe(hbRes.glyphs[i].glyphId);
        expect(rbRes.glyphs[i].xAdvance).toBe(hbRes.glyphs[i].xAdvance);
        expect(rbRes.glyphs[i].xOffset).toBe(hbRes.glyphs[i].xOffset);
        expect(rbRes.glyphs[i].yOffset).toBe(hbRes.glyphs[i].yOffset);
      }
    }
  });

  it("matches HarfBuzz for Thai tone mark stacking and vowel reordering", () => {
    const { rbId, hbId } = loadFont("NotoSansThai-Regular.ttf");
    const testCases = ["สวัสดีครับ", "น้ำ", "ผู้ใหญ่", "กินข้าว"];

    for (const text of testCases) {
      const rbRes = rustybuzz.shape(rbId, text, { script: "Thai" });
      const hbRes = harfbuzz.shape(hbId, text, { script: "Thai" });

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

  it("matches HarfBuzz for Indic (Devanagari) conjunct ligatures and matras", () => {
    const { rbId, hbId } = loadFont("NotoSansDevanagari-Regular.ttf");
    const testCases = ["नमस्ते", "हिन्दी", "शक्ति", "विद्या"];

    for (const text of testCases) {
      const rbRes = rustybuzz.shape(rbId, text, { script: "Deva" });
      const hbRes = harfbuzz.shape(hbId, text, { script: "Deva" });

      expect(rbRes.glyphs.length).toBe(hbRes.glyphs.length);
      expect(rbRes.totalAdvance).toBeCloseTo(hbRes.totalAdvance, 1);

      for (let i = 0; i < rbRes.glyphs.length; i++) {
        expect(rbRes.glyphs[i].glyphId).toBe(hbRes.glyphs[i].glyphId);
        expect(rbRes.glyphs[i].xAdvance).toBe(hbRes.glyphs[i].xAdvance);
      }
    }
  });

  it("matches HarfBuzz for Khmer sub-consonant and coeng shaping", () => {
    const { rbId, hbId } = loadFont("NotoSansKhmer-Regular.ttf");
    const testCases = ["ជំរាបសួរ", "ភាសាខ្មែរ", "កម្ពុជា"];

    for (const text of testCases) {
      const rbRes = rustybuzz.shape(rbId, text, { script: "Khmr" });
      const hbRes = harfbuzz.shape(hbId, text, { script: "Khmr" });

      expect(rbRes.glyphs.length).toBe(hbRes.glyphs.length);
      expect(rbRes.totalAdvance).toBeCloseTo(hbRes.totalAdvance, 1);

      for (let i = 0; i < rbRes.glyphs.length; i++) {
        expect(rbRes.glyphs[i].glyphId).toBe(hbRes.glyphs[i].glyphId);
      }
    }
  });

  it("extracts vertical metrics from native vhea table or synthesizes standard OpenType box", () => {
    const { rbId: vheaFontId } = loadFont("NotoSansAdlamUnjoined-Regular.ttf");
    const vheaMetrics = rustybuzz.getFontMetrics(vheaFontId);

    expect(vheaMetrics.vertical).toBeDefined();
    expect(vheaMetrics.vertical!.ascender).toBe(500);
    expect(vheaMetrics.vertical!.descender).toBe(-500);
    expect(vheaMetrics.vertical!.lineGap).toBe(1000);

    // Font without vhea table synthesizes standard vertical metrics
    const { rbId: noVheaId } = loadFont("NotoFangsongKSSVertical-Regular.ttf");
    const synthMetrics = rustybuzz.getFontMetrics(noVheaId);
    expect(synthMetrics.vertical).toBeDefined();
    expect(synthMetrics.vertical!.ascender).toBe(500);
    expect(synthMetrics.vertical!.descender).toBe(-500);
  });

  it("shapes vertical CJK text with top-to-bottom advance", () => {
    const { rbId } = loadFont("NotoFangsongKSSVertical-Regular.ttf");
    const text = "天地玄黃";

    const rbRes = rustybuzz.shape(rbId, text, { direction: "ttb" });

    expect(rbRes.glyphs.length).toBe(4);

    // In OpenType vertical layout, yAdvance is negative (top-to-bottom in font space)
    for (const g of rbRes.glyphs) {
      expect(g.yAdvance).toBeLessThan(0);
      expect(Math.abs(g.yAdvance)).toBe(2000);
    }
    // Total advance accumulates absolute vertical advance along the layout axis
    expect(rbRes.totalAdvance).toBe(8000);
  });

  it("auto-detects script and direction from text without explicit options", () => {
    const { rbId: hebrId } = loadFont("NotoSansHebrew-Regular.ttf");
    const { rbId: thaiId } = loadFont("NotoSansThai-Regular.ttf");

    // Hebrew text without direction option should auto-detect RTL
    const hebrewAuto = rustybuzz.shape(hebrId, "שָׁלוֹם");
    const hebrewExplicit = rustybuzz.shape(hebrId, "שָׁלוֹם", { direction: "rtl" });
    expect(hebrewAuto.glyphs).toEqual(hebrewExplicit.glyphs);

    // Thai text without script option should auto-detect Thai script
    const thaiAuto = rustybuzz.shape(thaiId, "สวัสดี");
    const thaiExplicit = rustybuzz.shape(thaiId, "สวัสดี", { script: "Thai" });
    expect(thaiAuto.glyphs).toEqual(thaiExplicit.glyphs);
  });
});
