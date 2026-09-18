import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { FontRef } from "../src/font-ref.js";
import { HarfbuzzBackend } from "../src/harfbuzz-backend.js";
import { RustybuzzBackend, initShapingWasm } from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fontsDir = path.join(__dirname, "fixtures/fonts");

describe("R6.6 OpenType Features & Variations", () => {
  let rustybuzz: RustybuzzBackend;
  let harfbuzz: HarfbuzzBackend;
  let interBytes: Uint8Array;
  let rbFontId: number;
  let hbFontId: number;
  let fontRef: FontRef;

  beforeAll(async () => {
    await initShapingWasm();
    rustybuzz = new RustybuzzBackend();
    harfbuzz = new HarfbuzzBackend();

    interBytes = fs.readFileSync(path.join(fontsDir, "InterVariable.ttf"));
    rbFontId = rustybuzz.registerFont(interBytes);
    hbFontId = harfbuzz.registerFont(interBytes);
    fontRef = new FontRef(rbFontId, rustybuzz);
  });

  it("discovers variation axes on variable fonts", () => {
    const axes = fontRef.getAxes();
    expect(axes.length).toBe(2);

    const opsz = axes.find((a) => a.tag === "opsz");
    expect(opsz).toBeDefined();
    expect(opsz?.min).toBe(14);
    expect(opsz?.max).toBe(32);
    expect(opsz?.default).toBe(14);

    const wght = axes.find((a) => a.tag === "wght");
    expect(wght).toBeDefined();
    expect(wght?.min).toBe(100);
    expect(wght?.max).toBe(900);
    expect(wght?.default).toBe(400);

    expect(fontRef.identity.axes).toEqual(axes);
  });

  it("shapes text with variation coordinates with 100% HarfBuzz parity", () => {
    const text = "123 Quick Fox";
    const weights = [100, 300, 400, 700, 900];

    for (const w of weights) {
      const rbRes = rustybuzz.shape(rbFontId, text, {
        variations: [{ tag: "wght", value: w }],
      });
      const hbRes = harfbuzz.shape(hbFontId, text, {
        variations: [{ tag: "wght", value: w }],
      });

      expect(rbRes.glyphs.length).toBe(hbRes.glyphs.length);
      expect(Math.abs(rbRes.totalAdvance - hbRes.totalAdvance)).toBeLessThanOrEqual(2);

      for (let i = 0; i < rbRes.glyphs.length; i++) {
        expect(rbRes.glyphs[i].glyphId).toBe(hbRes.glyphs[i].glyphId);
        expect(Math.abs(rbRes.glyphs[i].xAdvance - hbRes.glyphs[i].xAdvance)).toBeLessThanOrEqual(
          1,
        );
        expect(rbRes.glyphs[i].yAdvance).toBe(hbRes.glyphs[i].yAdvance);
        expect(rbRes.glyphs[i].xOffset).toBe(hbRes.glyphs[i].xOffset);
        expect(rbRes.glyphs[i].yOffset).toBe(hbRes.glyphs[i].yOffset);
      }
    }

    // Verify advances change with weight
    const lightRes = rustybuzz.shape(rbFontId, "123", {
      variations: [{ tag: "wght", value: 100 }],
    });
    const boldRes = rustybuzz.shape(rbFontId, "123", {
      variations: [{ tag: "wght", value: 900 }],
    });
    expect(boldRes.totalAdvance).toBeGreaterThan(lightRes.totalAdvance);
  });

  it("handles OpenType tabular numbers (tnum) feature with HarfBuzz parity", () => {
    const text = "1234567890";
    const rbDefault = rustybuzz.shape(rbFontId, text);
    const rbTnum = rustybuzz.shape(rbFontId, text, {
      features: [{ tag: "tnum", value: 1 }],
    });
    const hbTnum = harfbuzz.shape(hbFontId, text, {
      features: [{ tag: "tnum", value: 1 }],
    });

    // Glyph IDs change when tnum is enabled
    expect(rbTnum.glyphs[0].glyphId).not.toBe(rbDefault.glyphs[0].glyphId);
    expect(rbTnum.glyphs[0].glyphId).toBe(1360);

    // Parity vs HarfBuzz
    expect(rbTnum.glyphs.length).toBe(hbTnum.glyphs.length);
    for (let i = 0; i < rbTnum.glyphs.length; i++) {
      expect(rbTnum.glyphs[i].glyphId).toBe(hbTnum.glyphs[i].glyphId);
      expect(rbTnum.glyphs[i].xAdvance).toBe(hbTnum.glyphs[i].xAdvance);
    }
  });

  it("handles contextual alternates / ligatures (calt) with HarfBuzz parity", () => {
    const arrow = "->";
    const rbLigate = rustybuzz.shape(rbFontId, arrow, {
      features: [{ tag: "calt", value: 1 }],
    });
    const rbNoLigate = rustybuzz.shape(rbFontId, arrow, {
      features: [{ tag: "calt", value: 0 }],
    });

    const hbLigate = harfbuzz.shape(hbFontId, arrow, {
      features: [{ tag: "calt", value: 1 }],
    });
    const hbNoLigate = harfbuzz.shape(hbFontId, arrow, {
      features: [{ tag: "calt", value: 0 }],
    });

    // In Inter, "->" with calt forms a single ligature glyph [1805]
    expect(rbLigate.glyphs.length).toBe(1);
    expect(rbLigate.glyphs[0].glyphId).toBe(1805);
    expect(hbLigate.glyphs.length).toBe(1);
    expect(hbLigate.glyphs[0].glyphId).toBe(1805);

    // With calt: 0, it stays as 2 separate glyphs
    expect(rbNoLigate.glyphs.length).toBe(2);
    expect(hbNoLigate.glyphs.length).toBe(2);
    expect(rbNoLigate.glyphs[0].glyphId).toBe(hbNoLigate.glyphs[0].glyphId);
    expect(rbNoLigate.glyphs[1].glyphId).toBe(hbNoLigate.glyphs[1].glyphId);
  });

  it("extracts variation metrics via getMetrics(variations)", () => {
    const defaultMetrics = fontRef.metrics;
    const varMetrics = fontRef.getMetrics([{ tag: "wght", value: 900 }]);

    expect(defaultMetrics.unitsPerEm).toBe(2048);
    expect(varMetrics.unitsPerEm).toBe(2048);
    expect(varMetrics.ascender).toBeGreaterThan(0);
    expect(varMetrics.descender).toBeLessThan(0);
  });

  it("extracts variable font vector glyph outlines", () => {
    const glyphId = 1360; // digit 1
    const commandsDefault = fontRef.getGlyphOutline(glyphId);
    const commandsBold = fontRef.getGlyphOutline(glyphId, [{ tag: "wght", value: 900 }]);

    expect(commandsDefault.length).toBeGreaterThan(0);
    expect(commandsBold.length).toBeGreaterThan(0);
  });
});
