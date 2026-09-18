import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import {
  RustybuzzBackend,
  createFontRefSync,
  generateToUnicodeCMap,
  initShapingWasm,
  isFontEmbeddingAllowed,
  isFontSubsettingAllowed,
} from "../src/index.js";
import { subsetFont } from "../src/subsetter.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.join(__dirname, "fixtures/fonts/OpenSans-Regular.ttf");
const arabicPath = path.join(__dirname, "fixtures/fonts/NotoNaskhArabic-Regular.ttf");

describe("R6.7 Font Embedding & Subsetting", () => {
  let openSansBytes: Uint8Array;
  let arabicBytes: Uint8Array;

  beforeAll(async () => {
    await initShapingWasm();
    openSansBytes = fs.readFileSync(openSansPath);
    arabicBytes = fs.readFileSync(arabicPath);
  });

  it("extracts fsType and evaluates embedding/subsetting permissions", () => {
    const fontRef = createFontRefSync(openSansBytes);
    expect(fontRef.identity.fsType).toBeDefined();
    expect(typeof fontRef.identity.fsType).toBe("number");
    // OpenSans has installable embedding (fsType = 0)
    expect(fontRef.identity.isEmbeddingAllowed).toBe(true);
    expect(fontRef.identity.isSubsettingAllowed).toBe(true);

    // Verify licensing helpers with known flags
    expect(isFontEmbeddingAllowed(0x0000)).toBe(true); // Installable
    expect(isFontEmbeddingAllowed(0x0002)).toBe(false); // Restricted
    expect(isFontEmbeddingAllowed(0x0004)).toBe(true); // Preview/Print
    expect(isFontEmbeddingAllowed(0x0008)).toBe(true); // Editable

    expect(isFontSubsettingAllowed(0x0000)).toBe(true);
    expect(isFontSubsettingAllowed(0x0100)).toBe(false); // No subsetting
    expect(isFontSubsettingAllowed(0x0002)).toBe(false);
  });

  it("subsets TrueType font, dramatically reducing file size while preserving valid sfnt structure", () => {
    const backend = new RustybuzzBackend();
    const origId = backend.registerFont(openSansBytes);

    // Shape a short text to get used glyph IDs
    const shapeRes = backend.shape(origId, "Hello World!");
    const usedGids = [0, ...shapeRes.glyphs.map((g) => g.glyphId)];

    const subsetBytes = subsetFont(openSansBytes, usedGids);

    expect(subsetBytes.length).toBeGreaterThan(0);
    // Original OpenSans is ~147 KB, subset should be drastically smaller (< 40 KB)
    expect(subsetBytes.length).toBeLessThan(openSansBytes.length / 3);

    // Verify TrueType magic header (0x00010000)
    const view = new DataView(subsetBytes.buffer, subsetBytes.byteOffset);
    expect(view.getUint32(0)).toBe(0x00010000);

    // Verify that the subset font is valid and can be registered in the WASM engine
    const subsetId = backend.registerFont(subsetBytes);
    expect(subsetId).toBeGreaterThan(0);

    // Verify shaping on the subset font works cleanly
    const subsetRes = backend.shape(subsetId, "Hello World!");
    expect(subsetRes.glyphs.length).toBe(shapeRes.glyphs.length);
  });

  it("subsets complex Arabic font with cursive joining and marks", () => {
    const backend = new RustybuzzBackend();
    const origId = backend.registerFont(arabicBytes);

    const arabicText = "مرحبا";
    const shapeRes = backend.shape(origId, arabicText, { direction: "rtl", script: "Arab" });
    const usedGids = [0, ...shapeRes.glyphs.map((g) => g.glyphId)];

    const subsetBytes = subsetFont(arabicBytes, usedGids);
    expect(subsetBytes.length).toBeLessThan(arabicBytes.length / 2);

    const subsetId = backend.registerFont(subsetBytes);
    expect(subsetId).toBeGreaterThan(0);
  });

  it("generates Adobe ToUnicode CMap stream for glyph ID to Unicode mapping", () => {
    const cidMap = new Map<number, number>([
      [1, 0x0041], // 'A'
      [2, 0x0042], // 'B'
      [3, 0x0645], // Arabic Meem
      [4, 0x20ac], // Euro sign €
      [5, 0x1f600], // Grinning face emoji (surrogate pair)
    ]);

    const cmap = generateToUnicodeCMap(cidMap, "Docen-ToUnicode");
    expect(cmap).toContain("/CMapName /Docen-ToUnicode def");
    expect(cmap).toContain("5 beginbfchar");
    expect(cmap).toContain("<0001> <0041>");
    expect(cmap).toContain("<0002> <0042>");
    expect(cmap).toContain("<0003> <0645>");
    expect(cmap).toContain("<0004> <20AC>");
    expect(cmap).toContain("<0005> <D83DDE00>"); // Surrogate pair for U+1F600
    expect(cmap).toContain("endbfchar");
    expect(cmap).toContain("endcmap");
  });
});
