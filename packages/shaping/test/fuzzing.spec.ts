import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { RustybuzzBackend, initShapingWasm } from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.join(__dirname, "fixtures/fonts/OpenSans-Regular.ttf");

describe("R6.1 WASM Boundary & Fuzzing Safety", () => {
  let rustybuzz: RustybuzzBackend;
  let validFontBytes: Uint8Array;

  beforeAll(async () => {
    await initShapingWasm();
    rustybuzz = new RustybuzzBackend();
    validFontBytes = fs.readFileSync(openSansPath);
  });

  it("safely rejects 0-byte font buffer without panic", () => {
    expect(() => rustybuzz.registerFont(new Uint8Array(0))).toThrow();
  });

  it("safely rejects truncated and random junk bytes without crashing WASM", () => {
    const corruptInputs = [
      new Uint8Array([0x00, 0x01, 0x00]),
      new Uint8Array(32).fill(0xff),
      new Uint8Array(1024).map((_, i) => (i * 37) % 256),
    ];

    for (const corrupt of corruptInputs) {
      expect(() => rustybuzz.registerFont(corrupt)).toThrow();
    }
  });

  it("handles empty and single-character strings gracefully", () => {
    const fontId = rustybuzz.registerFont(validFontBytes);
    const emptyRes = rustybuzz.shape(fontId, "");
    expect(emptyRes.glyphs.length).toBe(0);
    expect(emptyRes.totalAdvance).toBe(0);

    const singleRes = rustybuzz.shape(fontId, "A");
    expect(singleRes.glyphs.length).toBe(1);
    expect(singleRes.totalAdvance).toBeGreaterThan(0);

    rustybuzz.dropFont(fontId);
  });

  it("handles large text input (30,000 characters) without stack or memory overflow", () => {
    const fontId = rustybuzz.registerFont(validFontBytes);
    const largeText = "The quick brown fox jumps over the lazy dog. ".repeat(650); // ~30k chars

    const res = rustybuzz.shape(fontId, largeText);
    expect(res.glyphs.length).toBeGreaterThan(25000);
    expect(res.totalAdvance).toBeGreaterThan(0);

    rustybuzz.dropFont(fontId);
  });

  it("handles out-of-range glyph IDs for outline queries without crashing", () => {
    const fontId = rustybuzz.registerFont(validFontBytes);

    expect(rustybuzz.getGlyphOutline(fontId, 999999)).toEqual([]);
    expect(rustybuzz.getGlyphOutline(fontId, 0)).toBeDefined();

    rustybuzz.dropFont(fontId);
  });

  it("safely rejects operations on invalid or already dropped font IDs", () => {
    const fontId = rustybuzz.registerFont(validFontBytes);
    rustybuzz.dropFont(fontId);

    expect(() => rustybuzz.shape(fontId, "Hello")).toThrow();
    expect(() => rustybuzz.getFontMetrics(fontId)).toThrow();
  });

  it("maintains heap stability across 100 rapid alloc/register/shape/drop cycles", () => {
    for (let i = 0; i < 100; i++) {
      const id = rustybuzz.registerFont(validFontBytes);
      const res = rustybuzz.shape(id, `Cycle test ${i}`);
      expect(res.glyphs.length).toBeGreaterThan(5);
      rustybuzz.dropFont(id);
    }
  });
});
