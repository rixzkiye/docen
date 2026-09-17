import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { createFontRef, initShapingWasm } from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.join(__dirname, "fixtures/fonts/OpenSans-Regular.ttf");

describe("R6.1 FontRef High-Level API", () => {
  let fontBytes: Uint8Array;

  beforeAll(async () => {
    await initShapingWasm();
    fontBytes = fs.readFileSync(openSansPath);
  });

  it("constructs FontRef and populates identity metadata", async () => {
    const font = await createFontRef(fontBytes);

    expect(font.familyName).toBe("Open Sans");
    expect(font.postscriptName).toBe("OpenSans-Regular");
    expect(font.unitsPerEm).toBe(2048);
    expect(font.identity.glyphCount).toBeGreaterThan(900);

    font.dispose();
  });

  it("shapes text and returns glyph runs and advances", async () => {
    const font = await createFontRef(fontBytes);
    const result = font.shape("OpenType Shaping");

    expect(result.glyphs.length).toBeGreaterThan(0);
    expect(result.totalAdvance).toBeGreaterThan(0);

    font.dispose();
  });

  it("extracts glyph vector outlines via FontRef", async () => {
    const font = await createFontRef(fontBytes);
    const outline = font.getGlyphOutline(43); // 'H'

    expect(outline.length).toBeGreaterThan(0);
    expect(outline[0].type).toBe("M");

    font.dispose();
  });

  it("prevents operations after dispose", async () => {
    const font = await createFontRef(fontBytes);
    font.dispose();

    expect(() => font.shape("Test")).toThrow(/disposed/);
    expect(() => font.getGlyphOutline(1)).toThrow(/disposed/);
  });
});
