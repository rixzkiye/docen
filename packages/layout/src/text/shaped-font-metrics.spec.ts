// The registered-font Word ratio must come from the face's own OS/2 metrics.
// The oracle below parses usWinAscent/usWinDescent/unitsPerEm straight from
// the sfnt bytes — no WASM involved — so the integration (metrics buffer →
// FontRef.metrics → the published ratio) is checked against the font file
// itself, not against a second copy of the implementation.

// @vitest-environment node
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { initShapingWasm } from "@docen/shaping";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { browserFontMetrics, clearRegisteredFontRatios } from "../font";
import { registerShapingFont } from "./shaped-measurer";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fontPath = path.resolve(
  __dirname,
  "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
);

/** OS/2 usWinAscent (offset 74) / usWinDescent (76) / head unitsPerEm (18). */
function winMetricsOf(bytes: Uint8Array): { ascent: number; descent: number; upem: number } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const numTables = view.getUint16(4);
  let os2 = -1;
  let head = -1;
  for (let i = 0; i < numTables; i++) {
    const rec = 12 + i * 16;
    const tag = String.fromCharCode(bytes[rec]!, bytes[rec + 1]!, bytes[rec + 2]!, bytes[rec + 3]!);
    if (tag === "OS/2") os2 = view.getUint32(rec + 8);
    if (tag === "head") head = view.getUint32(rec + 8);
  }
  if (os2 < 0 || head < 0) throw new Error("font is missing OS/2 or head");
  return {
    ascent: view.getUint16(os2 + 74),
    descent: view.getUint16(os2 + 76),
    upem: view.getUint16(head + 18),
  };
}

describe("registered-font Word ratio", () => {
  beforeAll(async () => {
    await initShapingWasm();
  });

  afterAll(() => {
    clearRegisteredFontRatios();
  });

  it("publishes Word's formula from the face's OS/2 metrics", () => {
    const bytes = new Uint8Array(fs.readFileSync(fontPath));
    const oracle = winMetricsOf(bytes);
    const sum = oracle.ascent + oracle.descent;
    const expected = (sum + 2 * Math.round(0.15 * sum)) / oracle.upem;

    const family = "Open Sans Ratio Oracle";
    const fontRef = registerShapingFont(family, bytes);

    expect(fontRef.metrics.winAscent).toBe(oracle.ascent);
    expect(fontRef.metrics.winDescent).toBe(oracle.descent);
    expect(browserFontMetrics.normalRatio({ family })).toBeCloseTo(expected, 5);
  });
});
