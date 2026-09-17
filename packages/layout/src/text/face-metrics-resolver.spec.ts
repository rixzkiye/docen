// A face registered with the shaping font manager — even directly, without
// going through registerShapingFont — must resolve its vertical metrics from
// its own tables and never touch the DOM probe. The oracle parses OS/2 +
// head straight from the sfnt bytes (no WASM); the environment has a DOM, so
// a fallback would surface as the 1.2 / 0.85 constants and an append call.

// @vitest-environment happy-dom
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createFontRefSync, initShapingWasm } from "@docen/shaping";
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest";

import { browserFontMetrics, clearFontMetricCache, clearRegisteredFontRatios } from "../font";
import { baselineShareOf } from "./measure";
import { getShapingFontManager } from "./shaped-measurer";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fontPath = path.resolve(
  __dirname,
  "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
);

/** OS/2 usWinAscent (74) / usWinDescent (76) / head unitsPerEm (18). */
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

describe("shaping-manager face metrics resolver", () => {
  beforeAll(async () => {
    // happy-dom defines `window`, so the loader would fetch the relative wasm
    // URL; hand it the bytes instead (the Node spec path).
    const wasmPath = path.resolve(__dirname, "../../../shaping/wasm/docen_shaping.wasm");
    await initShapingWasm(new Uint8Array(fs.readFileSync(wasmPath)));
  });

  afterEach(() => {
    clearFontMetricCache();
    vi.restoreAllMocks();
  });

  afterAll(() => {
    clearRegisteredFontRatios();
    clearFontMetricCache();
  });

  it("resolves ratio and baseline share from the registered file without DOM access", () => {
    const bytes = new Uint8Array(fs.readFileSync(fontPath));
    const oracle = winMetricsOf(bytes);
    const sum = oracle.ascent + oracle.descent;
    const family = "Open Sans Manager Oracle";
    const fontRef = createFontRefSync(bytes);
    // The direct manager path: no setRegisteredFontMetrics call anywhere.
    getShapingFontManager().registerActiveFont(family, fontRef);
    const append = vi.spyOn(document.body, "append");
    const create = vi.spyOn(document, "createElement");

    try {
      expect(browserFontMetrics.normalRatio({ family })).toBeCloseTo(
        (sum + 2 * Math.round(0.15 * sum)) / oracle.upem,
        5,
      );
      expect(baselineShareOf(family, false, false)).toBeCloseTo(oracle.ascent / oracle.upem, 5);
      // The probe path would append a span; nothing may be appended here.
      expect(append).not.toHaveBeenCalled();
      expect(create).not.toHaveBeenCalled();
    } finally {
      fontRef.dispose();
    }
  });

  it("still probes an untabulated family the shaping layer does not know", () => {
    const append = vi.spyOn(document.body, "append");
    expect(browserFontMetrics.normalRatio({ family: "Zzz Untabulated Face" })).toBe(1.2);
    expect(append).toHaveBeenCalled();
  });
});
