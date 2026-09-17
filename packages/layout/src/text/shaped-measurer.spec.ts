import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { FontRef, createFontRefSync, initShapingWasm } from "@docen/shaping";
import { beforeAll, describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutTextStyle } from "../layout-doc";
import { ShapedMeasurer, isShapingEnabled, setShapingEnabled } from "./shaped-measurer";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.resolve(
  __dirname,
  "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
);

describe("R6.3 ShapedMeasurer (TextMeasurer with Opt-in Shaping)", () => {
  let openSansBytes: Uint8Array;
  let fontRef: FontRef;

  beforeAll(async () => {
    installFakeCanvas();
    await initShapingWasm();
    openSansBytes = fs.readFileSync(openSansPath);
    fontRef = createFontRefSync(openSansBytes);
  });

  it("behaves as standard TextMeasurer when shaping is disabled", () => {
    setShapingEnabled(false);
    expect(isShapingEnabled()).toBe(false);

    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: false });
    measurer.registerFont("Open Sans", fontRef);
    const style: LayoutTextStyle = {
      family: "Open Sans",
      sizePx: 16,
    };

    const width = measurer.widthOf("Hello World", style);
    expect(width).toBeGreaterThan(0);
    expect(measurer.shapeRun("Hello World", style)).toBeUndefined();
  });

  it("produces shaped glyph run and deterministic advances when opt-in enabled", () => {
    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Open Sans", fontRef);
    const style: LayoutTextStyle = {
      family: "Open Sans",
      sizePx: 16,
    };

    const run = measurer.shapeRun("AV", style);
    expect(run).toBeDefined();
    expect(run?.glyphs.length).toBe(2);
    expect(run?.totalAdvancePx).toBeGreaterThan(0);

    const width = measurer.widthOf("AV", style);
    expect(width).toBe(run?.totalAdvancePx);
  });

  it("correctly computes OpenType ligature substitution on fi pair vs unligatured sum", () => {
    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Open Sans", fontRef);
    const style: LayoutTextStyle = {
      family: "Open Sans",
      sizePx: 16,
    };

    const runFi = measurer.shapeRun("fi", style)!;
    const runF = measurer.shapeRun("f", style)!;
    const runI = measurer.shapeRun("i", style)!;

    // OpenType liga replaces "f" + "i" with a single "fi" ligature glyph
    expect(runFi.glyphs.length).toBe(1);
    expect(runF.glyphs.length).toBe(1);
    expect(runI.glyphs.length).toBe(1);
    expect(runFi.glyphs[0].glyphId).not.toBe(runF.glyphs[0].glyphId);
    expect(runFi.glyphs[0].glyphId).not.toBe(runI.glyphs[0].glyphId);
    expect(runFi.totalAdvancePx).toBeCloseTo(runF.totalAdvancePx + runI.totalAdvancePx, 2);
  });

  it("caches shaped glyph runs content-addressed", () => {
    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Open Sans", fontRef);
    const style: LayoutTextStyle = {
      family: "Open Sans",
      sizePx: 14,
    };

    const run1 = measurer.shapeRun("Deterministic Text", style);
    const run2 = measurer.shapeRun("Deterministic Text", style);

    expect(run1).toBe(run2); // exact same cached object reference
  });
});
