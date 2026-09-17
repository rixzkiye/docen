import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { FontRef, createFontRefSync, initShapingWasm } from "@docen/shaping";
import { beforeAll, describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutTextStyle } from "../layout-doc";
import {
  createMeasurer,
  isShapingEnabled,
  setShapingEnabled,
  ShapedMeasurer,
} from "./shaped-measurer";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.resolve(
  __dirname,
  "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
);

describe("R6.3 & R6.8 ShapedMeasurer (Default Flip & Rollback)", () => {
  let openSansBytes: Uint8Array;
  let fontRef: FontRef;

  beforeAll(async () => {
    installFakeCanvas();
    await initShapingWasm();
    openSansBytes = fs.readFileSync(openSansPath);
    fontRef = createFontRefSync(openSansBytes);
    setShapingEnabled(true);
  });

  it("enables deterministic shaping by default (R6.8 default flip)", () => {
    setShapingEnabled(true);
    expect(isShapingEnabled()).toBe(true);

    const measurer = createMeasurer(fakeFontMetrics);
    expect(measurer).toBeInstanceOf(ShapedMeasurer);
    expect((measurer as ShapedMeasurer).enabled).toBe(true);
  });

  it("supports rollback to standard TextMeasurer via setShapingEnabled(false)", () => {
    try {
      setShapingEnabled(false);
      expect(isShapingEnabled()).toBe(false);

      const measurer = createMeasurer(fakeFontMetrics);
      expect(measurer).not.toBeInstanceOf(ShapedMeasurer);
    } finally {
      setShapingEnabled(true);
    }
  });

  it("behaves as standard TextMeasurer when shaping is disabled", () => {
    try {
      setShapingEnabled(false);
      expect(isShapingEnabled()).toBe(false);

      const measurer = new ShapedMeasurer(fakeFontMetrics, { enabled: false });
      measurer.registerFont("Open Sans", fontRef);
      const style: LayoutTextStyle = {
        family: "Open Sans",
        sizePx: 16,
      };

      const width = measurer.widthOf("Hello World", style);
      expect(width).toBeGreaterThan(0);
      expect(measurer.shapeRun("Hello World", style)).toBeUndefined();
    } finally {
      setShapingEnabled(true);
    }
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

  it("shapes RTL scripts with correct cluster and advance layout", () => {
    const arabicPath = path.resolve(
      __dirname,
      "../../../shaping/test/fixtures/fonts/NotoNaskhArabic-Regular.ttf",
    );
    const arabicBytes = fs.readFileSync(arabicPath);
    const arabicFont = createFontRefSync(arabicBytes);

    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Noto Naskh Arabic", arabicFont);

    const style: LayoutTextStyle = {
      family: "Noto Naskh Arabic",
      sizePx: 16,
      direction: "rtl",
    };

    const run = measurer.shapeRun("العربية", style);
    expect(run).toBeDefined();
    expect(run?.direction).toBe("rtl");
    expect(run?.totalAdvancePx).toBeGreaterThan(0);
    expect(run?.glyphs.length).toBeGreaterThan(0);
  });

  it("shapes vertical text runs with top-to-bottom advances", () => {
    const vertPath = path.resolve(
      __dirname,
      "../../../shaping/test/fixtures/fonts/NotoFangsongKSSVertical-Regular.ttf",
    );
    const vertBytes = fs.readFileSync(vertPath);
    const vertFont = createFontRefSync(vertBytes);

    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Noto Fangsong", vertFont);

    const style: LayoutTextStyle = {
      family: "Noto Fangsong",
      sizePx: 20,
      vertical: true,
    };

    const run = measurer.shapeRun("天地玄黃", style);
    expect(run).toBeDefined();
    expect(run?.direction).toBe("ttb");
    expect(run?.totalAdvancePx).toBeGreaterThan(0);
    // In vertical layout, yPx advances down the layout axis
    expect(run?.glyphs.length).toBe(4);
    for (let i = 1; i < run!.glyphs.length; i++) {
      expect(run!.glyphs[i]!.yPx).toBeGreaterThan(run!.glyphs[i - 1]!.yPx);
    }
  });

  it("applies OpenType number spacing and ligature features from style", () => {
    const interPath = path.resolve(
      __dirname,
      "../../../shaping/test/fixtures/fonts/InterVariable.ttf",
    );
    const interBytes = fs.readFileSync(interPath);
    const interFont = createFontRefSync(interBytes);

    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Inter", interFont);

    // Test tabular number spacing
    const propRun = measurer.shapeRun("123", {
      family: "Inter",
      sizePx: 16,
      numSpacing: "proportional",
    })!;
    const tabRun = measurer.shapeRun("123", {
      family: "Inter",
      sizePx: 16,
      numSpacing: "tabular",
    })!;

    expect(propRun.glyphs[0].glyphId).not.toBe(tabRun.glyphs[0].glyphId);
    expect(tabRun.glyphs[0].glyphId).toBe(1360);

    // Test ligature disabling (w:ligatures none)
    const ligRun = measurer.shapeRun("->", {
      family: "Inter",
      sizePx: 16,
    })!;
    const noLigRun = measurer.shapeRun("->", {
      family: "Inter",
      sizePx: 16,
      ligatures: "none",
    })!;

    expect(ligRun.glyphs.length).toBe(1);
    expect(ligRun.glyphs[0].glyphId).toBe(1805);
    expect(noLigRun.glyphs.length).toBe(2);

    // Test variable font weight axis
    const lightRun = measurer.shapeRun("123", {
      family: "Inter",
      sizePx: 16,
      fontWeight: 100,
    })!;
    const boldRun = measurer.shapeRun("123", {
      family: "Inter",
      sizePx: 16,
      fontWeight: 900,
    })!;

    expect(boldRun.totalAdvancePx).toBeGreaterThan(lightRun.totalAdvancePx);
  });
});
