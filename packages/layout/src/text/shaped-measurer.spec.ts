import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { FontManager, FontRef, createFontRefSync, initShapingWasm } from "@docen/shaping";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutInline, LayoutTextStyle } from "../layout-doc";
import { itemGlyphLayoutOf } from "./glyphs";
import { packLines } from "./line-break";
import {
  createMeasurer,
  isShapingEnabled,
  registerShapingFont,
  setShapingEnabled,
  ShapedMeasurer,
} from "./shaped-measurer";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.resolve(
  __dirname,
  "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
);

describe("R6.3 & R6.8 ShapedMeasurer (Opt-in Shaping & Canvas Default)", () => {
  let openSansBytes: Uint8Array;
  let fontRef: FontRef;

  beforeAll(async () => {
    installFakeCanvas();
    await initShapingWasm();
    openSansBytes = fs.readFileSync(openSansPath);
    fontRef = createFontRefSync(openSansBytes);
    setShapingEnabled(false);
  });

  afterAll(() => {
    setShapingEnabled(false);
  });

  it("defaults to canvas until shaping is explicitly opted in (R6.8 flip reverted)", () => {
    setShapingEnabled(false);
    expect(isShapingEnabled()).toBe(false);

    const measurer = createMeasurer(fakeFontMetrics);
    expect(measurer).not.toBeInstanceOf(ShapedMeasurer);
    expect(measurer.glyphRunOf("Hello", { family: "Open Sans", sizePx: 16 })).toBeUndefined();

    try {
      setShapingEnabled(true);
      expect(isShapingEnabled()).toBe(true);
      expect(createMeasurer(fakeFontMetrics)).toBeInstanceOf(ShapedMeasurer);
    } finally {
      setShapingEnabled(false);
    }
  });

  it("opts in through the fontManager option without the global flag", () => {
    setShapingEnabled(false);
    const measurer = createMeasurer(fakeFontMetrics, {
      fontManager: new FontManager({ enableOpfs: false }),
    });
    expect(measurer).toBeInstanceOf(ShapedMeasurer);
  });

  it("supports rollback to standard TextMeasurer via setShapingEnabled(false)", () => {
    try {
      setShapingEnabled(false);
      expect(isShapingEnabled()).toBe(false);

      const measurer = createMeasurer(fakeFontMetrics);
      expect(measurer).not.toBeInstanceOf(ShapedMeasurer);
    } finally {
      setShapingEnabled(false);
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
      setShapingEnabled(false);
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

  it("produces glyph runs through the shared font manager and carries unitsPerEm", () => {
    registerShapingFont("Open Sans", openSansBytes);
    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    const style: LayoutTextStyle = { family: "Open Sans", sizePx: 16 };

    const run = measurer.glyphRunOf("fi", style)!;
    expect(run).toBeDefined();
    expect(run.unitsPerEm).toBe(2048); // Open Sans — the upem=1000 assumption was the bug
    expect(run.glyphs).toHaveLength(1); // Open Sans liga
    expect(run.glyphs[0]!.glyphId).toBeGreaterThan(0);
    expect(measurer.widthOf("fi", style)).toBeCloseTo(run.totalAdvancePx, 5);
  });

  it("lays out RTL glyphs in visual order (first glyph at the right edge)", () => {
    const arabicBytes = fs.readFileSync(
      path.resolve(__dirname, "../../../shaping/test/fixtures/fonts/NotoNaskhArabic-Regular.ttf"),
    );
    const arabicFont = createFontRefSync(arabicBytes);
    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Noto Naskh Arabic", arabicFont);
    const run = measurer.shapeRun("با", {
      family: "Noto Naskh Arabic",
      sizePx: 16,
      direction: "rtl",
    })!;
    expect(run.glyphs.length).toBeGreaterThanOrEqual(2);
    const first = run.glyphs[0]!;
    const last = run.glyphs[run.glyphs.length - 1]!;
    expect(first.xPx).toBeGreaterThan(last.xPx);
  });

  it("keeps canvas fallback for letter-spaced runs the painter cannot place", () => {
    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Open Sans", fontRef);
    const style: LayoutTextStyle = { family: "Open Sans", sizePx: 16, letterSpacingPx: 2 };
    expect(measurer.glyphRunOf("spaced", style)).toBeUndefined();
    expect(measurer.widthOf("spaced", style)).toBeGreaterThan(0);
  });

  it("picks the script slot for single-script runs and skips mixed runs", () => {
    const cjkBytes = fs.readFileSync(
      path.resolve(
        __dirname,
        "../../../shaping/test/fixtures/fonts/NotoFangsongKSSVertical-Regular.ttf",
      ),
    );
    const cjkFont = createFontRefSync(cjkBytes);
    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Open Sans", fontRef);
    measurer.registerFont("CJK Face", cjkFont);
    const slots: LayoutTextStyle = {
      family: { latin: "Open Sans", eastAsia: "CJK Face" },
      sizePx: 16,
    };

    expect(measurer.glyphRunOf("天地玄黃", slots)?.fontName).toBe("CJK Face");
    expect(measurer.glyphRunOf("Hello", slots)?.fontName).toBe("Open Sans");
    expect(measurer.glyphRunOf("中文 test", slots)).toBeUndefined();
  });

  it("attaches produced glyph runs to laid-out items and maps caret clusters", () => {
    installFakeCanvas();
    setShapingEnabled(false);
    const measurer = new ShapedMeasurer(fakeFontMetrics, { optIn: true });
    measurer.registerFont("Open Sans", fontRef);
    const style: LayoutTextStyle = { family: "Open Sans", sizePx: 16 };
    const inline: LayoutInline[] = [{ kind: "text", text: "fi office", style }];
    const lines = packLines(inline, {
      measurer,
      width: 200,
      lineHeight: ({ naturalPx }) => naturalPx,
    });
    const items = lines.flatMap((line) => line.items).filter((item) => item.kind === "text");
    expect(items.length).toBeGreaterThan(0);
    const withRun = items.filter((item) => item.glyphRun);
    expect(withRun.length).toBeGreaterThan(0);

    // The caret map consumes the produced run: the fi ligature's cluster is
    // divided across its two graphemes.
    const fiItem = withRun.find((item) => item.text.startsWith("fi"))!;
    const layout = itemGlyphLayoutOf(fiItem, style);
    expect(layout.lens.slice(0, 2)).toEqual([1, 1]);
    expect(layout.widths[0]!).toBeGreaterThan(0);
    expect(layout.endX).toBeGreaterThan(0);
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
