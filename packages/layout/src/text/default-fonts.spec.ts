import { createHash } from "node:crypto";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { FontManager } from "@docen/shaping";
import { describe, expect, it, vi } from "vitest";

import { browserFontMetrics } from "../font";
import { DEFAULT_FONT_FACES, registerDefaultFonts } from "./default-fonts";
import { ShapedMeasurer, clearMissingFontWarnings, createMeasurer } from "./shaped-measurer";

/**
 * Cross-environment font-metrics golden (item 10).
 *
 * The production font bytes and the shaped runs they produce are hashed into
 * `test/font-metrics-golden.json`. The hash covers only deterministic inputs
 * (bundled bytes + fixed text/size/style + glyph ids/advances/clusters), so
 * the same golden verifies on any machine and any browser: run
 * `pnpm exec vp test run packages/layout/src/text/default-fonts.spec.ts`.
 *
 * Regenerate after an intentional font/engine change:
 * `DOCEN_FONT_GOLDEN_UPDATE=1 pnpm exec vp test run packages/layout/src/text/default-fonts.spec.ts`
 */

const GOLDEN_PATH = fileURLToPath(new URL("../../test/font-metrics-golden.json", import.meta.url));
const FONT_ROOT = new URL("../../assets/fonts/", import.meta.url);
const WASM_PATH = new URL("../../../shaping/wasm/docen_shaping.wasm", import.meta.url);
const UPDATE = process.env.DOCEN_FONT_GOLDEN_UPDATE === "1";

interface GoldenCase {
  readonly id: string;
  readonly family: string;
  readonly text: string;
  readonly sizePx: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** Deterministic hash of the shaped run. */
  readonly hash: string;
  readonly glyphCount: number;
  readonly totalAdvancePx: number;
}

interface GoldenFile {
  readonly version: number;
  readonly procedure: string;
  readonly wasmSha256: string;
  readonly fonts: Readonly<Record<string, string>>;
  readonly cases: readonly GoldenCase[];
}

function sha256(bytes: Uint8Array | string): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function canonicalHash(run: NonNullable<ReturnType<ShapedMeasurer["shapeRun"]>>): string {
  const payload = [
    (run.fontName ?? "").toLowerCase(),
    run.fontSizePx,
    run.unitsPerEm,
    run.totalAdvancePx.toFixed(6),
    ...run.glyphs.map((g) =>
      [
        g.glyphId,
        g.xAdvance,
        g.yAdvance,
        g.xOffset,
        g.yOffset,
        g.cluster,
        g.xPx.toFixed(6),
        g.yPx.toFixed(6),
      ].join(":"),
    ),
  ].join("|");
  return sha256(payload);
}

/** The sample matrix: default families × styles, plus script/feature probes. */
const CASES: readonly Omit<GoldenCase, "hash" | "glyphCount" | "totalAdvancePx">[] = [
  { id: "calibri-kern-liga", family: "Calibri", text: "AVATAR office fi fl", sizePx: 16 },
  { id: "calibri-bold", family: "Calibri", text: "AVATAR office fi fl", sizePx: 16, bold: true },
  {
    id: "calibri-italic",
    family: "Calibri",
    text: "AVATAR office fi fl",
    sizePx: 16,
    italic: true,
  },
  {
    id: "calibri-bold-italic",
    family: "Calibri",
    text: "AVATAR office fi fl",
    sizePx: 16,
    bold: true,
    italic: true,
  },
  { id: "calibri-light", family: "Calibri Light", text: "Heading one two", sizePx: 22 },
  { id: "cambria", family: "Cambria", text: "The quick brown fox", sizePx: 16 },
  { id: "arial", family: "Arial", text: "The quick brown fox", sizePx: 16 },
  { id: "times", family: "Times New Roman", text: "The quick brown fox", sizePx: 16 },
  { id: "latin-accents", family: "Calibri", text: "Cześć naïve Ærø", sizePx: 16 },
  { id: "cyrillic", family: "Arial", text: "Привет мир", sizePx: 16 },
  { id: "digits", family: "Cambria", text: "0123456789", sizePx: 11 },
];

async function computeCases(): Promise<{ cases: GoldenCase[]; fonts: Record<string, string> }> {
  await registerDefaultFonts();
  const measurer = new ShapedMeasurer(browserFontMetrics, { enabled: true });
  const cases = CASES.map((sample) => {
    const run = measurer.shapeRun(sample.text, {
      family: sample.family,
      sizePx: sample.sizePx,
      ...(sample.bold ? { bold: true } : {}),
      ...(sample.italic ? { italic: true } : {}),
    });
    if (!run) throw new Error(`shaping produced no run for ${sample.id}`);
    return {
      ...sample,
      hash: canonicalHash(run),
      glyphCount: run.glyphs.length,
      totalAdvancePx: Number(run.totalAdvancePx.toFixed(6)),
    };
  });
  const fonts: Record<string, string> = {};
  for (const face of new Set(DEFAULT_FONT_FACES.map((f) => f.file))) {
    fonts[face] = sha256(new Uint8Array(readFileSync(fileURLToPath(new URL(face, FONT_ROOT)))));
  }
  return { cases, fonts };
}

describe("production font set (item 10)", () => {
  it("golden: font bytes and shaped metrics hash as recorded", async () => {
    const { cases, fonts } = await computeCases();
    const wasmSha256 = sha256(new Uint8Array(readFileSync(fileURLToPath(WASM_PATH))));
    const computed: GoldenFile = {
      version: 1,
      procedure:
        "DOCEN_FONT_GOLDEN_UPDATE=1 pnpm exec vp test run packages/layout/src/text/default-fonts.spec.ts",
      wasmSha256,
      fonts,
      cases,
    };

    if (UPDATE) {
      writeFileSync(GOLDEN_PATH, `${JSON.stringify(computed, null, 2)}\n`);
      return;
    }

    if (!existsSync(GOLDEN_PATH)) {
      throw new Error(
        `font golden missing at ${GOLDEN_PATH} — regenerate with DOCEN_FONT_GOLDEN_UPDATE=1`,
      );
    }
    const golden = JSON.parse(readFileSync(GOLDEN_PATH, "utf8")) as GoldenFile;
    expect(golden.version).toBe(1);
    expect(golden.wasmSha256).toBe(wasmSha256);
    expect(golden.fonts).toEqual(fonts);
    // Compare per case for a readable failure diff.
    for (const expected of golden.cases) {
      const actual = computed.cases.find((c) => c.id === expected.id);
      expect(actual, `case ${expected.id}`).toMatchObject({
        family: expected.family,
        text: expected.text,
        sizePx: expected.sizePx,
        hash: expected.hash,
        glyphCount: expected.glyphCount,
        totalAdvancePx: expected.totalAdvancePx,
      });
    }
    expect(computed.cases).toHaveLength(golden.cases.length);
  });

  it("registers the bundled default families and selects weight/slant faces", async () => {
    const registered = await registerDefaultFonts();
    expect(registered).toContain("Calibri bold italic");
    const measurer = new ShapedMeasurer(browserFontMetrics, { enabled: true });
    const regular = measurer.getFont("Calibri");
    const bold = measurer.getFont("Calibri", { bold: true });
    const italic = measurer.getFont("Calibri", { italic: true });
    expect(regular).toBeDefined();
    expect(bold).toBeDefined();
    expect(italic).toBeDefined();
    expect(bold).not.toBe(regular);
    expect(italic).not.toBe(regular);
    // A style with no exact slot falls back to the regular face, not canvas.
    const noFallback = measurer.shapeRun("x", {
      family: "Cambria",
      sizePx: 16,
      fontWeight: 900,
    });
    expect(noFallback).toBeDefined();
  });

  it("warns once per family when shaping falls back to canvas", () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      clearMissingFontWarnings();
      const isolated = new ShapedMeasurer(browserFontMetrics, {
        enabled: true,
        fontManager: new FontManager({ enableOpfs: false }),
      });
      const style = { family: "No Such Face", sizePx: 16 };
      expect(isolated.shapeRun("hello", style)).toBeUndefined();
      expect(isolated.shapeRun("hello again", style)).toBeUndefined();
      expect(warn).toHaveBeenCalledTimes(1);
      expect(warn.mock.calls[0]?.[0]).toContain("No Such Face");
    } finally {
      warn.mockRestore();
    }
  });

  it("keeps the DOCEN_SHAPING_DISABLED=1 rollback silent", () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const previous = process.env.DOCEN_SHAPING_DISABLED;
    try {
      process.env.DOCEN_SHAPING_DISABLED = "1";
      clearMissingFontWarnings();
      const measurer = createMeasurer(browserFontMetrics);
      expect(measurer).not.toBeInstanceOf(ShapedMeasurer);
      // A forced-enabled instance also reports disabled through shapeRun: the
      // env rollback is checked before the per-instance flag.
      expect(
        new ShapedMeasurer(browserFontMetrics).shapeRun("hello", {
          family: "Calibri",
          sizePx: 16,
        }),
      ).toBeUndefined();
      expect(warn).not.toHaveBeenCalled();
    } finally {
      if (previous === undefined) delete process.env.DOCEN_SHAPING_DISABLED;
      else process.env.DOCEN_SHAPING_DISABLED = previous;
      warn.mockRestore();
    }
  });
});
