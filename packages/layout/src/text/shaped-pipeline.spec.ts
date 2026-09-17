// The R6 pipeline-level gates that need real font-table metrics and the
// shaping WASM: (1) repeated layouts of the same document hash identically
// without relying on the canvas, (2) the shaped measurer's per-keystroke
// layout stays within the documented regression budget of the canvas measurer.
//
// @vitest-environment node
import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createFontRefSync, initShapingWasm } from "@docen/shaping";
import { beforeAll, describe, expect, it } from "vitest";

import { installFakeCanvas } from "../../test/fake-canvas";
import { layoutBlock } from "../block/block";
import { browserFontMetrics } from "../font";
import type { LayoutParagraph } from "../layout-doc";
import { TextMeasurer } from "./measure";
import { ShapedMeasurer } from "./shaped-measurer";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fontsDir = path.resolve(__dirname, "../../../shaping/test/fixtures/fonts");

const paragraph = (text: string, cjk: string): LayoutParagraph => ({
  kind: "paragraph",
  defaultTextStyle: { family: "Calibri", sizePx: 16 },
  inline: [
    { kind: "text", text, style: { family: "Calibri", sizePx: 16 } },
    { kind: "text", text: cjk, style: { family: "SimSun", sizePx: 16 } },
  ],
});

function shapedMeasurer(): ShapedMeasurer {
  const measurer = new ShapedMeasurer(browserFontMetrics, { enabled: true });
  measurer.registerFont(
    "Calibri",
    createFontRefSync(fs.readFileSync(path.join(fontsDir, "OpenSans-Regular.ttf"))),
  );
  measurer.registerFont(
    "SimSun",
    createFontRefSync(fs.readFileSync(path.join(fontsDir, "NotoFangsongKSSVertical-Regular.ttf"))),
  );
  return measurer;
}

/** SHA-256 of a laid result with process-local backend handles removed — the
 *  fontId is a WASM registration number, not document content. */
function layoutHash(laid: unknown): string {
  return createHash("sha256")
    .update(JSON.stringify(laid, (key, value) => (key === "fontId" ? undefined : value)))
    .digest("hex");
}

describe("shaped pipeline gates", () => {
  beforeAll(async () => {
    // The breaker now consumes shaped advances for registered faces (T2);
    // the synthetic fake only backs the canvas fallbacks (glyph painting and
    // unregistered faces) and keeps Node deterministic.
    installFakeCanvas();
    await initShapingWasm();
  });

  it("hashes identical LayoutDoc across repeated runs with font-table metrics", () => {
    const doc = paragraph("Deterministic office text", "中文测试");
    const measurer = shapedMeasurer();

    const first = layoutBlock(doc, 600, undefined, measurer);
    measurer.clearCache();
    const second = layoutBlock(doc, 600, undefined, measurer);
    // A fresh measurer + fresh WASM registration must lay out identically
    // (only the volatile fontId differs and is normalized out of the hash).
    const third = layoutBlock(doc, 600, undefined, shapedMeasurer());

    const hashA = layoutHash(first);
    expect(layoutHash(second)).toBe(hashA);
    expect(layoutHash(third)).toBe(hashA);

    const json = JSON.stringify(first);
    expect(json).toContain("glyphRun");
    expect(json).toContain("totalAdvancePx");
  });

  it("keeps per-keystroke shaped layout within the documented regression budget", () => {
    const base = "The quick brown fox jumps over the lazy dog";
    const style = { family: "Calibri", sizePx: 16 };
    const canvas = new TextMeasurer(browserFontMetrics);
    const shaped = shapedMeasurer();

    const keystrokes = 120;
    const run = (measurer: TextMeasurer): number => {
      const t0 = performance.now();
      for (let i = 1; i <= keystrokes; i++) {
        layoutBlock(
          {
            kind: "paragraph",
            defaultTextStyle: style,
            inline: [{ kind: "text", text: base.slice(0, i), style }],
          },
          600,
          undefined,
          measurer,
        );
      }
      return performance.now() - t0;
    };

    // Warm both paths (WASM is cold on the first shape).
    run(canvas);
    run(shaped);
    const ratios: number[] = [];
    for (let round = 0; round < 5; round++) {
      const canvasMs = run(canvas);
      const shapedMs = run(shaped);
      ratios.push(shapedMs / canvasMs);
    }
    const sorted = [...ratios].sort((a, b) => a - b);
    const best = sorted[0]!;
    const median = sorted[Math.floor(sorted.length / 2)]!;
    console.log(
      `per-keystroke layout ratio (shaped/canvas): best ${best.toFixed(2)}, median ${median.toFixed(2)} (target <= 1.10; bench runs concurrently with the suite)`,
    );

    // The plan's T1 target on the isolated comparison (best round — suite runs
    // concurrently), with a loose under-load sanity bound so a real regression
    // cannot hide in the noise.
    expect(best).toBeLessThanOrEqual(1.1);
    expect(median).toBeLessThanOrEqual(3);
  });
});
