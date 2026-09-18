// T2 breaker gate: with a registered shaped font, the line breaker must wrap
// on shaped advances, not the canvas metrics. The oracle is the measurers'
// own widthOf — one shaped, one canvas — so the test proves the breaker
// consumed the shaped path without re-implementing the packer.

// @vitest-environment node
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createFontRefSync, initShapingWasm } from "@docen/shaping";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { installFakeCanvas } from "../../test/fake-canvas";
import { layoutBlock } from "../block/block";
import { browserFontMetrics, clearRegisteredFontRatios } from "../font";
import type { LayoutBlock, LayoutParagraph, LayoutTextStyle } from "../layout-doc";
import { TextMeasurer } from "./measure";
import { ShapedMeasurer } from "./shaped-measurer";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fontPath = path.resolve(
  __dirname,
  "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
);

const FAMILY = "Open Sans Breaker";
const STYLE: LayoutTextStyle = { family: FAMILY, sizePx: 16 };
const WIDTH = 200;
/** One unbreakable run: pretext segments it as a single word, so the shaped
 *  whole-run advance is exactly the segment advance. */
const WORD = "AV".repeat(60);

const paragraph = (): LayoutParagraph => ({
  kind: "paragraph",
  defaultTextStyle: STYLE,
  inline: [{ kind: "text", text: WORD, style: STYLE }],
});

const firstLineText = (laid: unknown): string => {
  const lines = (laid as { lines: { items: { text?: string }[] }[] }).lines;
  return lines[0]!.items.map((i) => i.text ?? "").join("");
};

describe("shaped breaker (T2)", () => {
  let canvas: TextMeasurer;
  let shaped: ShapedMeasurer;

  beforeAll(async () => {
    installFakeCanvas();
    await initShapingWasm();
    canvas = new TextMeasurer(browserFontMetrics);
    shaped = new ShapedMeasurer(browserFontMetrics, { enabled: true });
    shaped.registerFont(FAMILY, createFontRefSync(fs.readFileSync(fontPath)));
  });

  afterAll(() => {
    clearRegisteredFontRatios();
  });

  it("wraps on shaped advances, not the canvas ones", () => {
    // Sanity: the two measurement sources disagree on this run (fake canvas
    // measures 10px/char; OpenSans shapes narrower), so a differing break
    // proves which source the packer used.
    expect(canvas.widthOf(WORD, STYLE)).not.toBeCloseTo(shaped.widthOf(WORD, STYLE), 0);

    const canvasLaid = layoutBlock(paragraph() as LayoutBlock, WIDTH, undefined, canvas);
    const shapedLaid = layoutBlock(paragraph() as LayoutBlock, WIDTH, undefined, shaped);
    const canvasFirst = firstLineText(canvasLaid);
    const shapedFirst = firstLineText(shapedLaid);

    // The shaped path measures this run differently, so a differing break
    // proves which source the packer consumed.
    expect(shapedFirst).not.toBe(canvasFirst);

    // The shaped break fits under the shaped metric…
    expect(shaped.widthOf(shapedFirst, STYLE)).toBeLessThanOrEqual(WIDTH + 0.5);
    // …while the canvas-chosen break overflows when measured shaped — i.e.
    // the packer did NOT use the canvas widths.
    expect(shaped.widthOf(canvasFirst, STYLE)).toBeGreaterThan(WIDTH);
  });
});
