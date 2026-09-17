// Measurer comparison: the canvas TextMeasurer against the opt-in
// ShapedMeasurer (rustybuzz + fontations) over Latin and complex-script
// (Arabic, cursive joining + marks) text. The warm variants measure the
// typing steady state (cached strings); the `unique` variants force a cache
// miss every run — the open/paste/first-keystroke shape. Run with
// `pnpm exec vp test bench shaped-vs-canvas`.

import fs from "node:fs";
import path from "node:path";

import { createFontRefSync, initShapingWasm } from "@docen/shaping";
import { beforeAll, bench, describe } from "vitest";

import { installFakeCanvas } from "../../test/fake-canvas";
import { layoutBlock } from "../block/block";
import { browserFontMetrics } from "../font";
import type { LayoutBlock, LayoutParagraph } from "../layout-doc";
import { TextMeasurer } from "./measure";
import { ShapedMeasurer } from "./shaped-measurer";

const fontsDir = path.resolve(__dirname, "../../../shaping/test/fixtures/fonts");
const LATIN =
  "The quick brown fox jumps over the lazy dog. AVWA To fi ffi — office typography. ".repeat(8);
const ARABIC = "النص العربي مع تشكيل ووصل الحروف ".repeat(8);

beforeAll(async () => {
  installFakeCanvas();
  await initShapingWasm();
  shaped = new ShapedMeasurer(browserFontMetrics, { enabled: true });
  shaped.registerFont(
    "Open Sans",
    createFontRefSync(fs.readFileSync(path.join(fontsDir, "OpenSans-Regular.ttf"))),
  );
  shaped.registerFont(
    "Noto Naskh Arabic",
    createFontRefSync(fs.readFileSync(path.join(fontsDir, "NotoNaskhArabic-Regular.ttf"))),
  );
});

const canvas = new TextMeasurer(browserFontMetrics);
let shaped: ShapedMeasurer;

const style = (family: string): { family: string; sizePx: number } => ({ family, sizePx: 16 });

const paragraph = (family: string): LayoutParagraph => ({
  kind: "paragraph",
  inline: [{ kind: "text", text: family === "Open Sans" ? LATIN : ARABIC, style: style(family) }],
  defaultTextStyle: style(family),
});

let unique = 0;

describe("shaped vs canvas measurement", () => {
  bench("canvas widthOf (latin, warm)", () => {
    canvas.widthOf(LATIN, style("Open Sans"));
  });
  bench("shaped widthOf (latin, warm)", () => {
    shaped.widthOf(LATIN, style("Open Sans"));
  });
  bench("canvas widthOf (latin, unique)", () => {
    canvas.widthOf(`${LATIN}${unique++}`, style("Open Sans"));
  });
  bench("shaped widthOf (latin, unique)", () => {
    shaped.widthOf(`${LATIN}${unique++}`, style("Open Sans"));
  });
  bench("canvas widthOf (arabic, warm)", () => {
    canvas.widthOf(ARABIC, style("Noto Naskh Arabic"));
  });
  bench("shaped widthOf (arabic, warm)", () => {
    shaped.widthOf(ARABIC, style("Noto Naskh Arabic"));
  });
  bench("canvas layoutBlock (latin paragraph)", () => {
    layoutBlock(paragraph("Open Sans") as LayoutBlock, 468, undefined, canvas);
  });
  bench("shaped layoutBlock (latin paragraph)", () => {
    layoutBlock(paragraph("Open Sans") as LayoutBlock, 468, undefined, shaped);
  });
});
