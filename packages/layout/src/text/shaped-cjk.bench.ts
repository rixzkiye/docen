// Shaped-path benchmark on a Latin + CJK document (the default-flip
// evidence): whole-run width measurement, a keystroke on a 60-paragraph body
// (the incremental layout the editor runs per typed character), and cold
// shaping throughput, canvas vs opt-in shaped.
//
// The canvas side uses the deterministic fake OffscreenCanvas (Node has no
// browser font stack); the shaped side uses OpenSans + the system Noto Sans
// CJK SC collection face — the real bytes. Run with
// `pnpm exec vp test bench shaped-cjk`.

import fs from "node:fs";
import path from "node:path";

import { createFontRefSync, initShapingWasm } from "@docen/shaping";
import { beforeAll, bench, describe } from "vitest";

import { installFakeCanvas } from "../../test/fake-canvas";
import { layoutFlowSections } from "../flow/flow";
import { browserFontMetrics } from "../font";
import type { LayoutParagraph, LayoutTextStyle } from "../layout-doc";
import { TextMeasurer } from "./measure";
import { ShapedMeasurer } from "./shaped-measurer";

const fixtures = path.resolve(__dirname, "../../../shaping/test/fixtures/fonts");
const wasm = path.resolve(__dirname, "../../../shaping/wasm/docen_shaping.wasm");
const CJK_FONT = "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc";

const FAMILY = "Bench Noto CJK";
const LATIN_FAMILY = "Bench Open Sans";
const STYLE: LayoutTextStyle = { family: FAMILY, sizePx: 16 };
const SAMPLE =
  "混排段落 Latin words wrap at spaces while 中文可以在任意字之间断开，标点遵守禁则。 ".repeat(4);
const PAGE = { contentWidthPx: 468, contentHeightPx: 1200 };

const paragraph = (text: string, family = FAMILY): LayoutParagraph => ({
  kind: "paragraph",
  inline: [{ kind: "text", text, style: { family, sizePx: 16 } }],
  defaultTextStyle: { family, sizePx: 16 },
});

let canvas: TextMeasurer;
let shaped: ShapedMeasurer;

beforeAll(async () => {
  if (!fs.existsSync(CJK_FONT)) throw new Error(`missing CJK collection ${CJK_FONT}`);
  installFakeCanvas();
  await initShapingWasm(new Uint8Array(fs.readFileSync(wasm)));
  canvas = new TextMeasurer(browserFontMetrics);
  shaped = new ShapedMeasurer(browserFontMetrics, { enabled: true });
  shaped.registerFont(
    LATIN_FAMILY,
    createFontRefSync(new Uint8Array(fs.readFileSync(path.join(fixtures, "OpenSans-Regular.ttf")))),
  );
  shaped.registerFont(
    FAMILY,
    createFontRefSync(new Uint8Array(fs.readFileSync(CJK_FONT)), { index: 2 }),
  );
});

describe("widthOf, latin+cjk sample", () => {
  bench("canvas (warm)", () => {
    canvas.widthOf(SAMPLE, STYLE);
  });
  bench("shaped (warm)", () => {
    shaped.widthOf(SAMPLE, STYLE);
  });
  bench("canvas (cold)", () => {
    canvas.clearCache();
    canvas.widthOf(SAMPLE, STYLE);
  });
  bench("shaped (cold)", () => {
    shaped.clearCache();
    shaped.widthOf(SAMPLE, STYLE);
  });
});

describe("typing: one edited paragraph in a 60-paragraph body", () => {
  let editSeq = 0;
  const body = (edit: string): LayoutParagraph[] =>
    Array.from({ length: 60 }, (_, i) =>
      paragraph(i === 30 ? `${SAMPLE}${edit}` : `${SAMPLE}${i}`),
    );

  bench("canvas keystroke", () => {
    // A fresh edit string per iteration = a real keystroke (new segment).
    layoutFlowSections([{ blocks: body(`k${editSeq++}`), opts: PAGE }], canvas);
  });
  bench("shaped keystroke", () => {
    layoutFlowSections([{ blocks: body(`k${editSeq++}`), opts: PAGE }], shaped);
  });
});

describe("cold shaping throughput", () => {
  let unique = 0;
  bench("shaped unique run (shape + metrics + cache insert)", () => {
    shaped.clearCache();
    shaped.widthOf(`${SAMPLE} ${unique++}`, STYLE);
  });
});
