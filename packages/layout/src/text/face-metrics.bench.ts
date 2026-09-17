// Face-metrics seam benchmark: the DOM span probe (canvas era) against
// fontations-backed face metrics for fonts registered with the shaping font
// manager. The two sides use distinct family names (an unregistered probe
// family can never resolve through the manager), so the comparison isolates
// the metrics source: one append/getBoundingClientRect/remove cycle per
// probe vs one resolver hit + formula over the same number of faces.
//
// Node has no layout engine, so the probe's rect is stubbed; a browser
// additionally pays a forced style/layout per probe. Treat the DOM side as a
// lower bound and the call-count evidence (see face-metrics-resolver.spec.ts)
// as the authoritative "no DOM for registered faces" proof.
//
// Run with `pnpm exec vp test bench face-metrics`.

import fs from "node:fs";
import path from "node:path";

import { createFontRefSync, initShapingWasm } from "@docen/shaping";
import { beforeAll, bench, describe } from "vitest";

import { installFakeCanvas } from "../../test/fake-canvas";
import { layoutBlock } from "../block/block";
import { browserFontMetrics, clearFontMetricCache } from "../font";
import type { LayoutParagraph } from "../layout-doc";
import { TextMeasurer } from "./measure";
import { getShapingFontManager } from "./shaped-measurer";

const fixtures = path.resolve(__dirname, "../../../shaping/test/fixtures/fonts");
const wasm = path.resolve(__dirname, "../../../shaping/wasm/docen_shaping.wasm");

/** Registered faces (real bytes, resolver-backed metrics). */
const FACES: [family: string, file: string][] = [
  ["Open Sans", "OpenSans-Regular.ttf"],
  ["Noto Sans Hebrew", "NotoSansHebrew-Regular.ttf"],
  ["Noto Sans Thai", "NotoSansThai-Regular.ttf"],
  ["Noto Sans Khmer", "NotoSansKhmer-Regular.ttf"],
  ["Noto Sans Devanagari", "NotoSansDevanagari-Regular.ttf"],
];
/** Never registered — every metric goes through the table/probe path. */
const PROBE_FACES = FACES.map((_, i) => `Probe Face ${i}`);

const SAMPLE = "The quick brown fox 文本换行 布局引擎 ";
const style = (family: string): { family: string; sizePx: number } => ({ family, sizePx: 16 });

/** One paragraph cycling every face through several runs. */
const paragraph = (families: readonly string[]): LayoutParagraph => ({
  kind: "paragraph",
  inline: Array.from({ length: 40 }, (_, i) => {
    const family = families[i % families.length]!;
    return i % 3 === 2
      ? { kind: "break" as const }
      : { kind: "text" as const, text: `${SAMPLE}${i} `, style: style(family) };
  }),
  defaultTextStyle: style(families[0]!),
});

let probeCalls = 0;

/** The minimal DOM the metrics probe touches, with a 1.35em rect. */
function installFakeDom(): void {
  const canvasCtx = {
    font: "",
    measureText: () => ({ width: 0, fontBoundingBoxAscent: 0, fontBoundingBoxDescent: 0 }),
  };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- duck-typed probe surface
  (globalThis as any).document = {
    createElement: (tag: string) =>
      tag === "span"
        ? {
            style: {},
            textContent: "",
            getBoundingClientRect: () => {
              probeCalls++;
              return { height: 135, width: 0 };
            },
            remove: () => {},
          }
        : { getContext: () => canvasCtx },
    body: { append: () => {} },
  };
}

beforeAll(async () => {
  installFakeCanvas();
  installFakeDom();
  await initShapingWasm(new Uint8Array(fs.readFileSync(wasm)));
  for (const [family, file] of FACES) {
    getShapingFontManager().registerActiveFont(
      family,
      createFontRefSync(new Uint8Array(fs.readFileSync(path.join(fixtures, file)))),
    );
  }
});

describe("face metrics resolution, 5 faces, cold", () => {
  bench("via DOM probe", () => {
    clearFontMetricCache();
    probeCalls = 0;
    for (const family of PROBE_FACES) browserFontMetrics.normalRatio({ family });
    if (probeCalls !== 5) throw new Error(`expected 5 DOM probes, got ${probeCalls}`);
  });
  bench("via fontations resolver", () => {
    clearFontMetricCache();
    probeCalls = 0;
    for (const [family] of FACES) browserFontMetrics.normalRatio({ family });
    if (probeCalls !== 0) throw new Error(`expected 0 DOM probes, got ${probeCalls}`);
  });
});

describe("paragraph layout, 5 mixed faces, cold measurer + cold metrics", () => {
  bench("via DOM probe metrics", () => {
    clearFontMetricCache();
    const measurer = new TextMeasurer(browserFontMetrics);
    layoutBlock(paragraph(PROBE_FACES), 468, undefined, measurer);
  });
  bench("via fontations metrics", () => {
    clearFontMetricCache();
    const measurer = new TextMeasurer(browserFontMetrics);
    layoutBlock(paragraph(FACES.map(([f]) => f)), 468, undefined, measurer);
  });
});
