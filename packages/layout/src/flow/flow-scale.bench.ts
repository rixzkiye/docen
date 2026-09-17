// Flow-scale benchmark: how the paged layout walk scales with document size.
// Each 60-line paragraph fills a 1200px page, so `N pages` below is a
// one-paragraph-per-page body — the README's large-document shape at the
// layout layer. `cold` gives every run fresh content (the open/paste path);
// `incremental` re-lays a body whose paragraphs are structurally unchanged
// (the typing path — the paragraph memo hands back the previous result).
// Run with `pnpm exec vp test bench`.

import { bench, describe } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutParagraph } from "../layout-doc";
import { TextMeasurer } from "../text/measure";
import { layoutFlowSections } from "./flow";

installFakeCanvas();
const measurer = new TextMeasurer(fakeFontMetrics);

const latin = { family: "serif", sizePx: 16 };
const exact20 = { lineHeight: { rule: "exact" as const, px: 20 }, beforePx: 0, afterPx: 0 };
const PAGE = { contentWidthPx: 468, contentHeightPx: 1200 };

/** A paragraph of `lines` hard-broken lines — exact 20px lines keep the
 *  layout arithmetic and immune to font metrics (flow.spec's shape). The
 *  `salt` marks cold content so the memo never serves it; the per-index text
 *  keeps paragraphs distinct within a body. */
const para = (lines: number, salt = 0, index = 0): LayoutParagraph => {
  const text = `x${salt}-${index}`;
  return {
    kind: "paragraph",
    inline: [
      { kind: "text", text, style: latin },
      ...Array.from({ length: lines - 1 }, () => [
        { kind: "break" as const },
        { kind: "text" as const, text, style: latin },
      ]).flat(),
    ],
    spacing: exact20,
    defaultTextStyle: latin,
    widowControl: false,
  };
};

const body = (paragraphs: number, salt: number): LayoutParagraph[] =>
  Array.from({ length: paragraphs }, (_, i) => para(60, salt, i));

let coldSalt = 0;

describe("flow scale (one 60-line paragraph per page)", () => {
  for (const pages of [100, 250, 500]) {
    bench(`${pages} pages, cold`, () => {
      coldSalt++;
      layoutFlowSections([{ blocks: body(pages, coldSalt), opts: PAGE }], measurer);
    });
    bench(`${pages} pages, incremental`, () => {
      layoutFlowSections([{ blocks: body(pages, 0), opts: PAGE }], measurer);
    });
  }
});
