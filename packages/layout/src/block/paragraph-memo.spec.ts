import { describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutBlockContext, LayoutParagraph } from "../layout-doc";
import { TextMeasurer } from "../text/measure";
import { layoutParagraph } from "./paragraph";

installFakeCanvas();
const measurer = new TextMeasurer(fakeFontMetrics);

const latin = { family: "serif", sizePx: 16 };

const para = (over: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
  kind: "paragraph",
  inline: [{ kind: "text", text: "word", style: latin }],
  defaultTextStyle: latin,
  ...over,
});

describe("layoutParagraph memo", () => {
  it("re-lays when only a picture's src changed (paint data must not go stale)", () => {
    const picture = (src: string): LayoutParagraph =>
      para({ inline: [{ kind: "picture", widthPx: 120, heightPx: 60, src }] });
    // Prime the memo, then ask for the same geometry with a new source: the
    // hit hands back the laid copy's inline atoms, which the painter reads,
    // so a src-only change must miss instead of painting the old bitmap.
    layoutParagraph(picture("data:image/png;base64,AAAA"), 500, undefined, measurer);
    const next = layoutParagraph(picture("data:image/png;base64,BBBB"), 500, undefined, measurer);
    expect(next.inline[0]).toMatchObject({ src: "data:image/png;base64,BBBB" });
  });

  it("keys on the flow position when absolute float zones are in play", () => {
    const text = Array.from({ length: 40 }, () => "word").join(" ");
    const zones = [{ x0Px: 0, widthPx: 400, topPx: 0, bottomPx: 100 }];
    const ctx = (startY: number): LayoutBlockContext =>
      ({ floatZones: zones, startY, onGrid: true }) as LayoutBlockContext;
    const p = para({ inline: [{ kind: "text", text, style: latin }] });
    // The zone only reaches the startY = 0 position: the memo must not hand
    // the narrow (wrapped) lines back to the later, unobstructed layout.
    const near = layoutParagraph(p, 500, ctx(0), measurer);
    const far = layoutParagraph(p, 500, ctx(1000), measurer);
    expect(far.lines.length).toBeLessThan(near.lines.length);
  });
});
