import { describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutBalloonAnchor, LayoutParagraph } from "../layout-doc";
import { TextMeasurer } from "../text/measure";
import { layoutFlow } from "./flow";

/**
 * The right-margin balloon stack: anchors resolve to their line's center,
 * cards pack top-down without overlapping, and body text geometry is
 * untouched.
 */

installFakeCanvas();
const measurer = new TextMeasurer(fakeFontMetrics);

const latin = { family: "serif", sizePx: 16 };
const LINE = 20;
const exact20 = { lineHeight: { rule: "exact" as const, px: LINE }, beforePx: 0, afterPx: 0 };

/** A paragraph with exactly `n` lines (hard breaks); inline indices run
 *  0, 2, 4 … (text, break, text, …). */
const para = (n: number, extra: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
  kind: "paragraph",
  inline: [
    { kind: "text", text: "x", style: latin },
    ...Array.from({ length: n - 1 }, () => [
      { kind: "break" as const },
      { kind: "text" as const, text: "x", style: latin },
    ]).flat(),
  ],
  spacing: exact20,
  defaultTextStyle: latin,
  widowControl: false,
  ...extra,
});

const anchor = (over: Partial<LayoutBalloonAnchor> = {}): LayoutBalloonAnchor => ({
  id: 1,
  kind: "revision",
  color: "FF0000",
  label: "Ada",
  text: "changed",
  inlineIndex: 0,
  ...over,
});

/** Page size 500×N with a 50px left margin and a 300px content column — the
 *  150px right margin band holds 130px-wide cards (under the 200px cap). */
const flow = (blocks: Parameters<typeof layoutFlow>[0], height = 100) =>
  layoutFlow(
    blocks,
    { contentWidthPx: 300, contentHeightPx: height, pageWidthPx: 500, contentLeftPx: 50 },
    measurer,
  );

describe("balloon packing", () => {
  it("packs one card per anchor in the right margin, in anchor order", () => {
    const pages = flow([
      para(3, {
        balloons: [
          anchor({ id: 3, inlineIndex: 4, text: "third" }),
          anchor({ id: 1, inlineIndex: 0, text: "first" }),
          anchor({ id: 2, inlineIndex: 2, text: "second" }),
        ],
      }),
    ]);
    expect(pages).toHaveLength(1);
    const balloons = pages[0]!.balloons!;
    // Sorted by anchor Y regardless of the input order.
    expect(balloons.map((b) => b.id)).toEqual([1, 2, 3]);
    // Cards live right of the content column, inside the page band.
    for (const card of balloons) {
      expect(card.xPx).toBe(300 + 12);
      expect(card.widthPx).toBe(130);
      expect(card.anchorXPx).toBe(300);
    }
    // Anchor centers track the 20px lines.
    expect(balloons.map((b) => b.anchorYPx)).toEqual([10, 30, 50]);
  });

  it("never overlaps cards that share a crowded line", () => {
    const pages = flow([
      para(1, {
        balloons: [
          anchor({ id: 1, inlineIndex: 0, text: "one two three four five six seven eight" }),
          anchor({ id: 2, inlineIndex: 0, text: "nine ten eleven twelve thirteen" }),
          anchor({ id: 3, kind: "comment", color: "2E74B5", inlineIndex: 0, text: "note" }),
        ],
      }),
    ]);
    const balloons = pages[0]!.balloons!;
    expect(balloons).toHaveLength(3);
    for (let i = 1; i < balloons.length; i++) {
      const previous = balloons[i - 1]!;
      expect(balloons[i]!.yPx).toBeGreaterThanOrEqual(previous.yPx + previous.heightPx);
    }
  });

  it("stacks ten balloons on one crowded line without overlap", () => {
    const anchors = Array.from({ length: 10 }, (_, index) =>
      anchor({ id: index + 1, inlineIndex: 0, text: `note ${index + 1}` }),
    );
    const pages = flow([para(1, { balloons: anchors })]);
    const balloons = pages[0]!.balloons!;
    expect(balloons).toHaveLength(10);
    expect(balloons.map((b) => b.id)).toEqual(Array.from({ length: 10 }, (_, i) => i + 1));
    for (let i = 1; i < balloons.length; i++) {
      const previous = balloons[i - 1]!;
      expect(balloons[i]!.yPx).toBeGreaterThanOrEqual(previous.yPx + previous.heightPx);
    }
  });

  it("follows a range's line onto the page it lands on", () => {
    // 12 lines → 5 per 100px page; the anchor sits on line 7 (inline index 14).
    const pages = flow([para(12, { balloons: [anchor({ id: 7, inlineIndex: 14 })] })]);
    expect(pages).toHaveLength(3);
    expect(pages[0]!.balloons).toBeUndefined();
    const tail = pages[1]!.balloons!;
    expect(tail).toHaveLength(1);
    expect(tail[0]!.id).toBe(7);
    expect(tail[0]!.anchorYPx).toBeGreaterThanOrEqual(0);
    expect(tail[0]!.anchorYPx).toBeLessThanOrEqual(100);
    expect(pages[2]!.balloons).toBeUndefined();
  });

  it("keeps body text geometry identical with and without anchors", () => {
    const blocks = [para(3), para(12), para(2)];
    const plain = flow(blocks);
    const anchored = flow([
      para(3, { balloons: [anchor({ id: 1, inlineIndex: 0 })] }),
      para(12, { balloons: [anchor({ id: 2, inlineIndex: 14 })] }),
      para(2, { balloons: [anchor({ id: 3, inlineIndex: 0 })] }),
    ]);
    const geometry = (pages: ReturnType<typeof flow>) =>
      pages.map((page) =>
        page.items.map((item) => ({ yPx: item.yPx, heightPx: item.block.heightPx })),
      );
    expect(geometry(anchored)).toEqual(geometry(plain));
    expect(anchored).toHaveLength(plain.length);
    // The line geometry itself is untouched too.
    const lineGeometry = (pages: ReturnType<typeof flow>) =>
      pages.flatMap((page) =>
        page.items.flatMap((item) =>
          item.block.kind === "paragraph" ? item.block.lines.map((line) => line.yPx) : [],
        ),
      );
    expect(lineGeometry(anchored)).toEqual(lineGeometry(plain));
  });

  it("packs comments and revisions together and keeps their colors", () => {
    const pages = flow([
      para(2, {
        balloons: [
          anchor({ id: 1, kind: "comment", color: "2E74B5", label: "AL", inlineIndex: 0 }),
          anchor({ id: 2, kind: "revision", color: "FF0000", inlineIndex: 2 }),
        ],
      }),
    ]);
    const balloons = pages[0]!.balloons!;
    expect(balloons.map((b) => b.kind)).toEqual(["comment", "revision"]);
    expect(balloons.map((b) => b.color)).toEqual(["2E74B5", "FF0000"]);
  });

  it("carries no balloons for empty pages or undecorated paragraphs", () => {
    const pages = flow([para(2), { kind: "pageBreak" }, para(1)]);
    expect(pages.length).toBeGreaterThan(1);
    for (const page of pages) expect(page.balloons).toBeUndefined();
  });
});
