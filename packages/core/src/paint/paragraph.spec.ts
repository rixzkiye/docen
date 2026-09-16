import type { FontMetrics, LaidOutParagraph, ProjectedFlowBox } from "@docen/layout";
import { Group } from "leafer-ui";
import { describe, expect, it, vi } from "vitest";

import type { PaintContext } from "./context";
import { paintParagraph } from "./paragraph";

/**
 * The tracked-format-change change bar: a 2px vertical rule in the left margin
 * beside every line carrying the revision (w:pPrChange marks all lines,
 * w:rPrChange only the lines its run lands on).
 *
 * leafer-ui is doubled with a shape-recording element stub: the painter's
 * contract is LayoutDoc → Leafer elements 1:1, and the real module builds a
 * canvas context at import time that no headless DOM provides.
 */
vi.mock("leafer-ui", () => {
  class StubElement {
    children: StubElement[] = [];
    constructor(attrs: Record<string, unknown> = {}) {
      Object.assign(this, attrs);
    }
    add(...items: StubElement[]): void {
      this.children.push(...items);
    }
  }
  const make = () => class extends StubElement {};
  // Every leafer-ui binding the painter's module graph imports.
  return {
    Box: make(),
    Ellipse: make(),
    Group: make(),
    Image: make(),
    ImageManager: make(),
    Line: make(),
    Path: make(),
    Rect: make(),
    Resource: make(),
    Text: make(),
  };
});

const flow: ProjectedFlowBox = {
  pageWidthPx: 794,
  pageHeightPx: 1123,
  contentWidthPx: 400,
  contentHeightPx: 900,
  contentLeftPx: 96,
  contentTopPx: 96,
};

const ctx: PaintContext = {
  metrics: {} as FontMetrics,
  flow,
  pageIndex: 0,
  pageCount: 1,
  layer: "body",
  rerender: () => {},
};

const line = (yPx: number, inlineIndex: number, text: string) => ({
  yPx,
  heightPx: 20,
  naturalPx: 20,
  textEmPx: 16,
  endInlineIndex: inlineIndex + 1,
  final: yPx === 0,
  items: [{ kind: "text" as const, inlineIndex, text, xPx: 0, widthPx: 30 }],
  maxWidthPx: 400,
});

const para = (over: Partial<LaidOutParagraph>): LaidOutParagraph => ({
  kind: "paragraph",
  heightPx: 20,
  beforePx: 0,
  afterPx: 0,
  lines: [line(0, 0, "word")],
  inline: [{ kind: "text", text: "word", style: { family: "serif", sizePx: 16 } }],
  ...over,
});

/** The 2px margin bars among a paragraph group's painted children. */
const bars = (tree: Group): { x?: number; y?: number; fill?: string }[] =>
  (tree.children ?? [])
    .map((child) => child as unknown as { x?: number; y?: number; width?: number; fill?: string })
    .filter((child) => child.width === 2);

describe("format-change margin bars", () => {
  it("draws a paragraph-level w:pPrChange bar beside every line", () => {
    const tree = new Group();
    paintParagraph(
      tree,
      para({
        lines: [line(0, 0, "one"), line(20, 0, "two")],
        formatChange: { color: "2E74B5" },
      }),
      40,
      0,
      ctx,
    );
    const drawn = bars(tree);
    expect(drawn).toHaveLength(2);
    expect(drawn.map((bar) => bar.fill)).toEqual(["#2E74B5", "#2E74B5"]);
    // The rule sits just left of the content box (the margin gutter).
    expect(drawn.map((bar) => bar.x)).toEqual([35, 35]);
    expect(drawn.map((bar) => bar.y)).toEqual([0, 20]);
  });

  it("draws a run-level w:rPrChange bar only beside the line its run lands on", () => {
    const tree = new Group();
    paintParagraph(
      tree,
      para({
        lines: [line(0, 0, "one"), line(20, 1, "two")],
        inline: [
          { kind: "text", text: "one", style: { family: "serif", sizePx: 16 } },
          {
            kind: "text",
            text: "two",
            style: { family: "serif", sizePx: 16 },
            formatChange: { color: "FF0000" },
          },
        ],
      }),
      40,
      0,
      ctx,
    );
    const drawn = bars(tree);
    expect(drawn).toHaveLength(1);
    expect(drawn[0]).toMatchObject({ x: 35, y: 20, fill: "#FF0000" });
  });

  it("draws no bar without a format change", () => {
    const tree = new Group();
    paintParagraph(tree, para({}), 40, 0, ctx);
    expect(bars(tree)).toEqual([]);
  });
});
