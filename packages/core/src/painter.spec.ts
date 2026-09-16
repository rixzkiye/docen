import type { FontMetrics, LaidOutBalloon, ProjectedFlowBox } from "@docen/layout";
import { Group } from "leafer-ui";
import { describe, expect, it, vi } from "vitest";

import type { PaintContext } from "./paint/context";
import { balloonAt, paintBalloons } from "./painter";

/**
 * Balloon painting: cards in the page margin (rounded box, author label in
 * the revision color, wrapped body lines, connector to the anchored line) plus
 * the page-local hit boxes the stage routes clicks with. leafer-ui is doubled
 * with a shape-recording stub — the painter's contract is LayoutDoc → Leafer
 * elements 1:1 (the real module builds a canvas context at import time).
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
  pageWidthPx: 500,
  pageHeightPx: 700,
  contentWidthPx: 300,
  contentHeightPx: 600,
  contentLeftPx: 50,
  contentTopPx: 20,
};

const ctx: PaintContext = {
  metrics: {} as FontMetrics,
  flow,
  pageIndex: 0,
  pageCount: 1,
  layer: "body",
  rerender: () => {},
};

const card = (over: Partial<LaidOutBalloon> = {}): LaidOutBalloon => ({
  id: 7,
  kind: "comment",
  color: "2E74B5",
  label: "AL",
  lines: ["note body"],
  xPx: 312,
  yPx: 40,
  widthPx: 130,
  heightPx: 44,
  anchorXPx: 300,
  anchorYPx: 50,
  ...over,
});

/** Every painted element, depth first. */
const flatten = (group: Group): Record<string, unknown>[] => {
  const out: Record<string, unknown>[] = [];
  const walk = (node: { children?: unknown[] }): void => {
    for (const child of node.children ?? []) {
      out.push(child as Record<string, unknown>);
      walk(child as { children?: unknown[] });
    }
  };
  walk(group);
  return out;
};

describe("paintBalloons", () => {
  it("paints the card in the margin with its label, body and connector", () => {
    const tree = new Group();
    const painted = paintBalloons(tree, [card()], ctx);
    const elements = flatten(tree);
    // Card box + its page-local placement (the hit box carries the geometry).
    const box = elements.find((el) => el.width === 130 && el.height === 44);
    expect(box).toMatchObject({ stroke: "#2E74B5", cornerRadius: 3 });
    const group = elements.find((el) => el.x === 50 + 312 && el.y === 20 + 40);
    expect(group).toBeDefined();
    // Author header in the accent color, body line below it.
    const texts = elements.filter((el) => typeof el.text === "string");
    expect(texts.map((el) => el.text)).toEqual(["AL", "note body"]);
    expect(texts[0]).toMatchObject({ fill: "#2E74B5" });
    // Connector from the content edge's anchor to the card's left edge.
    const line = elements.find((el) => Array.isArray(el.points));
    expect(line?.points).toEqual([50 + 300, 20 + 50, 50 + 312, 20 + 40 + 22]);
    // Hit box, page-local.
    expect(painted.boxes).toEqual([
      { id: 7, kind: "comment", x: 362, y: 60, width: 130, height: 44 },
    ]);
    expect([...painted.groups.keys()]).toEqual(["comment:7"]);
  });

  it("keys hover groups per kind:id and paints every card", () => {
    const tree = new Group();
    const painted = paintBalloons(
      tree,
      [
        card(),
        card({ id: 7, kind: "revision", color: "FF0000", label: "Ada", lines: [] }),
        card({ id: 9, kind: "comment", color: "FF0000", label: "", lines: [] }),
      ],
      ctx,
    );
    expect([...painted.groups.keys()]).toEqual(["comment:7", "revision:7", "comment:9"]);
    expect(painted.boxes.map((box) => box.id)).toEqual([7, 7, 9]);
  });

  it("paints nothing without balloons", () => {
    const tree = new Group();
    const painted = paintBalloons(tree, undefined, ctx);
    expect(flatten(tree)).toEqual([]);
    expect(painted.boxes).toEqual([]);
    expect(painted.groups.size).toBe(0);
  });
});

describe("balloonAt", () => {
  it("returns the topmost box under the point and null off every card", () => {
    const boxes = [
      { id: 1, kind: "comment" as const, x: 10, y: 10, width: 50, height: 20 },
      { id: 2, kind: "revision" as const, x: 40, y: 15, width: 50, height: 20 },
    ];
    expect(balloonAt(boxes, 45, 20)?.id).toBe(2);
    expect(balloonAt(boxes, 15, 20)?.id).toBe(1);
    expect(balloonAt(boxes, 5, 20)).toBeNull();
    expect(balloonAt(boxes, 45, 50)).toBeNull();
  });
});
