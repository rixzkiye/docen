// @vitest-environment node
import type { LaidOutMathItem } from "@docen/layout";
import { describe, expect, it, vi } from "vitest";

import { paintMath } from "./math";

vi.mock("leafer-ui", () => {
  class StubElement {
    type: string;
    props: Record<string, unknown>;
    children: StubElement[] = [];
    constructor(type: string, attrs: Record<string, unknown> = {}) {
      this.type = type;
      this.props = attrs;
      Object.assign(this, attrs);
    }
    add(...items: StubElement[]): void {
      this.children.push(...items);
    }
    descendants(): StubElement[] {
      return this.children.flatMap((c) => [c, ...c.descendants()]);
    }
  }
  const make = (type: string) =>
    class extends StubElement {
      constructor(attrs: Record<string, unknown> = {}) {
        super(type, attrs);
      }
    };
  return {
    Group: make("Group"),
    Line: make("Line"),
    Path: make("Path"),
    Rect: make("Rect"),
    Text: make("Text"),
  };
});

describe("paintMath", () => {
  it("renders a dashed fallback box with label when elements are absent", () => {
    const tree = new (class {
      children: any[] = [];
      add(...items: any[]): void {
        this.children.push(...items);
      }
    })();

    const item: LaidOutMathItem = {
      kind: "math",
      inlineIndex: 0,
      xPx: 0,
      widthPx: 60,
      heightPx: 24,
      label: "x = y",
    };

    paintMath(tree as any, item, 10, 20);

    expect(tree.children.length).toBe(2);
    const [rect, text] = tree.children;
    expect(rect.type).toBe("Rect");
    expect(rect.dashPattern).toEqual([3, 2]);
    expect(rect.x).toBe(10);
    expect(rect.y).toBe(20);

    expect(text.type).toBe("Text");
    expect(text.text).toBe("x = y");
    expect(text.italic).toBe(true);
  });

  it("renders structured math elements: fractions, radicals, and slots", () => {
    const tree = new (class {
      children: any[] = [];
      add(...items: any[]): void {
        this.children.push(...items);
      }
      descendants(): any[] {
        return this.children.flatMap((c) => [c, ...(c.descendants ? c.descendants() : [])]);
      }
    })();

    const item: LaidOutMathItem = {
      kind: "math",
      inlineIndex: 0,
      xPx: 0,
      widthPx: 100,
      heightPx: 40,
      label: "\\frac{a}{b}",
      data: {
        widthPx: 100,
        heightPx: 40,
        baselinePx: 30,
        elements: [
          {
            kind: "fraction",
            xPx: 5,
            yPx: 5,
            widthPx: 30,
            heightPx: 30,
            fractionLine: { x1: 0, y1: 15, x2: 30, y2: 15 },
            slots: [
              { xPx: 10, yPx: 0, widthPx: 10, heightPx: 12, text: "a" },
              { xPx: 10, yPx: 18, widthPx: 10, heightPx: 12, text: "b" },
            ],
          },
          {
            kind: "radical",
            xPx: 45,
            yPx: 5,
            widthPx: 40,
            heightPx: 30,
            symbol: "3",
            slots: [{ xPx: 15, yPx: 5, widthPx: 20, heightPx: 20, empty: true }],
          },
        ],
      },
    };

    paintMath(tree as any, item, 10, 20);

    expect(tree.children.length).toBe(1);
    const container = tree.children[0];
    expect(container.type).toBe("Group");
    expect(container.x).toBe(10);
    expect(container.y).toBe(20);

    const descendants = container.descendants();

    // Fraction has a Line for fractionLine
    const line = descendants.find((d: any) => d.type === "Line");
    expect(line).toBeDefined();
    expect(line.points).toEqual([0, 15, 30, 15]);

    // Radical has a Path
    const path = descendants.find((d: any) => d.type === "Path");
    expect(path).toBeDefined();

    // Degree text "3"
    const degText = descendants.find((d: any) => d.type === "Text" && d.text === "3");
    expect(degText).toBeDefined();

    // Slot texts "a" and "b"
    const slotA = descendants.find((d: any) => d.type === "Text" && d.text === "a");
    const slotB = descendants.find((d: any) => d.type === "Text" && d.text === "b");
    expect(slotA).toBeDefined();
    expect(slotB).toBeDefined();

    // Empty slot has a dashed placeholder Rect
    const emptySlot = descendants.find((d: any) => d.type === "Rect" && d.dashPattern);
    expect(emptySlot).toBeDefined();
    expect(emptySlot.dashPattern).toEqual([2, 2]);
  });
});
