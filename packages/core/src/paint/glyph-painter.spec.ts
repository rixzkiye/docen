import type { LaidOutGlyphRun } from "@docen/layout";
import type { PathCommand } from "@docen/shaping";
import { describe, expect, it, vi } from "vitest";

vi.mock("leafer-ui", () => {
  class StubElement {
    children: StubElement[] = [];
    constructor(public attrs: Record<string, unknown> = {}) {
      Object.assign(this, attrs);
    }
    add(...items: StubElement[]): void {
      this.children.push(...items);
    }
  }
  const make = () => class extends StubElement {};
  return {
    Group: make(),
    Path: make(),
    Line: make(),
  };
});

const { Group } = await import("leafer-ui");
const { GlyphOutlineCache, commandsToSvgPath, paintGlyphRun } = await import("./glyph-painter");

describe("R6.4 Glyph Painter (Vector Outline Painting)", () => {
  it("converts PathCommand array to standard SVG path data with Y inversion", () => {
    const commands: PathCommand[] = [
      { type: "M", x: 10, y: 20 },
      { type: "L", x: 30, y: 40 },
      { type: "Q", cx: 50, cy: 60, x: 70, y: 80 },
      { type: "C", cx1: 90, cy1: 100, cx2: 110, cy2: 120, x: 130, y: 140 },
      { type: "Z" },
    ];

    const svg = commandsToSvgPath(commands);
    expect(svg).toBe("M 10 -20 L 30 -40 Q 50 -60 70 -80 C 90 -100 110 -120 130 -140 Z");
  });

  it("handles empty command array gracefully", () => {
    expect(commandsToSvgPath([])).toBe("");
  });

  it("caches and retrieves glyph outlines in GlyphOutlineCache", () => {
    const cache = new GlyphOutlineCache(3);
    cache.setGlyphPath(1, 10, "M 0 0 L 10 10 Z");
    cache.setGlyphPath(1, 20, "M 5 5 L 15 15 Z");

    expect(cache.getGlyphPath(1, 10)).toBe("M 0 0 L 10 10 Z");
    expect(cache.getGlyphPath(1, 20)).toBe("M 5 5 L 15 15 Z");

    // Add more entries to trigger LRU eviction
    cache.setGlyphPath(1, 30, "path3");
    cache.setGlyphPath(1, 40, "path4");

    // Key 1:10 was least recently used, should be evicted
    expect(cache.getGlyphPath(1, 10)).toBeUndefined();
    expect(cache.getGlyphPath(1, 40)).toBe("path4");
  });

  it("paints glyph runs into tree using cached vector paths", () => {
    const tree = new Group();
    const cache = new GlyphOutlineCache();
    cache.setGlyphPath(1, 42, "M 0 0 L 10 0 L 10 10 Z");

    const glyphRun: LaidOutGlyphRun = {
      fontId: 1,
      fontName: "TestFont",
      fontSizePx: 20,
      totalAdvancePx: 25,
      glyphs: [
        {
          glyphId: 42,
          cluster: 0,
          xAdvance: 1000,
          yAdvance: 0,
          xOffset: 0,
          yOffset: 0,
          xPx: 0,
          yPx: 0,
        },
      ],
    };

    const painted = paintGlyphRun(tree as any, glyphRun, {
      x: 50,
      y: 100,
      fill: "#ff0000",
      outlineCache: cache,
      unitsPerEm: 1000,
    });

    expect(painted).toBe(true);
    expect(tree.children.length).toBe(1);

    const child = tree.children[0] as any;
    expect(child.path).toBe("M 0 0 L 10 0 L 10 10 Z");
    expect(child.x).toBe(50);
    expect(child.y).toBe(100);
    expect(child.fill).toBe("#ff0000");
    expect(child.scaleX).toBeCloseTo(0.02); // 20 / 1000
    expect(child.scaleY).toBeCloseTo(0.02);
  });
});
