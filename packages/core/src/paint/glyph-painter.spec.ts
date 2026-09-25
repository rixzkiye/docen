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

  it("scales outlines by the run's unitsPerEm (2048-em fonts paint half as large)", () => {
    const tree = new Group();
    const cache = new GlyphOutlineCache();
    cache.setGlyphPath(1, 42, "M 0 0 L 10 0 L 10 10 Z");
    const glyphRun: LaidOutGlyphRun = {
      fontId: 1,
      fontSizePx: 20,
      unitsPerEm: 2048,
      totalAdvancePx: 25,
      glyphs: [
        {
          glyphId: 42,
          cluster: 0,
          xAdvance: 1000,
          yAdvance: 0,
          xOffset: 0,
          yOffset: 0,
          xPx: 10,
          yPx: 0,
        },
      ],
    };

    paintGlyphRun(tree as any, glyphRun, { x: 0, y: 0, outlineCache: cache });
    const child = tree.children[0] as any;
    expect(child.scaleX).toBeCloseTo(20 / 2048, 6);
    expect(child.scaleY).toBeCloseTo(20 / 2048, 6);
    // advanceScale stretches positions and the x scale together
    tree.children.length = 0;
    paintGlyphRun(tree as any, glyphRun, { x: 0, y: 0, outlineCache: cache, advanceScale: 2 });
    const stretched = tree.children[0] as any;
    expect(stretched.x).toBe(20);
    expect(stretched.scaleX).toBeCloseTo((20 / 2048) * 2, 6);
    expect(stretched.scaleY).toBeCloseTo(20 / 2048, 6);
  });

  it("keys cached outlines by variation coordinates", () => {
    const cache = new GlyphOutlineCache();
    cache.setGlyphPath(1, 10, "plain");
    expect(cache.getGlyphPath(1, 10, [{ tag: "wght", value: 700 }])).toBeUndefined();
  });

  it("wordStretch grows only the space advances — word outlines stay natural", () => {
    const tree = new Group();
    const cache = new GlyphOutlineCache();
    for (const gid of [1, 2, 3, 4, 5]) cache.setGlyphPath(1, gid, `M 0 0 L ${gid} 0 Z`);
    // An empty outline for the space glyph: it must still advance the gap.
    cache.setGlyphPath(1, 4, "");

    const glyphRun: LaidOutGlyphRun = {
      fontId: 1,
      fontName: "TestFont",
      fontSizePx: 10,
      unitsPerEm: 1000,
      totalAdvancePx: 50,
      direction: "ltr",
      glyphs: [0, 1, 2, 3, 4].map((i) => ({
        glyphId: i + 1,
        cluster: i,
        xAdvance: 1000,
        yAdvance: 0,
        xOffset: 0,
        yOffset: 0,
        xPx: i * 10,
        yPx: 0,
      })),
    };

    const painted = paintGlyphRun(tree as any, glyphRun, {
      x: 0,
      y: 0,
      outlineCache: cache,
      wordStretch: { text: "dan c ", gapPx: 6 },
    });

    expect(painted).toBe(true);
    // The space paints no ink; the four word glyphs do.
    expect(tree.children.length).toBe(4);
    const xs = tree.children.map((c: any) => c.x);
    // "d", "a", "n" keep their natural pen positions; "c" (after the space)
    // shifts by exactly the gap.
    expect(xs).toEqual([0, 10, 20, 46]);
    // The outlines are never x-scaled by a word gap.
    for (const child of tree.children as any[]) expect(child.scaleX).toBeCloseTo(0.01, 6);

    // Trailing whitespace past `stretchEnd` hangs unstretched.
    tree.children.length = 0;
    paintGlyphRun(tree as any, glyphRun, {
      x: 0,
      y: 0,
      outlineCache: cache,
      wordStretch: { text: "dan c ", gapPx: 6, stretchEnd: 3 },
    });
    expect((tree.children as any[]).map((c: any) => c.x)).toEqual([0, 10, 20, 40]);
  });
});
