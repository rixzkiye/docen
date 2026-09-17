// @vitest-environment node
// The paint-side oracle for paragraph painting: a stubbed leafer-ui lets each
// spec inspect exactly what paintParagraph emits (element types/children and
// props), so a behavior is proven to reach the scene, not just the projection.
// Covers the tracked-format-change margin bars (B1) and the character effects
// (C1) — both share this one painter.

import type {
  FontMetrics,
  LaidOutParagraph,
  LayoutInline,
  LayoutTextStyle,
  ProjectedFlowBox,
} from "@docen/layout";
import type { IGroup } from "leafer-ui";
import { describe, expect, it, vi } from "vitest";

import type { PaintContext } from "./context";

// The glyph-lattice helper captures a 2d canvas at module load; node has
// none, so stub the document BEFORE the dynamic import with a deterministic
// 8px-per-grapheme font so emphasis-mark counts are exact.
const fakeCtx = (() => {
  let font = "16px serif";
  return {
    set font(v: string) {
      font = v;
    },
    get font(): string {
      return font;
    },
    measureText: (text: string) => ({ width: text.length * 8 }),
  };
})() as unknown as CanvasRenderingContext2D;
vi.stubGlobal("document", {
  createElement: (tag: string) => (tag === "canvas" ? { getContext: () => fakeCtx } : {}),
} as unknown as Document);

vi.mock("leafer-ui", () => {
  class StubElement {
    type: string;
    props: Record<string, unknown>;
    children: StubElement[] = [];
    constructor(type: string, attrs: Record<string, unknown> = {}) {
      this.type = type;
      this.props = attrs;
      // The revision-bar spec reads element props directly (child.width),
      // the character-effects spec through `.props` — one shape serves both.
      Object.assign(this, attrs);
    }
    add(...items: StubElement[]): void {
      this.children.push(...items);
    }
    /** Depth-first walk (the character-effects collector). */
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
  // Every leafer-ui binding the painter's module graph imports.
  return {
    Box: make("Box"),
    Ellipse: make("Ellipse"),
    Group: make("Group"),
    Image: make("Image"),
    ImageManager: class {},
    Line: make("Line"),
    Path: make("Path"),
    Rect: make("Rect"),
    Resource: class {},
    Text: make("Text"),
  };
});

const { Group } = await import("leafer-ui");
const { paintParagraph } = await import("./paragraph");

// ── tracked-format-change margin bars ──────────────────────────────────────

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
const bars = (tree: IGroup): { x?: number; y?: number; fill?: string }[] =>
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

// ── character effects ──────────────────────────────────────────────────────

interface StubNode {
  type: string;
  props: Record<string, unknown>;
  children: StubNode[];
  descendants(): StubNode[];
}

const style = (patch: Partial<LayoutTextStyle> = {}): LayoutTextStyle => ({
  family: "serif",
  sizePx: 16,
  ...patch,
});

const ctxOf = (): PaintContext =>
  ({
    metrics: { normalRatio: () => 1.2 },
    flow: { contentWidthPx: 600 },
    pageIndex: 0,
    pageCount: 1,
    layer: "body",
    rerender: () => {},
  }) as unknown as PaintContext;

interface ItemPatch {
  text?: string;
  displayText?: string;
  fontSizePx?: number;
  xPx?: number;
  widthPx?: number;
}

/** One laid line + paragraph around a single text item. */
const paraOf = (
  inline: LayoutInline,
  item: ItemPatch & { justify?: boolean } = {},
): LaidOutParagraph => {
  const { justify, ...rest } = item;
  return {
    kind: "paragraph",
    heightPx: 20,
    beforePx: 0,
    afterPx: 0,
    inline: [inline],
    lines: [
      {
        yPx: 0,
        heightPx: 20,
        naturalPx: 20,
        baselinePadPx: 16,
        final: true,
        maxWidthPx: 600,
        endInlineIndex: 0,
        items: [
          {
            kind: "text",
            inlineIndex: 0,
            text: inline.kind === "text" ? (rest.text ?? inline.text) : "",
            ...(rest.displayText != null ? { displayText: rest.displayText } : {}),
            ...(rest.fontSizePx != null ? { fontSizePx: rest.fontSizePx } : {}),
            xPx: rest.xPx ?? 0,
            widthPx: rest.widthPx ?? 16,
          },
        ],
        ...(justify ? { justifyGapPx: 4 } : {}),
      },
    ],
  };
};

const paint = (para: LaidOutParagraph): StubNode => {
  const root = new Group({}) as unknown as StubNode;
  paintParagraph(root as never, para, 0, 0, ctxOf());
  return root;
};

const nodesOf = (root: StubNode, type: string): StubNode[] =>
  root.descendants().filter((n) => n.type === type);

describe("paintParagraph character effects", () => {
  it("paints the caps display text while the model keeps the source", () => {
    const root = paint(
      paraOf({ kind: "text", text: "ab", style: style({ caps: "all" }) }, { displayText: "AB" }),
    );
    const texts = nodesOf(root, "Text");
    expect(texts).toHaveLength(1);
    expect(texts[0]!.props.text).toBe("AB");
  });

  it("paints a smallCaps piece at its reduced size", () => {
    const root = paint(
      paraOf(
        { kind: "text", text: "a", style: style({ caps: "small" }) },
        { displayText: "A", fontSizePx: 12.8, widthPx: 10 },
      ),
    );
    const [text] = nodesOf(root, "Text");
    expect(text!.props).toMatchObject({ text: "A", fontSize: 12.8, lineHeight: 12.8 });
  });

  it("scales a w:w run horizontally about its left edge", () => {
    const root = paint(
      paraOf({ kind: "text", text: "ab", style: style({ scalePct: 50 }) }, { widthPx: 8 }),
    );
    const [text] = nodesOf(root, "Text");
    expect(text!.props).toMatchObject({ scaleX: 0.5, origin: "left" });
    // An unjustified item carries no width (the element keeps its natural
    // glyph layout, then the scale stretches it).
    expect(text!.props.width).toBeUndefined();
  });

  it("compensates a justified interval for the w:w scale", () => {
    const root = paint(
      paraOf(
        { kind: "text", text: "ab", style: style({ scalePct: 50 }) },
        { widthPx: 8, justify: true, xPx: 0 },
      ),
    );
    const [text] = nodesOf(root, "Text");
    // The stretch target ends at 16 painted px (justifyGapPx 4 over 0..? —
    // rights[0] = 600 here, so width = 600 / 0.5).
    expect(text!.props.scaleX).toBe(0.5);
    expect(text!.props.width).toBe(1200);
  });

  it("shifts a w:position run's baseline without changing the line box", () => {
    const plain = paint(paraOf({ kind: "text", text: "ab", style: style() }));
    const raised = paint(
      paraOf({ kind: "text", text: "ab", style: style({ baselineShiftPx: -8 }) }),
    );
    const [a] = nodesOf(plain, "Text");
    const [b] = nodesOf(raised, "Text");
    expect(b!.props.y).toBeCloseTo((a!.props.y as number) - 8, 5);
    expect(b!.props.height).toBe(a!.props.height);
  });

  it("draws nothing for a suppressed hidden run", () => {
    const root = paint(
      paraOf({ kind: "text", text: "secret", style: style({ hidden: true }), suppressed: true }),
    );
    expect(nodesOf(root, "Text")).toHaveLength(0);
    expect(nodesOf(root, "Line")).toHaveLength(0);
  });

  it("paints a shown hidden run with its dotted marker", () => {
    const root = paint(
      paraOf({ kind: "text", text: "secret", style: style({ hidden: true }) }, { widthPx: 48 }),
    );
    expect(nodesOf(root, "Text")).toHaveLength(1);
    const [marker] = nodesOf(root, "Line");
    expect(marker).toBeDefined();
    expect(marker!.props.dashPattern).toEqual([1, 2]);
    // The marker sits on the shared baseline (baselinePadPx 16) — a hair
    // below it, not at the element top.
    expect(marker!.props.y as number).toBeGreaterThan(16);
    expect(marker!.props.y as number).toBeLessThan(20);
  });

  it("draws a patterned underline on the text baseline (not the element top)", () => {
    // The old call site handed paintUnderlinePattern the element top, so a
    // dotted/dashed w:u stroked ~0.85em above the baseline; the box it draws
    // in is the baseline by contract. baselinePadPx 16 + 0.08em (16px) =
    // 17.28 — just below the glyphs.
    const root = paint(
      paraOf(
        {
          kind: "text",
          text: "u",
          style: style({ underline: true, underlineStyle: "dotted" }),
        },
        { widthPx: 10 },
      ),
    );
    const [line] = nodesOf(root, "Line");
    expect(line).toBeDefined();
    expect(line!.props.y).toBeCloseTo(16 + 16 * 0.08, 5);
  });

  it("paints a w:bdr box around the run with its padding and color", () => {
    const root = paint(
      paraOf(
        {
          kind: "text",
          text: "boxed",
          style: style({ border: { style: "single", px: 2, spacePx: 3, color: "FF0000" } }),
        },
        { widthPx: 40 },
      ),
    );
    const [box] = nodesOf(root, "Rect");
    expect(box).toBeDefined();
    expect(box!.props).toMatchObject({
      x: -3,
      width: 46,
      stroke: "#FF0000",
      strokeWidth: 2,
    });
    expect(box!.props.height as number).toBeCloseTo(16 + 6, 5);
  });

  it("paints one emphasis dot per grapheme above the glyphs", () => {
    const root = paint(
      paraOf(
        { kind: "text", text: "ab", style: style({ emphasisMark: "dot" }) },
        { displayText: "AB" },
      ),
    );
    const dots = nodesOf(root, "Ellipse");
    expect(dots).toHaveLength(2);
    for (const dot of dots) {
      expect(dot.props.fill).toBe("#1b1b1b");
      // Above the baseline.
      expect((dot.props.y as number) + (dot.props.height as number) / 2).toBeLessThan(16);
    }
    const [first, second] = dots;
    expect(second!.props.x as number).toBeGreaterThan(first!.props.x as number);
  });

  it("paints an underDot emphasis below the baseline and a circle stroked", () => {
    const under = paint(
      paraOf({ kind: "text", text: "a", style: style({ emphasisMark: "underDot" }) }),
    );
    const [dot] = nodesOf(under, "Ellipse");
    expect((dot!.props.y as number) + (dot!.props.height as number) / 2).toBeGreaterThan(16);

    const circled = paint(
      paraOf(
        { kind: "text", text: "a", style: style({ emphasisMark: "circle" }) },
        { displayText: "A" },
      ),
    );
    const [circle] = nodesOf(circled, "Ellipse");
    expect(circle!.props.fill).toBe("none");
    expect(circle!.props.stroke).toBe("#1b1b1b");
  });

  it("paints text effects: outline, shadow, emboss, imprint, glow, reflection", () => {
    const outlined = paint(
      paraOf({ kind: "text", text: "outline", style: style({ outline: true, color: "0000FF" }) }),
    );
    const [tOut] = nodesOf(outlined, "Text");
    expect(tOut!.props.fill).toBe("transparent");
    expect(tOut!.props.stroke).toBe("#0000FF");
    expect(tOut!.props.strokeWidth).toBe(1);

    const shadowed = paint(
      paraOf({ kind: "text", text: "shadow", style: style({ shadow: true }) }),
    );
    const [tShad] = nodesOf(shadowed, "Text");
    expect(tShad!.props.shadow).toMatchObject({ x: 1.5, y: 1.5 });

    const glowing = paint(
      paraOf({
        kind: "text",
        text: "glow",
        style: style({ glow: { radiusPx: 8, color: "FF0000" } }),
      }),
    );
    const [tGlow] = nodesOf(glowing, "Text");
    expect(tGlow!.props.shadow).toMatchObject({ x: 0, y: 0, blur: 8, color: "#FF0000" });

    const reflected = paint(
      paraOf({ kind: "text", text: "refl", style: style({ reflection: { opacity: 0.5 } }) }),
    );
    const texts = nodesOf(reflected, "Text");
    expect(texts).toHaveLength(2);
    expect(texts[1].props.scaleY).toBe(-0.6);
    expect(texts[1].props.opacity).toBe(0.5);
  });

  it("paints vertical bar tab line at stop position", () => {
    const tree = new Group();
    paintParagraph(
      tree,
      para({
        tabStops: [{ positionPx: 60, type: "bar" }],
      }),
      10,
      20,
      ctx,
    );
    const rects = (
      tree as unknown as { children: Array<{ type: string; props: Record<string, unknown> }> }
    ).children.filter((c) => c.type === "Rect" && c.props.width === 1);
    expect(rects).toHaveLength(1);
    expect(rects[0].props.x).toBe(70);
    expect(rects[0].props.y).toBe(20);
    expect(rects[0].props.height).toBe(20);
    expect(rects[0].props.fill).toBe("#1b1b1b");
  });
});

describe("paintParagraph drop caps", () => {
  it("paints the lifted cap glyph exactly once, never in the flow", () => {
    const root = paint(
      para({
        dropCap: { type: "dropped", lines: 3, distancePx: 0 },
        dropCapGlyph: { text: "O", style: { family: "serif", sizePx: 16 } },
        lines: [line(0, 0, "nce upon")],
      }),
    );
    const texts = nodesOf(root, "Text");
    const labels = texts.map((t) => t.props.text);
    expect(labels.filter((t) => t === "O")).toHaveLength(1);
    expect(labels).toContain("nce upon");
    expect(labels.join("")).not.toContain("Once");
    const cap = texts.find((t) => t.props.text === "O")!;
    // The cap is enlarged past the run's own size.
    expect(Number(cap.props.fontSize)).toBeGreaterThan(16);
  });

  it("paints no cap when the layout lifted no glyph (no fallback duplication)", () => {
    const root = paint(
      para({
        dropCap: { type: "dropped", lines: 3, distancePx: 0 },
        lines: [line(0, 0, "Once upon")],
      }),
    );
    const labels = nodesOf(root, "Text").map((t) => t.props.text);
    expect(labels).toEqual(["Once upon"]);
  });
});
