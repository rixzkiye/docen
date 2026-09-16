// @vitest-environment node
// Hidden-run (w:vanish suppressed) caret/selection geometry. Lives in its own
// file for two reasons:
// 1. The canvas stub must exist BEFORE @docen/layout is first imported — the
//    layout advance module captures its hidden canvas at load. caret-map.spec
//    statically imports @docen/docx (which loads @docen/layout) too early, so
//    its glyph lattice measures zero-width runs (the pre-existing failures).
// 2. The PM document is a hand-rolled fake node tree, so no Tiptap/prosemirror
//    import runs before the stub (prosemirror-view needs a real DOM).

import { describe, expect, it, vi } from "vitest";

// CaretMap grabs a 2d canvas at module load; stub 10px-per-grapheme
// measurement first.
const fakeCtx = {
  set font(_v: string) {},
  measureText: (text: string) => ({ width: text.length * 10 }),
} as unknown as CanvasRenderingContext2D;
vi.stubGlobal("document", {
  createElement: (tag: string) => (tag === "canvas" ? { getContext: () => fakeCtx } : {}),
} as unknown as Document);

const { CaretMap } = await import("./caret-map");

/** The PM textblock shape the zip touches: one text child. */
interface FakeTextChild {
  isText: true;
  textContent: string;
  nodeSize: number;
}

const fakeDoc = (text: string) => {
  const child: FakeTextChild = { isText: true, textContent: text, nodeSize: text.length };
  const para = {
    isTextblock: true,
    textContent: text,
    nodeSize: text.length + 2,
    content: {
      size: text.length,
      forEach(cb: (child: FakeTextChild) => void): void {
        cb(child);
      },
    },
  };
  return {
    descendants(cb: (node: typeof para, pos: number) => boolean | undefined): void {
      cb(para, 0);
    },
  };
};

/** "see " + hidden "secret" + "!" on one line (10px/grapheme). The hidden
 *  atom keeps its characters — they are real PM positions — but paints zero
 *  width when suppressed, so the visible "!" starts at the hidden run's x. */
const hiddenPara = (suppressed: boolean): Record<string, unknown> => {
  const hiddenWidth = suppressed ? 0 : 60;
  return {
    kind: "paragraph",
    heightPx: 20,
    beforePx: 0,
    afterPx: 0,
    inline: [
      { kind: "text", text: "see ", style: { sizePx: 16, family: "Test" } },
      {
        kind: "text",
        text: "secret",
        style: { sizePx: 16, family: "Test", hidden: true },
        ...(suppressed ? { suppressed: true } : {}),
      },
      { kind: "text", text: "!", style: { sizePx: 16, family: "Test" } },
    ],
    lines: [
      {
        yPx: 0,
        heightPx: 20,
        naturalPx: 16,
        maxWidthPx: suppressed ? 50 : 110,
        items: [
          { kind: "text", text: "see ", xPx: 0, widthPx: 40, inlineIndex: 0 },
          { kind: "text", text: "secret", xPx: 40, widthPx: hiddenWidth, inlineIndex: 1 },
          { kind: "text", text: "!", xPx: 40 + hiddenWidth, widthPx: 10, inlineIndex: 2 },
        ],
      },
    ],
  };
};

const mapOf = (suppressed: boolean) => {
  const pages = [{ items: [{ yPx: 0, block: hiddenPara(suppressed) }] }];
  return new CaretMap(pages as never, fakeDoc("see secret!") as never, () => ({
    contentLeftPx: 0,
    contentTopPx: 0,
  }));
};

const xsOf = (map: InstanceType<typeof CaretMap>, positions: number[]): (number | undefined)[] =>
  positions.map((p) => map.caretRect(p)?.xPx);

describe("CaretMap hidden runs (w:vanish suppressed)", () => {
  it("keeps the caret lattice monotonic across a suppressed hidden run", () => {
    const map = mapOf(true);
    expect(map.valid).toBe(true);
    // PM positions 1..12: the visible boundaries plus the paragraph end. The
    // hidden run's positions all collapse onto its x (40) — a natural-width
    // lattice ran past the visible "!" and went non-monotonic.
    expect(xsOf(map, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])).toEqual([
      0, 10, 20, 30, 40, 40, 40, 40, 40, 40, 40, 50,
    ]);
  });

  it("maps clicks to visible boundaries, skipping the suppressed run", () => {
    const map = mapOf(true);
    expect(map.posAtPoint(0, 30, 5)).toBe(4); // end of the visible "see"
    // The collapsed hidden span resolves past the hidden run (Word skips
    // hidden text) — never to a position inside it.
    expect(map.posAtPoint(0, 40, 5)).toBe(11);
    expect(map.posAtPoint(0, 44, 5)).toBe(11); // on the visible "!"
    expect(map.posAtPoint(0, 50, 5)).toBe(12); // past the "!"
  });

  it("highlights the visible extent around a suppressed hidden run", () => {
    const map = mapOf(true);
    const [all] = map.selectionRects(1, 12);
    expect(all).toMatchObject({ xPx: 0, widthPx: 50 });
    // A selection of just the visible "!" starts at the collapsed x, not
    // after a phantom hidden width.
    const [bang] = map.selectionRects(11, 12);
    expect(bang).toMatchObject({ xPx: 40, widthPx: 10 });
  });

  it("treats a shown hidden run as normal text geometry", () => {
    const map = mapOf(false);
    expect(xsOf(map, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])).toEqual([
      0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110,
    ]);
    expect(map.posAtPoint(0, 45, 5)).toBe(5); // inside the shown run
    const [bang] = map.selectionRects(11, 12);
    expect(bang).toMatchObject({ xPx: 100, widthPx: 10 });
  });
});
