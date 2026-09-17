// itemGlyphLayoutOf — the shared per-grapheme geometry the painter's marks and
// the caret map consume. advance.ts captures its hidden canvas at module load,
// so the document stub must exist BEFORE the dynamic import (same protocol as
// the editor's caret-map spec). The fake measures case-aware advances from the
// font shorthand, so the display transform and the w:w scale are observable.

import { describe, expect, it, vi } from "vitest";

import type { LayoutTextStyle } from "../layout-doc";

const emOf = (font: string): number => {
  const m = /(\d+(?:\.\d+)?)px/.exec(font);
  return m ? Number(m[1]) : 16;
};
const fakeCtx = (() => {
  let font = "16px serif";
  return {
    set font(v: string) {
      font = v;
    },
    get font(): string {
      return font;
    },
    measureText(text: string): { width: number } {
      const em = emOf(font);
      let w = 0;
      for (const ch of text) {
        if (ch === " " || ch === "\t") w += em / 4;
        else if (ch >= "A" && ch <= "Z") w += (em * 5) / 8;
        else if (ch >= "\u4e00" && ch <= "\u9fff") w += em;
        else w += em / 2;
      }
      return { width: w };
    },
  };
})();
vi.stubGlobal("document", {
  createElement: (tag: string) => (tag === "canvas" ? { getContext: () => fakeCtx } : {}),
} as unknown as Document);

const { itemGlyphLayoutOf } = await import("./glyphs");

const latin: LayoutTextStyle = { family: "serif", sizePx: 16 };

describe("itemGlyphLayoutOf", () => {
  it("lays out the natural grapheme positions", () => {
    const plain = itemGlyphLayoutOf({ text: "ab" }, latin);
    expect(plain.xs).toEqual([0, 8]);
    expect(plain.widths).toEqual([8, 8]);
    expect(plain.endX).toBe(16);
  });

  it("measures the displayed (uppercased) glyphs for a caps run", () => {
    const caps = itemGlyphLayoutOf({ text: "ab", displayText: "AB" }, { ...latin, caps: "all" });
    expect(caps.widths).toEqual([10, 10]);
    expect(caps.endX).toBe(20);
  });

  it("uses a smallCaps piece's reduced size", () => {
    const small = itemGlyphLayoutOf(
      { text: "a", displayText: "A", fontSizePx: 12.8 },
      { ...latin, caps: "small" },
    );
    expect(small.widths).toEqual([8]); // uppercase 10 × 0.8
    expect(small.lens).toEqual([1]);
  });

  it("applies the w:w scale to positions and widths", () => {
    const scaled = itemGlyphLayoutOf({ text: "ab" }, { ...latin, scalePct: 200 });
    expect(scaled.xs).toEqual([0, 16]);
    expect(scaled.widths).toEqual([16, 16]);
    expect(scaled.endX).toBe(32);
  });

  it("scales a justify interval back into painted space", () => {
    // Painted target 20px at 200%: the natural layout fills 10px and the
    // returned end lands on the painted interval.
    const stretched = itemGlyphLayoutOf({ text: "a b" }, { ...latin, scalePct: 200 }, 40);
    expect(stretched.endX).toBe(40);
    expect(stretched.xs.at(-1)).toBeLessThan(40);
  });

  it("derives exact cluster positions for shaped OpenType glyph runs", () => {
    const glyphRun = {
      fontSizePx: 16,
      glyphs: [
        {
          glyphId: 10,
          cluster: 0,
          xAdvance: 1000,
          yAdvance: 0,
          xOffset: 0,
          yOffset: 0,
          xPx: 0,
          yPx: 0,
        },
        {
          glyphId: 20,
          cluster: 1,
          xAdvance: 1200,
          yAdvance: 0,
          xOffset: 0,
          yOffset: 0,
          xPx: 10,
          yPx: 0,
        },
      ],
      totalAdvancePx: 22,
    };

    const layout = itemGlyphLayoutOf({ text: "AB", glyphRun }, latin);
    expect(layout.xs).toEqual([0, 10]);
    expect(layout.widths).toEqual([10, 12]);
    expect(layout.endX).toBe(22);
    expect(layout.lens).toEqual([1, 1]);
  });

  it("accurately splits ligature clusters for caret placement inside ligatures", () => {
    // Ligature "fi": 1 glyph (gid 90) covering cluster 0 (length 2 chars)
    const glyphRun = {
      fontSizePx: 16,
      glyphs: [
        {
          glyphId: 90,
          cluster: 0,
          xAdvance: 1200,
          yAdvance: 0,
          xOffset: 0,
          yOffset: 0,
          xPx: 0,
          yPx: 0,
        },
      ],
      totalAdvancePx: 12,
    };

    const layout = itemGlyphLayoutOf({ text: "fi", glyphRun }, latin);
    // 2 graphemes: 'f' and 'i'
    expect(layout.xs.length).toBe(2);
    expect(layout.widths.length).toBe(2);
    expect(layout.xs[0]).toBe(0);
    expect(layout.xs[1]).toBe(6);
    expect(layout.widths[0]).toBe(6);
    expect(layout.widths[1]).toBe(6);
    expect(layout.endX).toBe(12);
  });
});
