// Character effects (w:caps / w:smallCaps / w:w / w:vanish / w:position):
// the projection carries the resolved attributes; this suite proves the
// measurement layer turns them into the right advances, display text and
// per-piece sizes. The canvas fake is case-sensitive (lowercase 0.5em,
// uppercase 0.625em, space 0.25em, CJK 1em) so an uppercase display form is
// observable in the measured width alone.

import { describe, expect, it } from "vitest";

import { fakeFontMetrics } from "../../test/fake-canvas";
import type { LayoutInline, LayoutTextStyle } from "../layout-doc";
import type { LaidOutLineItem } from "../layout-result";
import { packLines } from "./line-break";
import { capsPiecesOf, characterScaleOf, displayTextOf, kerningActive } from "./measure";
import { TextMeasurer } from "./measure";

// The spec's own deterministic canvas: advance fractions are derived from the
// font shorthand's px size (lowercase .5em, uppercase .625em, space .25em,
// CJK 1em).
const emOf = (font: string): number => {
  const m = /(\d+(?:\.\d+)?)px/.exec(font);
  return m ? Number(m[1]) : 16;
};
const advanceOf = (ch: string, em: number): number => {
  if (ch === " " || ch === "\t") return em / 4;
  if (ch >= "\u4e00" && ch <= "\u9fff") return em;
  if (ch >= "A" && ch <= "Z") return (em * 5) / 8;
  return em / 2;
};

if (typeof OffscreenCanvas === "undefined") {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- the minimal duck-typed surface pretext touches
  (globalThis as any).OffscreenCanvas = class {
    getContext(): unknown {
      let font = "16px serif";
      return {
        set font(v: string) {
          font = v;
        },
        get font(): string {
          return font;
        },
        measureText(s: string): { width: number } {
          const em = emOf(font);
          let w = 0;
          for (const ch of s) w += advanceOf(ch, em);
          return { width: w };
        },
      };
    }
  };
}

const measurer = new TextMeasurer(fakeFontMetrics);

// 16px em: lowercase 8, uppercase 10, space 4.
const latin: LayoutTextStyle = { family: "serif", sizePx: 16 };

const text = (t: string, style: LayoutTextStyle = latin): LayoutInline => ({
  kind: "text",
  text: t,
  style,
});

const pack = (inline: LayoutInline[], width = 1000) =>
  packLines(inline, { measurer, width, lineHeight: ({ naturalPx }) => naturalPx });

const textItems = (line: {
  items: LaidOutLineItem[];
}): Extract<LaidOutLineItem, { kind: "text" }>[] =>
  line.items.filter((i): i is Extract<LaidOutLineItem, { kind: "text" }> => i.kind === "text");

describe("capsPiecesOf", () => {
  it("returns one untouched piece without caps formatting", () => {
    expect(capsPiecesOf("aA1 ", undefined)).toEqual([
      { source: "aA1 ", display: "aA1 ", small: false },
    ]);
  });

  it("uppercases the whole run for allCaps", () => {
    expect(capsPiecesOf("aA1ß", "all")).toEqual([
      // ß → SS changes the UTF-16 length, so the source glyph is kept — the
      // caret lattice never shifts.
      { source: "aA1ß", display: "AA1ß", small: false },
    ]);
  });

  it("splits smallCaps at case boundaries, marking lowercase pieces small", () => {
    expect(capsPiecesOf("aB1cD", "small")).toEqual([
      { source: "a", display: "A", small: true },
      { source: "B1", display: "B1", small: false },
      { source: "c", display: "C", small: true },
      { source: "D", display: "D", small: false },
    ]);
  });
});

describe("displayTextOf / characterScaleOf / kerningActive", () => {
  it("keeps the UTF-16 length when the case mapping would change it", () => {
    const text = "straße";
    expect(displayTextOf(text, "all").length).toBe(text.length);
    expect(displayTextOf(text, "all")).toBe("STRAßE");
  });

  it("resolves the w:w percent to a scale (100/absent natural)", () => {
    expect(characterScaleOf({ family: "serif", sizePx: 16 })).toBe(1);
    expect(characterScaleOf({ family: "serif", sizePx: 16, scalePct: 100 })).toBe(1);
    expect(characterScaleOf({ family: "serif", sizePx: 16, scalePct: 50 })).toBe(0.5);
    expect(characterScaleOf({ family: "serif", sizePx: 16, scalePct: 200 })).toBe(2);
  });

  it("activates kerning when the run's size reaches the w:kern threshold", () => {
    // w:kern 16 (half-points) = 8pt; a 12pt run kerns, a 6pt run does not.
    expect(kerningActive({ family: "serif", sizePx: 16, kernPt: 8 })).toBe(true);
    expect(kerningActive({ family: "serif", sizePx: 8, kernPt: 8 })).toBe(false);
    expect(kerningActive({ family: "serif", sizePx: 16 })).toBe(false);
  });
});

describe("packLines character effects", () => {
  it("measures allCaps with the uppercase advances but keeps the source text", () => {
    const lines = pack([text("ab", { ...latin, caps: "all" })]);
    const [item] = textItems(lines[0]!);
    expect(item!.text).toBe("ab");
    expect(item!.displayText).toBe("AB");
    // Uppercase advances: 10 + 10, not the lowercase 8 + 8.
    expect(item!.widthPx).toBe(20);
  });

  it("renders smallCaps lowercase at 80% and original capitals at full size", () => {
    const lines = pack([text("aB", { ...latin, caps: "small" })]);
    const items = textItems(lines[0]!);
    expect(items).toHaveLength(2);
    // "a" → "A" at 12.8px: uppercase advance 10 × 0.8 = 8.
    expect(items[0]).toMatchObject({ text: "a", displayText: "A", widthPx: 8, fontSizePx: 12.8 });
    // The run's own capital keeps the full size (display === source, so no
    // displayText/fontSize override is written).
    expect(items[1]).toMatchObject({ text: "B", widthPx: 10 });
    expect(items[1]!.displayText).toBeUndefined();
    expect(items[1]!.fontSizePx).toBeUndefined();
  });

  it("scales the measured advance by w:w without touching the source text", () => {
    const lines = pack([text("ab", { ...latin, scalePct: 50 })]);
    const [item] = textItems(lines[0]!);
    expect(item!.text).toBe("ab");
    expect(item!.displayText).toBeUndefined();
    expect(item!.widthPx).toBe(8); // 16 × 50%
  });

  it("combines w:caps with w:w (uppercase advances, then the scale)", () => {
    const lines = pack([text("ab", { ...latin, caps: "all", scalePct: 200 })]);
    const [item] = textItems(lines[0]!);
    expect(item!.displayText).toBe("AB");
    expect(item!.widthPx).toBe(40); // (10 + 10) × 200%
  });

  it("zeroes a suppressed hidden run's advance but keeps its characters on the line", () => {
    const lines = pack([
      text("see "),
      { kind: "text", text: "secret", style: latin, suppressed: true },
    ]);
    const items = textItems(lines[0]!);
    expect(items).toHaveLength(2);
    expect(items[1]).toMatchObject({ text: "secret", widthPx: 0 });
    // The visible run keeps its own advance and the hidden run adds nothing.
    expect(items[0]!.widthPx).toBe(28); // "see " = 8+8+8+4
    // A run that is only hidden paints no line metric of its own: the line
    // floors at the paragraph strut instead of the hidden glyph size.
    const hiddenOnly = pack([
      { kind: "text", text: "xx", style: { ...latin, sizePx: 40 }, suppressed: true },
    ]);
    expect(hiddenOnly[0]!.naturalPx).toBe(0);
    expect(hiddenOnly[0]!.textEmPx).toBeUndefined();
  });

  it("leaves w:position out of the layout (glyphs shift, the line box does not)", () => {
    const plain = pack([text("ab")]);
    const raised = pack([text("ab", { ...latin, baselineShiftPx: -8 })]);
    expect(raised[0]!.heightPx).toBe(plain[0]!.heightPx);
    expect(raised[0]!.naturalPx).toBe(plain[0]!.naturalPx);
    expect(textItems(raised[0]!)[0]!.xPx).toBe(0);
  });
});
