import { describe, expect, it } from "vitest";

import { fakeFontMetrics, installFakeCanvas } from "../../test/fake-canvas";
import type { LayoutParagraph } from "../layout-doc";
import { TextMeasurer } from "../text/measure";
import { layoutBlock } from "./block";

/**
 * The tracked-format-change paint metadata (w:rPrChange / w:pPrChange) is
 * mirrored through the layout for the painter's change bar — it carries no
 * geometry of its own: measuring, wrapping and line heights stay identical.
 */

installFakeCanvas();
const measurer = new TextMeasurer(fakeFontMetrics);

const latin = { family: "serif", sizePx: 16 };

const para = (over: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
  kind: "paragraph",
  inline: [{ kind: "text", text: "word", style: latin }],
  defaultTextStyle: latin,
  ...over,
});

describe("format-change paint metadata", () => {
  it("mirrors a paragraph-level w:pPrChange onto the laid paragraph", () => {
    const out = layoutBlock(para({ formatChange: { color: "FF0000" } }), 500, undefined, measurer);
    expect(out.kind).toBe("paragraph");
    if (out.kind === "paragraph") expect(out.formatChange).toEqual({ color: "FF0000" });
  });

  it("carries a run-level w:rPrChange on the inline atom without changing geometry", () => {
    const plain = para();
    const marked = para({
      inline: [{ kind: "text", text: "word", style: latin, formatChange: { color: "2E74B5" } }],
    });
    const a = layoutBlock(plain, 500, undefined, measurer);
    const b = layoutBlock(marked, 500, undefined, measurer);
    expect(b.heightPx).toBeCloseTo(a.heightPx, 5);
    if (b.kind === "paragraph") {
      expect(b.inline[0]).toMatchObject({ kind: "text", formatChange: { color: "2E74B5" } });
    }
  });
});
