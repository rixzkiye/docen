import { parseParagraphDocx, renderParagraphDocx } from "@docen/docx";
// @vitest-environment happy-dom
import { projectDocumentOptions } from "@docen/docx/layout";
import { layoutParagraph, type LayoutParagraph, TextMeasurer } from "@docen/layout";
import { describe, expect, it } from "vitest";

describe("drop cap projection", () => {
  it("projects dropped drop cap correctly", () => {
    const doc = {
      sections: [
        {
          children: [
            {
              paragraph: {
                dropCap: { val: "drop", lines: 4, distance: 150 },
                children: ["Once upon a time in a faraway kingdom..."],
              },
            },
          ],
        },
      ],
    };
    const projected = projectDocumentOptions(doc as any);
    const para = projected.sections[0].blocks[0] as any;
    expect(para.kind).toBe("paragraph");
    expect(para.dropCap).toBeDefined();
    expect(para.dropCap.type).toBe("dropped");
    expect(para.dropCap.lines).toBe(4);
    expect(para.dropCap.distancePx).toBe(10); // 150 twips = 10 px
  });

  it("projects in-margin drop cap correctly", () => {
    const doc = {
      sections: [
        {
          children: [
            {
              paragraph: {
                dropCap: { val: "margin", lines: 3 },
                children: ["Paragraph with margin drop cap"],
              },
            },
          ],
        },
      ],
    };
    const projected = projectDocumentOptions(doc as any);
    const para = projected.sections[0].blocks[0] as any;
    expect(para.dropCap).toBeDefined();
    expect(para.dropCap.type).toBe("margin");
  });

  it("serializes dropCap attribute to framePr frame options", () => {
    const node = {
      type: "paragraph",
      attrs: {
        dropCap: { val: "drop", lines: 3, distance: 200 },
      },
    };
    const opts = renderParagraphDocx(node);
    expect(opts.frame).toEqual({
      dropCap: "drop",
      lines: 3,
      space: { horizontal: 200, vertical: 0 },
    });

    const parsed = parseParagraphDocx(opts as any);
    expect(parsed.dropCap).toEqual({
      val: "drop",
      lines: 3,
      distance: 200,
      vDistance: 0,
    });
  });

  it("creates selfZones during layoutParagraph when dropCap is defined", () => {
    const fakeCtx = {
      _font: "16px serif",
      set font(v: string) {
        this._font = v;
      },
      get font(): string {
        return this._font;
      },
      measureText(s: string) {
        return { width: s.length * 8 };
      },
    };
    (globalThis as any).OffscreenCanvas = class {
      getContext() {
        return fakeCtx;
      }
    };
    HTMLCanvasElement.prototype.getContext = () => fakeCtx as any;

    const para: LayoutParagraph = {
      kind: "paragraph",
      inline: [
        {
          kind: "text",
          text: "Once upon a time in a faraway kingdom",
          style: { family: "Calibri", sizePx: 16 },
        },
      ],
      dropCap: { type: "dropped", lines: 3, distancePx: 10 },
    };
    const measurer = new TextMeasurer({ normalRatio: () => 1.2 });
    const laid = layoutParagraph(para, 400, undefined, measurer);
    expect(laid.dropCap).toBeDefined();
    expect(laid.dropCap?.type).toBe("dropped");
    // At least one line should have an xOffsetPx due to the selfZone
    expect(laid.lines[0]?.xOffsetPx).toBeGreaterThan(0);
    // The cap glyph is lifted OUT of the flow (rendered once by the painter):
    // the packed lines start at "nce" and the glyph rides dropCapGlyph.
    const lineText = laid.lines
      .flatMap((line) => line.items.filter((item) => item.kind === "text").map((item) => item.text))
      .join("");
    expect(lineText.startsWith("nce upon")).toBe(true);
    expect(lineText).not.toContain("Once");
    expect(laid.dropCapGlyph).toEqual({
      text: "O",
      style: { family: "Calibri", sizePx: 16 },
    });
  });

  it("keeps a combining mark attached to the dropped grapheme", () => {
    (globalThis as any).OffscreenCanvas = class {
      getContext() {
        return {
          font: "16px serif",
          measureText: (s: string) => ({ width: s.length * 8 }),
        };
      }
    };
    HTMLCanvasElement.prototype.getContext = () =>
      ({
        font: "16px serif",
        measureText: (s: string) => ({ width: s.length * 8 }),
      }) as any;
    const para: LayoutParagraph = {
      kind: "paragraph",
      inline: [
        {
          kind: "text",
          text: "e\u0301clair was here",
          style: { family: "Calibri", sizePx: 16 },
        },
      ],
      dropCap: { type: "dropped", lines: 3 },
    };
    const measurer = new TextMeasurer({ normalRatio: () => 1.2 });
    const laid = layoutParagraph(para, 400, undefined, measurer);
    expect(laid.dropCapGlyph?.text).toBe("e\u0301");
    const lineText = laid.lines
      .flatMap((line) => line.items.filter((item) => item.kind === "text").map((item) => item.text))
      .join("");
    expect(lineText.startsWith("clair")).toBe(true);
  });
});
