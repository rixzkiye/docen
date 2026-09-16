import type { LayoutInline, LayoutParagraph } from "@docen/layout";
import type { DocumentOptions, ParagraphChild, SectionChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { projectDocumentOptions } from "./project";
import type { MarkupDisplay } from "./project/context";

/**
 * Revision display projection: Word's By-author palette (stable per author,
 * distinct between authors, cycling through the fixed slots) vs the fixed
 * By-change-type colors; plus the format-change paint markers.
 */

const REV_DATE = "2026-01-01T00:00:00Z";

const doc = (children: SectionChild[]): DocumentOptions => ({ sections: [{ children }] });

/** Project one paragraph and return its LayoutParagraph. */
function paraOf(children: ParagraphChild[], markup?: MarkupDisplay): LayoutParagraph {
  const blocks = projectDocumentOptions(doc([{ paragraph: { children } }]), markup).sections[0]!
    .blocks;
  const para = blocks[0];
  if (para?.kind !== "paragraph") throw new Error("expected a projected paragraph");
  return para;
}

const insertion = (id: number, author: string, text: string): ParagraphChild => ({
  insertion: { id, author, date: REV_DATE, children: [{ text }] },
});

const deletion = (id: number, author: string, text: string): ParagraphChild => ({
  deletion: { id, author, date: REV_DATE, children: [{ text }] },
});

/** The projected text atoms of a paragraph. */
const atoms = (para: LayoutParagraph): Extract<LayoutInline, { kind: "text" }>[] =>
  para.inline.filter((inline): inline is Extract<LayoutInline, { kind: "text" }> => {
    return inline.kind === "text";
  });

describe("revision author colors", () => {
  it("assigns Word's By-author palette in reviewer order and keeps it stable", () => {
    const para = paraOf([
      insertion(1, "Alice", "a"),
      insertion(2, "Bob", "b"),
      insertion(3, "Alice", "c"),
    ]);
    const [a, b, c] = atoms(para);
    // Slot 0 is the historical red; the second author gets the next slot.
    expect(a?.style).toMatchObject({ underline: true, color: "FF0000" });
    expect(b?.style.color).toBe("2E74B5");
    expect(b?.style.color).not.toBe(a?.style.color);
    // Same author, same color — wherever the revision appears.
    expect(c?.style.color).toBe(a?.style.color);
  });

  it("keeps a deletion in its author's color (struck, not underlined)", () => {
    const para = paraOf([
      insertion(1, "Alice", "a"),
      deletion(2, "Alice", "b"),
      deletion(3, "Bob", "c"),
    ]);
    const [a, b, c] = atoms(para);
    expect(b?.style).toMatchObject({
      strikethrough: true,
      underline: undefined,
      color: a?.style.color,
    });
    expect(c?.style.color).not.toBe(b?.style.color);
  });

  it("reproduces the legacy fixed red in By-change-type mode", () => {
    const para = paraOf([insertion(1, "Alice", "a"), deletion(2, "Bob", "b")], {
      view: "all",
      colors: "changeType",
    });
    const [a, b] = atoms(para);
    expect(a?.style).toMatchObject({ underline: true, color: "FF0000" });
    expect(b?.style).toMatchObject({ strikethrough: true, color: "FF0000" });
  });

  it("assigns palette slots even for filtered-out authors (stable under the filter)", () => {
    const para = paraOf([insertion(1, "Alice", "a"), insertion(2, "Bob", "b")], {
      view: "all",
      authors: ["Bob"],
    });
    const [a, b] = atoms(para);
    // Alice is filtered out → her text renders accepted (no revision look).
    expect(a?.style.color).toBeUndefined();
    expect(a?.style.underline).toBeUndefined();
    // Bob keeps the slot his reviewer order gave him — the filter never
    // recolors the authors that remain visible.
    expect(b?.style.color).toBe("2E74B5");
  });
});

describe("format-change indicators", () => {
  it("marks a w:rPrChange run with its author's color", () => {
    const para = paraOf([
      { text: "changed", revision: { id: 5, author: "Ada", date: REV_DATE } },
      { text: "plain" },
    ]);
    const [changed, plain] = atoms(para);
    expect(changed?.formatChange).toEqual({ color: "FF0000" });
    expect(plain?.formatChange).toBeUndefined();
  });

  it("marks a w:pPrChange paragraph with its author's color", () => {
    expect(paraOf([{ text: "body" }]).formatChange).toBeUndefined();

    const blocks = projectDocumentOptions(
      doc([
        {
          paragraph: {
            alignment: "center",
            revision: { id: 6, author: "Ada", date: REV_DATE, alignment: "left" },
            children: [{ text: "body" }],
          },
        },
      ]),
    ).sections[0]!.blocks;
    const changed = blocks[0];
    if (changed?.kind !== "paragraph") throw new Error("expected a projected paragraph");
    expect(changed.formatChange).toEqual({ color: "FF0000" });
    // The paragraph's CURRENT alignment still projects (accept keeps it).
    expect(changed.align).toBe("center");
  });

  it("uses the neutral gray for format changes in By-change-type mode", () => {
    const blocks = projectDocumentOptions(
      doc([
        {
          paragraph: {
            children: [{ text: "x", revision: { id: 7, author: "Ada", date: REV_DATE } }],
          },
        },
      ]),
      { view: "all", colors: "changeType" },
    ).sections[0]!.blocks;
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected a projected paragraph");
    expect(atoms(para)[0]?.formatChange).toEqual({ color: "808080" });
  });

  it("hides the indicator when the author is filtered out", () => {
    const para = paraOf([{ text: "changed", revision: { id: 8, author: "Ada", date: REV_DATE } }], {
      view: "all",
      authors: ["Bob"],
    });
    expect(atoms(para)[0]?.formatChange).toBeUndefined();
  });
});
