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

/** Project one paragraph carrying document-level comment entries. */
function paraWithComments(
  children: ParagraphChild[],
  comments: DocumentOptions["comments"],
  markup: MarkupDisplay,
): LayoutParagraph {
  const blocks = projectDocumentOptions(
    { sections: [{ children: [{ paragraph: { children } }] }], comments },
    markup,
  ).sections[0]!.blocks;
  const para = blocks[0];
  if (para?.kind !== "paragraph") throw new Error("expected a projected paragraph");
  return para;
}

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

  it("keeps one author's color across body and header (document-wide palette)", () => {
    const projected = projectDocumentOptions({
      sections: [
        {
          children: [
            {
              paragraph: { children: [insertion(1, "Alice", "a"), insertion(2, "Bob", "b")] },
            },
          ],
          headers: {
            default: [
              {
                paragraph: {
                  children: [insertion(3, "Bob", "c"), insertion(4, "Alice", "d")],
                },
              },
            ],
          },
        },
      ],
    });
    const bodyPara = projected.sections[0]!.blocks[0];
    if (bodyPara?.kind !== "paragraph") throw new Error("expected a projected paragraph");
    const headerPara = projected.sections[0]!.furniture.header?.[0];
    if (headerPara?.kind !== "paragraph") throw new Error("expected a projected header paragraph");
    const [alice, bob] = atoms(bodyPara);
    const [headerBob, headerAlice] = atoms(headerPara);
    // Body order assigns Alice slot 0, Bob slot 1.
    expect(alice?.style.color).toBe("FF0000");
    expect(bob?.style.color).toBe("2E74B5");
    // The header lists Bob first — the shared palette keeps each author's
    // body color instead of restarting the cycle per furniture walk.
    expect(headerBob?.style.color).toBe(bob?.style.color);
    expect(headerAlice?.style.color).toBe(alice?.style.color);
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

describe("balloon anchors", () => {
  /** A w:comment entry as the compile pass spreads it into DocumentOptions. */
  const comment: NonNullable<DocumentOptions["comments"]>[number] = {
    id: 7,
    author: "Alice",
    initials: "AL",
    date: REV_DATE,
    children: [{ children: [{ text: "note body" }] }],
  };

  it("anchors a comment at its range's first atom with the author card", () => {
    const para = paraWithComments(
      [
        { text: "before " },
        { commentRangeStart: { id: 7 } },
        { text: "marked" },
        { commentRangeEnd: { id: 7 } },
      ],
      [comment],
      { view: "all", balloons: "comments" },
    );
    expect(para.balloons).toEqual([
      { id: 7, kind: "comment", color: "FF0000", label: "AL", text: "note body", inlineIndex: 1 },
    ]);
  });

  it("anchors a run format change at its atom with the author label", () => {
    const para = paraOf(
      [{ text: "plain " }, { text: "styled", revision: { id: 5, author: "Ada", date: REV_DATE } }],
      { view: "all", balloons: "revisions" },
    );
    expect(para.balloons).toEqual([
      { id: 5, kind: "revision", color: "FF0000", label: "Ada", text: "styled", inlineIndex: 1 },
    ]);
  });

  it("anchors a paragraph format change at the paragraph's first line", () => {
    const blocks = projectDocumentOptions(
      {
        sections: [
          {
            children: [
              {
                paragraph: {
                  children: [{ text: "body" }],
                  revision: { id: 6, author: "Ada", date: REV_DATE, alignment: "left" },
                },
              },
            ],
          },
        ],
      },
      { view: "all", balloons: "revisions" },
    ).sections[0]!.blocks;
    const para = blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected a projected paragraph");
    expect(para.balloons).toEqual([
      { id: 6, kind: "revision", color: "FF0000", label: "Ada", inlineIndex: -1 },
    ]);
  });

  it("anchors a deletion with its struck excerpt only in the full view", () => {
    const full = paraOf([deletion(3, "Bob", "gone")], { view: "all", balloons: "revisions" });
    expect(full.balloons).toEqual([
      { id: 3, kind: "revision", color: "FF0000", label: "Bob", text: "gone", inlineIndex: 0 },
    ]);
    // Simple markup keeps deletions inline — B1's accepted rendering (the
    // host remaps simple → all for the canvas).
    const simple = paraOf([deletion(3, "Bob", "gone")], { view: "simple", balloons: "revisions" });
    expect(simple.balloons).toBeUndefined();
    expect(atoms(simple)[0]?.style.strikethrough).toBeUndefined();
  });

  it("scopes the balloon kinds by the Show Markup mode", () => {
    const children: ParagraphChild[] = [
      { commentRangeStart: { id: 7 } },
      { text: "both", revision: { id: 5, author: "Ada", date: REV_DATE } },
      { commentRangeEnd: { id: 7 } },
    ];
    const commentsOnly = paraWithComments(children, [comment], {
      view: "all",
      balloons: "comments",
    });
    expect(commentsOnly.balloons?.map((anchor) => anchor.kind)).toEqual(["comment"]);
    const revisionsOnly = paraWithComments(children, [comment], {
      view: "all",
      balloons: "revisions",
    });
    expect(revisionsOnly.balloons?.map((anchor) => anchor.kind)).toEqual(["revision"]);
    const both = paraWithComments(children, [comment], { view: "all", balloons: "all" });
    expect(both.balloons?.map((anchor) => anchor.kind)).toEqual(["comment", "revision"]);
    const none = paraWithComments(children, [comment], { view: "all", balloons: "none" });
    expect(none.balloons).toBeUndefined();
    // Absent markup keeps every annotation inline (the B1 default).
    const plain = paraWithComments(children, [comment], { view: "all" });
    expect(plain.balloons).toBeUndefined();
  });

  it("drops anchors for authors outside the Specific People filter", () => {
    const para = paraWithComments(
      [{ commentRangeStart: { id: 7 } }, { text: "marked" }, { commentRangeEnd: { id: 7 } }],
      [comment],
      { view: "all", authors: ["Bob"], balloons: "comments" },
    );
    expect(para.balloons).toBeUndefined();
  });

  it("drops every balloon outside the markup views", () => {
    const children: ParagraphChild[] = [
      { text: "x", revision: { id: 5, author: "Ada", date: REV_DATE } },
    ];
    for (const view of ["none", "original"] as const) {
      const para = paraOf(children, { view, balloons: "all" });
      expect(para.balloons).toBeUndefined();
    }
  });

  it("anchors a multi-paragraph comment range once, at its start", () => {
    const blocks = projectDocumentOptions(
      {
        sections: [
          {
            children: [
              {
                paragraph: {
                  children: [{ commentRangeStart: { id: 7 } }, { text: "head" }],
                },
              },
              { paragraph: { children: [{ text: "tail" }] } },
            ],
          },
        ],
        comments: [comment],
      },
      { view: "all", balloons: "comments" },
    ).sections[0]!.blocks;
    const [head, tail] = blocks;
    if (head?.kind !== "paragraph" || tail?.kind !== "paragraph")
      throw new Error("expected projected paragraphs");
    expect(head.balloons?.map((anchor) => anchor.id)).toEqual([7]);
    expect(tail.balloons).toBeUndefined();
  });

  it("keeps balloon anchors stable under the field-code view (Alt+F9)", () => {
    const children: ParagraphChild[] = [
      { text: "plain " },
      { text: "styled", revision: { id: 5, author: "Ada", date: REV_DATE } },
    ];
    const normal = projectDocumentOptions(
      { sections: [{ children: [{ paragraph: { children } }] }] },
      { view: "all", balloons: "revisions" },
      false,
    ).sections[0]!.blocks[0];
    const codes = projectDocumentOptions(
      { sections: [{ children: [{ paragraph: { children } }] }] },
      { view: "all", balloons: "revisions" },
      true,
    ).sections[0]!.blocks[0];
    if (normal?.kind !== "paragraph" || codes?.kind !== "paragraph")
      throw new Error("expected projected paragraphs");
    expect(codes.balloons).toEqual(normal.balloons);
  });
});
