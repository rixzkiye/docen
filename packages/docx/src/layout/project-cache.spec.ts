// Incremental projection equivalence: the compile + projection caches must
// never change what a render produces. Each step applies one transaction-like
// edit to a document whose untouched subtrees stay referentially stable (the
// editor's `pmNodeToJSON` builds exactly this shape), then compares the
// cached pipeline's output against a from-scratch compile + projection.
import type { JSONContent } from "@docen/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, createCompileCache } from "../converters/docx";
import { createProjectionCache, projectDocumentOptions, type ProjectedSection } from "./project";

/** One transaction-shaped edit: a copy-on-write JSON tree (only the edited
 *  path re-allocates, like ProseMirror's persistent nodes). */
interface Step {
  name: string;
  apply(doc: JSONContent): JSONContent;
}

const text = (t: string, marks?: JSONContent["marks"]): JSONContent => ({
  type: "text",
  text: t,
  ...(marks ? { marks } : {}),
});

const para = (...content: JSONContent[]): JSONContent => ({
  type: "paragraph",
  content: content.length > 0 ? content : undefined,
});

const passthrough = (data: unknown): JSONContent => ({
  type: "inlinePassthrough",
  attrs: { data: JSON.stringify(data) },
});

/** Rebuild the doc with `content[i]` replaced by `make(...)` — everything
 *  outside the path keeps its identity. */
function replaceBlock(
  doc: JSONContent,
  index: number,
  make: (block: JSONContent) => JSONContent,
): JSONContent {
  const content = (doc.content ?? []).slice();
  content[index] = make(content[index]!);
  return { ...doc, content };
}

/** The first text run's text replaced inside top-level block `index`. */
function editBlockText(index: number, next: string): Step {
  return {
    name: `text:${index}`,
    apply: (doc) =>
      replaceBlock(doc, index, (block) => ({
        ...block,
        content: (block.content ?? []).map((child, i) =>
          i === 0 && child.type === "text" ? { ...child, text: next } : child,
        ),
      })),
  };
}

/** A paragraph appended at the end of the body. */
function appendParagraph(message: string): Step {
  return {
    name: `append:${message}`,
    apply: (doc) => ({ ...doc, content: [...(doc.content ?? []), para(text(message))] }),
  };
}

function removeBlock(index: number): Step {
  return {
    name: `remove:${index}`,
    apply: (doc) => ({
      ...doc,
      content: (doc.content ?? []).filter((_, i) => i !== index),
    }),
  };
}

function replaceTable(): Step {
  return {
    name: "table",
    apply: (doc) => replaceBlock(doc, 6, (block) => JSON.parse(JSON.stringify(block))),
  };
}

function corpus(): JSONContent {
  return {
    type: "doc",
    attrs: {
      sectionProperties: {
        pageSize: { width: 11906, height: 16838 },
        pageMargin: { top: 1440, left: 1800, right: 1800, bottom: 1440 },
      },
      documentExtras: {
        comments: [
          {
            id: 1,
            author: "Ada",
            initials: "A",
            children: [{ paragraph: { children: [{ text: "a comment" }] } }],
          },
        ],
        footnotes: [
          {
            id: 1,
            children: [{ paragraph: { children: [{ text: "a note", style: "FootnoteText" }] } }],
          },
        ],
      },
    },
    content: [
      { type: "paragraph", attrs: { heading: "Title" }, content: [text("Projected cache")] },
      para(
        text("plain "),
        text("bold", [{ type: "bold" }]),
        text(" then "),
        text("italic", [{ type: "italic" }]),
      ),
      {
        type: "paragraph",
        attrs: { numbering: { reference: "list_1", level: 0 } },
        content: [text("first")],
      },
      {
        type: "paragraph",
        attrs: { numbering: { reference: "list_1", level: 0 } },
        content: [text("second")],
      },
      {
        type: "paragraph",
        attrs: { numbering: { reference: "list_1", level: 1 } },
        content: [text("nested")],
      },
      {
        type: "paragraph",
        attrs: { bullet: { level: 0 } },
        content: [text("bullet")],
      },
      {
        type: "table",
        content: [
          {
            type: "tableRow",
            content: [
              {
                type: "tableCell",
                content: [para(text("left"))],
              },
              {
                type: "tableCell",
                content: [para(text("right"))],
              },
            ],
          },
          {
            type: "tableRow",
            content: [
              {
                type: "tableCell",
                content: [para(text("bottom"))],
              },
              {
                type: "tableCell",
                content: [para(text("cell"))],
              },
            ],
          },
        ],
      },
      para(
        text("commented", [{ type: "insertion", attrs: { author: "Ada", id: "r1" } }]),
        passthrough({ commentRangeStart: { id: 1 } }),
        text("range"),
        passthrough({ commentRangeEnd: { id: 1 } }),
        passthrough({ commentReference: { id: 1 } }),
      ),
      para(
        text("note"),
        passthrough({ footnoteReference: 1 }),
        passthrough({ simpleField: { instruction: "PAGE" } }),
      ),
      {
        type: "paragraph",
        attrs: {
          sectionProperties: {
            type: "continuous",
            pageSize: { width: 11906, height: 16838 },
            pageMargin: { top: 1440, left: 1440, right: 1440, bottom: 1440 },
            columns: { count: 2, space: 720 },
          },
        },
        content: [text("section break")],
      },
      para(text("two-column body")),
    ],
  };
}

/** A cached projection must deep-equal a from-scratch one — the cache is an
 *  execution detail, never a behavior. */
function expectEquivalent(doc: JSONContent): { sections: ProjectedSection[] } {
  const compileCache = createCompileCache();
  const projectionCache = createProjectionCache();
  const cached = projectDocumentOptions(
    compileDocument(doc, undefined, compileCache),
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    projectionCache,
  );
  const fresh = projectDocumentOptions(compileDocument(doc));
  // Compare structurally, ignoring object identity (toEqual does both).
  expect(cached.sections).toEqual(fresh.sections);
  expect(cached.background).toEqual(fresh.background);
  return cached;
}

const steps: Step[] = [
  editBlockText(0, "Projected cache!"),
  editBlockText(1, "plain changed bold then italic"),
  appendParagraph("appended"),
  {
    name: "insert-list-item",
    apply: (doc) => ({
      ...doc,
      content: [
        ...(doc.content ?? []).slice(0, 3),
        para(text("inserted")),
        ...(doc.content ?? []).slice(3),
      ],
    }),
  },
  removeBlock(4),
  replaceTable(),
  editBlockText(7, "commented change"),
  removeBlock(2),
  appendParagraph("tail"),
  editBlockText(1, "done"),
];

describe("incremental projection (compile + projection cache)", () => {
  it("matches a fresh projection across a battery of transaction shapes", () => {
    let doc = corpus();
    for (const step of steps) {
      doc = step.apply(doc);
      expectEquivalent(doc);
    }
  });

  it("is deterministic when replayed from the same start", () => {
    let doc = corpus();
    const first = steps.map((step) => {
      doc = step.apply(doc);
      return expectEquivalent(doc);
    });
    let replay = corpus();
    const second = steps.map((step) => {
      replay = step.apply(replay);
      return expectEquivalent(replay);
    });
    // `toEqual` ignores identity but catches any state leak between runs.
    expect(second).toEqual(first);
  });

  it("reuses the previous projected blocks for untouched top-level blocks", () => {
    let doc = corpus();
    const compileCache = createCompileCache();
    const projectionCache = createProjectionCache();
    const render = (d: JSONContent) =>
      projectDocumentOptions(
        compileDocument(d, undefined, compileCache),
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        projectionCache,
      );
    const before = render(doc);
    doc = editBlockText(1, "tiny edit").apply(doc);
    const after = render(doc);
    expect(after.sections[0]!.blocks[0]).toBe(before.sections[0]!.blocks[0]);
    expect(after.sections[0]!.blocks[2]).toBe(before.sections[0]!.blocks[2]);
    // The edited paragraph re-projected.
    expect(after.sections[0]!.blocks[1]).not.toBe(before.sections[0]!.blocks[1]);
    // A section with no changed child keeps its whole projected object
    // (blocks array and flow box included).
    expect(after.sections[1]).toBe(before.sections[1]);
    expect(after.sections[1]!.flow).toBe(before.sections[1]!.flow);
  });

  it("re-projects the paragraph a newly inserted block bookmark anchors to", () => {
    // The paragraph object keeps its identity across both projections; only
    // the block-level bookmark child is new. The cached non-stateful paragraph
    // must not be reused while the bookmark buffer is non-empty, or the marker
    // would leak past its paragraph.
    const paragraph = { paragraph: { children: ["after"] } };
    const cache = createProjectionCache();
    const before = projectDocumentOptions(
      { sections: [{ children: [paragraph] }] },
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      cache,
    );
    expect(before.sections[0]!.blocks[0]).toMatchObject({ kind: "paragraph" });

    const after = projectDocumentOptions(
      { sections: [{ children: [{ bookmarkStart: { id: 7, name: "Bm" } }, paragraph] }] },
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      cache,
    );
    const block = after.sections[0]!.blocks[0];
    if (block?.kind !== "paragraph") throw new Error("expected paragraph");
    expect(block.bookmarks).toEqual([{ name: "Bm", inlineIndex: 0 }]);
  });
});
