// Incremental JSON serialization equivalence + identity reuse: the render
// path consumes `pmNodeToJSON` instead of `editor.getJSON()`, so the two must
// be indistinguishable — and untouched subtrees must stay referentially
// stable (the compile/projection caches key on that).
import { docxExtensions } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { createDocJsonCache, pmNodeToJSON } from "./doc-json";

const makeEditor = (content: unknown): Editor =>
  new Editor({ element: null, extensions: docxExtensions, content: content as never });

const baseDoc = {
  type: "doc",
  content: [
    { type: "paragraph", content: [{ type: "text", text: "Hello world" }] },
    {
      type: "paragraph",
      content: [
        { type: "text", text: "bold", marks: [{ type: "bold" }] },
        { type: "text", text: " and " },
        { type: "text", text: "italic", marks: [{ type: "italic" }] },
      ],
    },
    {
      type: "table",
      content: [
        {
          type: "tableRow",
          content: [
            {
              type: "tableCell",
              content: [{ type: "paragraph", content: [{ type: "text", text: "cell" }] }],
            },
            {
              type: "tableCell",
              content: [{ type: "paragraph", content: [{ type: "text", text: "other" }] }],
            },
          ],
        },
      ],
    },
    {
      type: "paragraph",
      content: [{ type: "text", text: "tail" }],
    },
  ],
};

describe("pmNodeToJSON", () => {
  it("matches editor.getJSON() on a structured document", () => {
    const editor = makeEditor(baseDoc);
    const cache = createDocJsonCache();
    expect(pmNodeToJSON(editor.state.doc, cache)).toEqual(editor.getJSON());
  });

  it("matches after a transaction edits one paragraph", () => {
    const editor = makeEditor(baseDoc);
    const cache = createDocJsonCache();
    pmNodeToJSON(editor.state.doc, cache);
    editor.commands.insertContentAt(1, { type: "text", text: "XY" });
    expect(pmNodeToJSON(editor.state.doc, cache)).toEqual(editor.getJSON());
    editor.commands.deleteRange({ from: 1, to: 3 });
    expect(pmNodeToJSON(editor.state.doc, cache)).toEqual(editor.getJSON());
  });

  it("keeps untouched top-level blocks referentially stable", () => {
    const editor = makeEditor(baseDoc);
    const cache = createDocJsonCache();
    const before = pmNodeToJSON(editor.state.doc, cache);
    // Insert inside the first paragraph only.
    editor.commands.insertContentAt(6, { type: "text", text: "!" });
    const after = pmNodeToJSON(editor.state.doc, cache);
    expect(after).toEqual(editor.getJSON());
    expect(after).not.toBe(before);
    // The edited paragraph is a fresh object, every other block is the same.
    expect(after.content![0]).not.toBe(before.content![0]);
    expect(after.content![1]).toBe(before.content![1]);
    expect(after.content![2]).toBe(before.content![2]);
    expect(after.content![3]).toBe(before.content![3]);
    // Selection-only transactions keep the whole doc identical.
    editor.commands.setTextSelection(2);
    const same = pmNodeToJSON(editor.state.doc, cache);
    expect(same).toBe(after);
  });

  it("survives a split/join with fresh objects only along the edit", () => {
    const editor = makeEditor(baseDoc);
    const cache = createDocJsonCache();
    const before = pmNodeToJSON(editor.state.doc, cache);
    editor.commands.splitBlock();
    const split = pmNodeToJSON(editor.state.doc, cache);
    expect(split).toEqual(editor.getJSON());
    editor.commands.joinBackward();
    const joined = pmNodeToJSON(editor.state.doc, cache);
    expect(joined).toEqual(editor.getJSON());
    // The table and the tail survive the split/join untouched.
    const tableBefore = before.content![2];
    expect(split.content![3]).toBe(tableBefore);
    expect(joined.content![2]).toBe(tableBefore);
  });
});
