import { Schema } from "@tiptap/pm/model";
import { EditorState } from "@tiptap/pm/state";
import { describe, expect, it } from "vitest";

import { MultiSelectionManager } from "./multi-range";

describe("MultiSelectionManager (W2.1 Non-contiguous selection)", () => {
  const schema = new Schema({
    nodes: {
      doc: { content: "paragraph+" },
      paragraph: { content: "text*", toDOM: () => ["p", 0] },
      text: { inline: true },
    },
    marks: {
      bold: {
        toDOM: () => ["strong", 0],
      },
      italic: {
        toDOM: () => ["em", 0],
      },
    },
  });

  const createDoc = (text: string) => {
    return schema.node("doc", null, [schema.node("paragraph", null, [schema.text(text)])]);
  };

  it("adds and normalizes non-contiguous disjoint ranges", () => {
    const mgr = new MultiSelectionManager();
    expect(mgr.hasRanges()).toBe(false);

    mgr.addRange({ from: 10, to: 15 });
    mgr.addRange({ from: 1, to: 5 });

    expect(mgr.hasRanges()).toBe(true);
    expect(mgr.ranges).toEqual([
      { from: 1, to: 5 },
      { from: 10, to: 15 },
    ]);
  });

  it("merges overlapping and adjacent ranges", () => {
    const mgr = new MultiSelectionManager();
    mgr.addRange({ from: 1, to: 5 });
    mgr.addRange({ from: 4, to: 8 });

    expect(mgr.ranges).toEqual([{ from: 1, to: 8 }]);
  });

  it("clears all ranges on clear()", () => {
    const mgr = new MultiSelectionManager();
    mgr.addRange({ from: 1, to: 5 });
    mgr.clear();
    expect(mgr.hasRanges()).toBe(false);
    expect(mgr.ranges).toEqual([]);
  });

  it("deletes multiple ranges in reverse document order", () => {
    const doc = createDoc("Hello World Amazing Editor");
    let state = EditorState.create({ doc, schema });
    const mockView = {
      dispatch: (tr: any) => {
        state = state.apply(tr);
      },
    };

    const mockEditor = {
      state,
      get view() {
        return mockView;
      },
    } as any;

    const mgr = new MultiSelectionManager();
    // In "Hello World Amazing Editor":
    // "World" is [7, 12]
    // "Editor" is [21, 27]
    mgr.addRange({ from: 7, to: 12 });
    mgr.addRange({ from: 21, to: 27 });

    const deleted = mgr.deleteContents(mockEditor);
    expect(deleted).toBe(true);
    expect(mgr.hasRanges()).toBe(false);
    expect(state.doc.textContent).toBe("Hello  Amazing ");
  });

  it("copies concatenated text of all ranges joined by newlines", () => {
    const doc = createDoc("The quick brown fox jumps over the lazy dog");
    const state = EditorState.create({ doc, schema });
    const mockEditor = { state } as any;

    const mgr = new MultiSelectionManager();
    // "quick" is [5, 10]
    // "fox" is [17, 20]
    // "lazy" is [36, 40]
    mgr.addRange({ from: 5, to: 10 });
    mgr.addRange({ from: 17, to: 20 });
    mgr.addRange({ from: 36, to: 40 });

    const text = mgr.copyText(mockEditor);
    expect(text).toBe("quick\nfox\nlazy");
  });

  it("applies formatting marks across all ranges", () => {
    const doc = createDoc("First Second Third");
    let state = EditorState.create({ doc, schema });
    const mockView = {
      dispatch: (tr: any) => {
        state = state.apply(tr);
      },
    };
    const mockEditor = {
      state,
      schema,
      get view() {
        return mockView;
      },
    } as any;

    const mgr = new MultiSelectionManager();
    // "First" is [1, 6]
    // "Third" is [14, 19]
    mgr.addRange({ from: 1, to: 6 });
    mgr.addRange({ from: 14, to: 19 });

    const applied = mgr.applyMark(mockEditor, "bold");
    expect(applied).toBe(true);

    const firstRun = state.doc.resolve(2).nodeAfter;
    expect(firstRun?.marks.some((m) => m.type.name === "bold")).toBe(true);

    const secondRun = state.doc.resolve(8).nodeAfter;
    expect(secondRun?.marks.some((m) => m.type.name === "bold")).toBe(false);

    const thirdRun = state.doc.resolve(15).nodeAfter;
    expect(thirdRun?.marks.some((m) => m.type.name === "bold")).toBe(true);
  });
});
