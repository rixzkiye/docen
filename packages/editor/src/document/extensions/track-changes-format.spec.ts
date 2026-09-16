import {
  Bold,
  Deletion,
  Document,
  FormatChange,
  Insertion,
  Italic,
  Paragraph,
  TextStyle,
  compileDocument,
  docxExtensions,
  parseFormatRecords,
  type RunFormatRecord,
} from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import type { Mark } from "@tiptap/pm/model";
import { afterEach, describe, expect, it } from "vitest";

import { updateSettings } from "../settings";
import { revisionAuthorOf, TrackChanges, collectRevisions } from "./track-changes";

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

/** A headless editor with the schema the format-change workflow touches:
 *  paragraph text + the revision marks + rPr marks to toggle + the extension. */
const build = (content?: Record<string, unknown>[]): EditorType => {
  // The settings store is module-global and this workspace runs specs
  // concurrently — every editor starts from the historical default identity so
  // one test's identity can never leak into another test's author assertions.
  updateSettings({ identity: { name: "" } });
  const editor = new Editor({
    element: null,
    extensions: [
      Document,
      Paragraph,
      Text,
      TextStyle,
      Bold,
      Italic,
      UndoRedo,
      Insertion,
      Deletion,
      FormatChange,
      TrackChanges,
    ],
    content: {
      type: "doc",
      content: [{ type: "paragraph", content: content ?? [{ type: "text", text: "hello world" }] }],
    },
  });
  // element:null skips Tiptap's mount (and with it plugin installation) — the
  // same gap the canvas edit bridge patches by registering the sorted list.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

const addMark = (editor: EditorType, name: string, from: number, to: number): void => {
  editor.commands.command(({ state, dispatch }) => {
    dispatch?.(state.tr.addMark(from, to, state.schema.marks[name]!.create()));
    return true;
  });
};

const removeMark = (editor: EditorType, name: string, from: number, to: number): void => {
  editor.commands.command(({ state, dispatch }) => {
    dispatch?.(state.tr.removeMark(from, to, state.schema.marks[name]!));
    return true;
  });
};

/** The formatChange marks in the doc, one entry per text node carrying one. */
const formatMarks = (editor: EditorType) => {
  const out: { attrs: Record<string, unknown>; from: number; to: number }[] = [];
  editor.state.doc.descendants((node, pos) => {
    if (!node.isText) return true;
    for (const mark of node.marks) {
      if (mark.type.name !== "formatChange") continue;
      out.push({
        attrs: mark.attrs as Record<string, unknown>,
        from: pos,
        to: pos + node.nodeSize,
      });
    }
    return true;
  });
  return out;
};

/** Every format record in the doc, flattened in document order. */
const formatRecords = (editor: EditorType): RunFormatRecord[] => {
  const out: RunFormatRecord[] = [];
  for (const mark of formatMarks(editor)) {
    out.push(...parseFormatRecords(mark.attrs.records));
  }
  return out;
};

/** Mark names on the text node at `pos` (default: the leading run). */
const markNames = (editor: EditorType, pos = 1): string[] => {
  const text = editor.state.doc.nodeAt(pos);
  return (text?.marks ?? []).map((mark: Mark) => mark.type.name);
};

/** The compiled run options of the first paragraph — the DOCX-facing proof. */
const compiledRuns = (editor: EditorType): Record<string, unknown>[] => {
  const out = compileDocument(editor.getJSON() as never, docxExtensions);
  const section = out.sections[out.sections.length - 1]!;
  const para = section.children[0] as { paragraph: { children?: Record<string, unknown>[] } };
  return para.paragraph.children ?? [];
};

const paragraphAttrs = (editor: EditorType): Record<string, unknown> =>
  editor.state.doc.firstChild!.attrs as Record<string, unknown>;

afterEach(() => {
  // The settings store is module-global (in-memory in the test env) — clear
  // the identity so the fallback tests stay independent of run order.
  updateSettings({ identity: { name: "" } });
});

describe.sequential("revision author identity", () => {
  it("falls back to Word's historical default when no identity is set", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    expect(formatRecords(editor)[0]!.author).toBe("docen");
    editor.destroy();
  });

  it("uses the settings user name for new revisions (text and format)", () => {
    const editor = build();
    updateSettings({ identity: { name: "Ada Lovelace" } });
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    expect(formatRecords(editor)[0]!.author).toBe("Ada Lovelace");
    // Typed text revisions share the same identity.
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.insertText("X", 1, 1));
      return true;
    });
    const insertion = collectRevisions(editor.state.doc).find((r) => r.type === "insertion");
    expect(insertion?.author).toBe("Ada Lovelace");
    editor.destroy();
  });

  it("resolves name > initials > default", () => {
    expect(revisionAuthorOf({ name: "", initials: "" })).toBe("docen");
    expect(revisionAuthorOf({ name: "", initials: "AL" })).toBe("AL");
    expect(revisionAuthorOf({ name: "Ada Lovelace", initials: "AL" })).toBe("Ada Lovelace");
  });
});

describe.sequential("format-change marking", () => {
  it("records a bold toggle as a format change holding the old props", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    const marks = formatMarks(editor);
    expect(marks).toHaveLength(1);
    expect(marks[0]).toMatchObject({ from: 1, to: 6 });
    const [record] = parseFormatRecords(marks[0]!.attrs.records);
    expect(record).toMatchObject({ author: "docen" });
    expect(record!.props).toEqual({});
    expect(String(record!.date)).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/);
    expect(markNames(editor)).toContain("bold");
    editor.destroy();
  });

  it("records the old props when removing an existing mark", () => {
    const editor = build([{ type: "text", marks: [{ type: "bold" }], text: "hello world" }]);
    editor.commands["track-changes"](true);
    removeMark(editor, "bold", 1, 6);
    const records = formatRecords(editor);
    expect(records).toHaveLength(1);
    expect(records[0]!.props).toEqual({ bold: true });
    expect(markNames(editor)).not.toContain("bold");
    editor.destroy();
  });

  it("records one snapshot per run for a mixed selection", () => {
    const editor = build([
      { type: "text", marks: [{ type: "bold" }], text: "ab" },
      { type: "text", text: "cd" },
    ]);
    editor.commands["track-changes"](true);
    addMark(editor, "italic", 1, 5);
    const records = formatRecords(editor);
    // Word records per formatting run: "ab" was bold, "cd" was plain.
    expect(records).toHaveLength(2);
    expect(records[0]!.props).toEqual({ bold: true });
    expect(records[1]!.props).toEqual({});
    editor.destroy();
  });

  it("does not mark formatting edits while tracking is off", () => {
    const editor = build();
    addMark(editor, "bold", 1, 6);
    expect(formatMarks(editor)).toEqual([]);
    editor.destroy();
  });

  it("lets undo/redo replays through untouched (history$ meta)", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    editor.commands.command(({ state, dispatch }) => {
      const tr = state.tr;
      // prosemirror-history tags its replay trs with the "history$" meta;
      // the marking plugin must let them through unrecorded.
      tr.setMeta("history$", { redo: false });
      dispatch?.(tr.addMark(1, 6, state.schema.marks.bold!.create()));
      return true;
    });
    expect(formatMarks(editor)).toEqual([]);
    editor.destroy();
  });

  it("merges repeated same-author edits into one record with the original props", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 4);
    addMark(editor, "italic", 1, 4);
    const records = formatRecords(editor);
    expect(records).toHaveLength(1);
    expect(records[0]!.props).toEqual({});
    expect(markNames(editor)).toEqual(expect.arrayContaining(["bold", "italic"]));
    editor.destroy();
  });
});

describe.sequential("format-change accept/reject", () => {
  it("accept drops the record and keeps the new formatting", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    editor.commands["accept-change"]();
    expect(formatMarks(editor)).toEqual([]);
    expect(markNames(editor)).toContain("bold");
    editor.destroy();
  });

  it("reject drops the record and restores the old props", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    editor.commands["reject-change"]();
    expect(formatMarks(editor)).toEqual([]);
    expect(markNames(editor)).not.toContain("bold");
    editor.destroy();
  });

  it("reject restores an old mark that was removed", () => {
    const editor = build([{ type: "text", marks: [{ type: "bold" }], text: "hello world" }]);
    editor.commands["track-changes"](true);
    removeMark(editor, "bold", 1, 6);
    editor.commands["reject-change"]();
    expect(markNames(editor)).toContain("bold");
    expect(formatMarks(editor)).toEqual([]);
    editor.destroy();
  });

  it("reject restores the exact prior marks — no reconstructed carrier", () => {
    const editor = build([
      {
        type: "text",
        text: "hello world",
        marks: [{ type: "bold" }, { type: "textStyle", attrs: { size: 24, color: "00B050" } }],
      },
    ]);
    editor.commands["track-changes"](true);
    addMark(editor, "italic", 1, 6);
    editor.commands["reject-change"]();
    const marks = markNames(editor);
    // Exactly the prior structure: bold + ONE textStyle carrier (not two).
    expect(marks.filter((name) => name === "textStyle")).toHaveLength(1);
    expect(marks).toEqual(expect.arrayContaining(["bold", "textStyle"]));
    expect(marks).not.toContain("italic");
    const style = editor.state.doc
      .nodeAt(1)!
      .marks.find((mark: Mark) => mark.type.name === "textStyle");
    expect(style?.attrs).toMatchObject({ size: 24, color: "00B050" });
    // The restored carrier is not stale: removing bold afterwards unbolds.
    removeMark(editor, "bold", 1, 6);
    expect(markNames(editor)).not.toContain("bold");
    expect(compiledRuns(editor)[0]).not.toHaveProperty("bold");
    expect(compiledRuns(editor)[0]).toMatchObject({ size: 24, color: "00B050" });
    editor.destroy();
  });

  it("reject clears every rPr a single multi-mark transaction added", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    editor.commands.command(({ state, dispatch }) => {
      const tr = state.tr;
      tr.addMark(1, 6, state.schema.marks.bold!.create());
      tr.addMark(1, 6, state.schema.marks.italic!.create());
      tr.addMark(1, 6, state.schema.marks.textStyle!.create({ size: 24 }));
      tr.addMark(1, 6, state.schema.marks.textStyle!.create({ color: "00B050" }));
      dispatch?.(tr);
      return true;
    });
    // One author, one record — reject restores the plain before-state.
    expect(formatRecords(editor)).toHaveLength(1);
    expect(formatRecords(editor)[0]!.props).toEqual({});
    editor.commands["reject-change"]();
    expect(markNames(editor)).toEqual([]);
    expect(formatMarks(editor)).toEqual([]);
    editor.destroy();
  });

  it("rejecting a mixed selection's records never rewrites the untouched run", () => {
    const editor = build([
      { type: "text", marks: [{ type: "bold" }], text: "ab" },
      { type: "text", text: "cd" },
    ]);
    editor.commands["track-changes"](true);
    addMark(editor, "italic", 1, 5);
    const records = formatRecords(editor);
    // Reject newest first (the plain run), then the bold one.
    editor.commands["reject-change"](String(records[1]!.id));
    editor.commands["reject-change"](String(records[0]!.id));
    expect(formatMarks(editor)).toEqual([]);
    expect(markNames(editor, 1)).toContain("bold");
    expect(markNames(editor, 3)).not.toContain("bold");
    const runs = compiledRuns(editor);
    expect(runs.map((run) => run.text)).toEqual(["ab", "cd"]);
    expect(runs[0]).toMatchObject({ bold: true });
    expect(runs[1]).not.toHaveProperty("bold");
    editor.destroy();
  });

  it("reject by pane id targets the right record", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 4);
    addMark(editor, "bold", 5, 9);
    const revisions = collectRevisions(editor.state.doc);
    expect(revisions).toHaveLength(2);
    expect(revisions[0]!.type).toBe("format");
    editor.commands["reject-change"](String(revisions[0]!.id));
    // The first record is gone; the second (an independent record) keeps bold.
    expect(formatMarks(editor)).toHaveLength(1);
    expect(formatMarks(editor)[0]!.from).toBe(5);
    editor.destroy();
  });
});

describe.sequential("per-author run records", () => {
  const twoAuthors = (editor: EditorType): void => {
    updateSettings({ identity: { name: "Ada" } });
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    updateSettings({ identity: { name: "Bob" } });
    addMark(editor, "italic", 1, 6);
  };

  it("keeps separate records for different authors on the same range", () => {
    const editor = build();
    twoAuthors(editor);
    const records = formatRecords(editor);
    expect(records.map((record) => record.author)).toEqual(["Ada", "Bob"]);
    const revisions = collectRevisions(editor.state.doc);
    expect(revisions.map((revision) => revision.author)).toEqual(["Ada", "Bob"]);
    expect(revisions.map((revision) => revision.type)).toEqual(["format", "format"]);
    editor.destroy();
  });

  it("rejecting the newest author's record keeps the older author's change", () => {
    const editor = build();
    twoAuthors(editor);
    const records = formatRecords(editor);
    editor.commands["reject-change"](String(records[1]!.id));
    expect(markNames(editor)).toContain("bold");
    expect(markNames(editor)).not.toContain("italic");
    expect(formatRecords(editor).map((record) => record.author)).toEqual(["Ada"]);
    // Ada's own reject then removes her bold.
    editor.commands["reject-change"](String(records[0]!.id));
    expect(markNames(editor)).not.toContain("bold");
    expect(formatMarks(editor)).toEqual([]);
    editor.destroy();
  });

  it("rejecting the older author's record keeps the newer author's change", () => {
    const editor = build();
    twoAuthors(editor);
    const records = formatRecords(editor);
    editor.commands["reject-change"](String(records[0]!.id));
    // Ada's bold is reverted; Bob's italic and his record survive.
    expect(markNames(editor)).not.toContain("bold");
    expect(markNames(editor)).toContain("italic");
    expect(formatRecords(editor).map((record) => record.author)).toEqual(["Bob"]);
    editor.destroy();
  });
});

describe.sequential("paragraph format changes (w:pPrChange)", () => {
  const setAlignment = (editor: EditorType, alignment: string | null): void => {
    editor.commands.command(({ state, dispatch }) => {
      const attrs = state.doc.firstChild!.attrs as Record<string, unknown>;
      dispatch?.(state.tr.setNodeMarkup(0, undefined, { ...attrs, alignment }));
      return true;
    });
  };

  it("records a paragraph attr change with the old props", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    setAlignment(editor, "center");
    const attrs = paragraphAttrs(editor);
    expect(attrs.alignment).toBe("center");
    expect(attrs.revision).toMatchObject({ author: "docen" });
    expect((attrs.revision as { alignment?: string }).alignment).toBeUndefined();
    editor.destroy();
  });

  it("accept keeps the new paragraph attrs and drops the record", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    setAlignment(editor, "center");
    editor.commands["accept-change"]();
    expect(paragraphAttrs(editor).alignment).toBe("center");
    expect(paragraphAttrs(editor).revision).toBeNull();
    editor.destroy();
  });

  it("reject restores the old paragraph attrs", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    setAlignment(editor, "center");
    editor.commands["reject-change"]();
    expect(paragraphAttrs(editor).alignment).toBeNull();
    expect(paragraphAttrs(editor).revision).toBeNull();
    editor.destroy();
  });

  it("reject restores a previous non-default value", () => {
    const editor = build();
    editor.commands.command(({ state, dispatch }) => {
      // Seed the non-default prior value the way a loaded document carries it.
      dispatch?.(
        state.tr.setNodeMarkup(0, undefined, { ...paragraphAttrs(editor), alignment: "right" }),
      );
      return true;
    });
    editor.commands["track-changes"](true);
    setAlignment(editor, "center");
    expect((paragraphAttrs(editor).revision as { alignment?: string }).alignment).toBe("right");
    editor.commands["reject-change"]();
    expect(paragraphAttrs(editor).alignment).toBe("right");
    editor.destroy();
  });

  it("attributes a second author's change to that author", () => {
    const editor = build();
    updateSettings({ identity: { name: "Ada" } });
    editor.commands["track-changes"](true);
    setAlignment(editor, "center");
    updateSettings({ identity: { name: "Bob" } });
    setAlignment(editor, "right");
    const revision = paragraphAttrs(editor).revision as { author?: string; alignment?: string };
    expect(revision.author).toBe("Bob");
    // Rejecting Bob's change restores the state before HIS edit (center) —
    // never the first author's before-state.
    editor.commands["reject-change"]();
    expect(paragraphAttrs(editor).alignment).toBe("center");
    expect(paragraphAttrs(editor).revision).toBeNull();
    editor.destroy();
  });

  it("lists the paragraph change in collectRevisions with a prop diff", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    setAlignment(editor, "center");
    const [revision] = collectRevisions(editor.state.doc);
    expect(revision?.type).toBe("format");
    expect(revision?.paraPos).toBe(0);
    expect(revision?.text).toContain("alignment");
    editor.destroy();
  });
});

describe.sequential("sweeps, shown-scope and history", () => {
  it("reject-all restores formatting and removes insertions in one step", () => {
    const editor = build([{ type: "text", marks: [{ type: "bold" }], text: "hello world" }]);
    editor.commands["track-changes"](true);
    removeMark(editor, "bold", 1, 6);
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.insertText("XY", 11, 11));
      return true;
    });
    editor.commands["reject-all-changes"]();
    expect(editor.state.doc.textBetween(0, editor.state.doc.content.size)).toBe("hello world");
    expect(markNames(editor)).toContain("bold");
    expect(formatMarks(editor)).toEqual([]);
    editor.destroy();
  });

  it("scopes the shown sweeps to the filter's authors", () => {
    const editor = build();
    updateSettings({ identity: { name: "Ada" } });
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 4);
    updateSettings({ identity: { name: "Bob" } });
    addMark(editor, "italic", 5, 9);
    expect(formatRecords(editor)).toHaveLength(2);
    editor.commands["reject-all-changes-shown"](["Ada"]);
    expect(formatRecords(editor).map((record) => record.author)).toEqual(["Bob"]);
    expect(markNames(editor, 5)).toContain("italic");
    editor.commands["accept-all-changes-shown"](["Bob"]);
    expect(formatMarks(editor)).toEqual([]);
    expect(markNames(editor, 5)).toContain("italic");
    editor.destroy();
  });

  it("undo after reject restores the record, redo re-applies the reject", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    expect(formatMarks(editor)).toHaveLength(1);
    editor.commands["reject-change"]();
    expect(formatMarks(editor)).toHaveLength(0);
    expect(markNames(editor)).not.toContain("bold");

    editor.commands.undo();
    expect(formatMarks(editor)).toHaveLength(1);
    expect(markNames(editor)).toContain("bold");

    editor.commands.redo();
    expect(formatMarks(editor)).toEqual([]);
    expect(markNames(editor)).not.toContain("bold");
    editor.destroy();
  });
});
