import {
  Bold,
  Deletion,
  Document,
  FormatChange,
  Insertion,
  Paragraph,
  TextStyle,
} from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
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
  const editor = new Editor({
    element: null,
    extensions: [
      Document,
      Paragraph,
      Text,
      TextStyle,
      Bold,
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

/** Mark names on the first text node (the paragraph's leading run). */
const markNames = (editor: EditorType): string[] => {
  const text = editor.state.doc.firstChild?.firstChild;
  return (text?.marks ?? []).map((mark: Mark) => mark.type.name);
};

const paragraphAttrs = (editor: EditorType): Record<string, unknown> =>
  editor.state.doc.firstChild!.attrs as Record<string, unknown>;

afterEach(() => {
  // The settings store is module-global (in-memory in the test env) — clear
  // the identity so the fallback tests stay independent of run order.
  updateSettings({ identity: { name: "" } });
});

describe("revision author identity", () => {
  it("falls back to Word's historical default when no identity is set", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    expect(formatMarks(editor)[0]!.attrs.author).toBe("docen");
    editor.destroy();
  });

  it("uses the settings user name for new revisions (text and format)", () => {
    updateSettings({ identity: { name: "Ada Lovelace" } });
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    expect(formatMarks(editor)[0]!.attrs.author).toBe("Ada Lovelace");
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

describe("format-change marking", () => {
  it("records a bold toggle as a format change holding the old props", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 6);
    const marks = formatMarks(editor);
    expect(marks).toHaveLength(1);
    expect(marks[0]).toMatchObject({ from: 1, to: 6 });
    expect(JSON.parse(marks[0]!.attrs.props as string)).toEqual({});
    expect(markNames(editor)).toContain("bold");
    expect(String(marks[0]!.attrs.date)).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/);
    editor.destroy();
  });

  it("records the old props when removing an existing mark", () => {
    const editor = build([{ type: "text", marks: [{ type: "bold" }], text: "hello world" }]);
    editor.commands["track-changes"](true);
    removeMark(editor, "bold", 1, 6);
    const marks = formatMarks(editor);
    expect(marks).toHaveLength(1);
    expect(JSON.parse(marks[0]!.attrs.props as string)).toEqual({ bold: true });
    expect(markNames(editor)).not.toContain("bold");
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
    removeMark(editor, "bold", 1, 4);
    const marks = formatMarks(editor);
    expect(marks).toHaveLength(1);
    expect(JSON.parse(marks[0]!.attrs.props as string)).toEqual({});
    editor.destroy();
  });
});

describe("format-change accept/reject", () => {
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

  it("reject by pane id targets the right record", () => {
    const editor = build();
    editor.commands["track-changes"](true);
    addMark(editor, "bold", 1, 4);
    addMark(editor, "bold", 5, 9);
    const revisions = collectRevisions(editor.state.doc);
    expect(revisions).toHaveLength(2);
    expect(revisions[0]!.type).toBe("format");
    editor.commands["reject-change"](String(revisions[0]!.id));
    // The first record is gone; the second (an independent insertion mark)
    // keeps its bold.
    expect(formatMarks(editor)).toHaveLength(1);
    expect(formatMarks(editor)[0]!.from).toBe(5);
    editor.destroy();
  });
});

describe("paragraph format changes (w:pPrChange)", () => {
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
    const editor = new Editor({
      element: null,
      extensions: [
        Document,
        Paragraph,
        Text,
        TextStyle,
        Bold,
        Insertion,
        Deletion,
        FormatChange,
        TrackChanges,
      ],
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { alignment: "right" },
            content: [{ type: "text", text: "hello world" }],
          },
        ],
      },
    });
    for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
    editor.commands["track-changes"](true);
    setAlignment(editor, "center");
    expect((paragraphAttrs(editor).revision as { alignment?: string }).alignment).toBe("right");
    editor.commands["reject-change"]();
    expect(paragraphAttrs(editor).alignment).toBe("right");
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

describe("mixed accept/reject sweeps", () => {
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
});
