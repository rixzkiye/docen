// @vitest-environment happy-dom
import { Document, InlinePassthrough, Paragraph, type JSONContent } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { DialogCommands, type DialogsHost } from "./dialogs";
import {
  formatBibliographyEntry,
  formatInTextCitation,
  type BibliographySource,
} from "./references";

const fieldAtom = (branch: object): JSONContent =>
  ({
    type: "inlinePassthrough",
    attrs: { data: JSON.stringify(branch) },
  }) as JSONContent;

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

const extrasOf = (): Record<string, unknown> => ({
  footnotes: [
    {
      id: 1,
      children: [{ style: "FootnoteText", children: [{ footnoteRef: true }, { text: "原始" }] }],
    },
  ],
});

const build = (): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough],
    content: {
      type: "doc",
      attrs: { documentExtras: extrasOf() },
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "ab" },
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ footnoteReference: 1 }) },
            },
            { type: "text", text: "c" },
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ endnoteReference: { id: 7 } }) },
            },
          ],
        },
      ],
    },
  });

const host = (editor: EditorType): DialogsHost => ({
  editor: () => editor,
  bridge: () => undefined,
  // No shadow root — #noteDialog()'s querySelector stays undefined and the
  // dialog opens are skipped (the commits under test don't need the element).
  element: () => ({ shadowRoot: null }) as unknown as HTMLElement,
  syncStatusLanguage: () => {},
});

const noteOk = (detail: unknown): Event => ({ detail }) as unknown as Event;

describe("DialogCommands note targets", () => {
  it("resolves the reference atom after the caret and just before it", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    // layout: 1 "ab" | 3 footnote atom | 4 "c" | 5 endnote atom
    editor.commands.setTextSelection(4);
    expect(dialogs.noteTarget()).toEqual({ kind: "footnote", id: 1, pos: 3 });
  });

  it("reads the object-form id and the endnote kind, hit from both sides", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(5); // the atom starts here
    expect(dialogs.noteTarget()).toEqual({ kind: "endnote", id: 7, pos: 5 });
    editor.commands.setTextSelection(6); // past it — the atom ends just before
    expect(dialogs.noteTarget()).toEqual({ kind: "endnote", id: 7, pos: 5 });
  });

  it("returns null when the caret sits on plain text", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(2);
    expect(dialogs.noteTarget()).toBeNull();
  });
});

describe("DialogCommands note commits", () => {
  it("insert appends a marker-led body per textarea line", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(2);
    dialogs.onNoteOk(noteOk({ kind: "footnote", text: "第一行\n第二行" }));
    const notes = (
      editor.state.doc.attrs.documentExtras as { footnotes: Array<Record<string, unknown>> }
    ).footnotes;
    expect(notes.map((n) => n.id)).toEqual([1, 2]);
    expect(notes[1].children).toEqual([
      { style: "FootnoteText", children: [{ footnoteRef: true }, { text: "第一行" }] },
      { style: "FootnoteText", children: [{ text: "第二行" }] },
    ]);
  });

  it("edit rewrites the referenced body in place, keeping style and marker", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4);
    dialogs.noteEditAtSelection(); // sets the pending target (no dialog here)
    dialogs.onNoteOk(noteOk({ kind: "footnote", text: "改后" }));
    const notes = (
      editor.state.doc.attrs.documentExtras as { footnotes: Array<Record<string, unknown>> }
    ).footnotes;
    expect(notes.map((n) => n.id)).toEqual([1]);
    expect(notes[0].children).toEqual([
      { style: "FootnoteText", children: [{ footnoteRef: true }, { text: "改后" }] },
    ]);
  });

  it("insert with empty text is a no-op", () => {
    const editor = build();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(2);
    dialogs.onNoteOk(noteOk({ kind: "footnote", text: "" }));
    const notes = (
      editor.state.doc.attrs.documentExtras as { footnotes: Array<Record<string, unknown>> }
    ).footnotes;
    expect(notes.map((n) => n.id)).toEqual([1]);
  });
});

// Field tests: layout 1 "ab" | 3 NUMCHARS atom | 4 "c".
const buildField = (): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, InlinePassthrough],
    content: {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "ab" },
            fieldAtom({ simpleField: { instruction: "NUMCHARS", cachedValue: "9" } }),
            { type: "text", text: "c" },
          ],
        },
      ],
    },
  });

describe("DialogCommands field targets", () => {
  it("resolves the field atom from both sides of the caret", () => {
    const editor = buildField();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4); // right after the atom
    const hit = dialogs.fieldTarget();
    expect(hit?.pos).toBe(3);
    expect(hit?.ref).toEqual({
      kind: "simpleField",
      instruction: "NUMCHARS",
      result: "9",
    });
    editor.commands.setTextSelection(3); // right at the atom
    expect(dialogs.fieldTarget()?.pos).toBe(3);
    editor.commands.setTextSelection(2); // plain text — nothing
    expect(dialogs.fieldTarget()).toBeNull();
  });
});

describe("DialogCommands field commits", () => {
  it("update re-derives the cached value from the live text", () => {
    const editor = buildField();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4);
    dialogs.fieldUpdateAtSelection();
    const node = editor.state.doc.nodeAt(3);
    expect(JSON.parse(String(node?.attrs.data))).toEqual({
      simpleField: { instruction: "NUMCHARS", cachedValue: "3" },
    });
  });

  it("dialog OK on a fresh dialog seeds an evaluated field at the caret", () => {
    const editor = buildField();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(2);
    dialogs.fieldInsert(); // clears the (unset) edit target — insert mode
    dialogs.onFieldOk({ detail: { instruction: 'DATE \\@ "yyyy"' } } as unknown as Event);
    const seeded = JSON.parse(String(editor.state.doc.nodeAt(2)?.attrs.data));
    expect(seeded.simpleField.instruction).toBe('DATE \\@ "yyyy"');
    expect(seeded.simpleField.cachedValue).toBe("2026");
  });

  it("dialog OK after edit rewrites the atom's instruction in place", () => {
    const editor = buildField();
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(4);
    dialogs.fieldEditAtSelection();
    dialogs.onFieldOk({ detail: { instruction: 'DATE \\@ "yy"' } } as unknown as Event);
    const node = editor.state.doc.nodeAt(3);
    expect(JSON.parse(String(node?.attrs.data))).toEqual({
      simpleField: { instruction: 'DATE \\@ "yy"', cachedValue: "26" },
    });
  });

  it("checkbox toggle flips checked and keeps the rest of the branch", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [fieldAtom({ formField: { checkBox: { checked: false, size: 20 } } })],
          },
        ],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection(1);
    expect(dialogs.fieldTarget()?.ref.kind).toBe("formField");
    dialogs.fieldToggleCheckboxAtSelection();
    expect(JSON.parse(String(editor.state.doc.nodeAt(1)?.attrs.data))).toEqual({
      formField: { checkBox: { checked: true, size: 20 } },
    });
  });
});

describe("DialogCommands note conversion", () => {
  it("converts all footnotes to endnotes with updated IDs and node references", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        attrs: {
          documentExtras: {
            footnotes: [
              {
                id: 1,
                children: [
                  {
                    style: "FootnoteText",
                    children: [{ footnoteRef: true }, { text: "Note 1" }],
                  },
                ],
              },
            ],
            endnotes: [
              {
                id: 1,
                children: [
                  { style: "EndnoteText", children: [{ endnoteRef: true }, { text: "End 1" }] },
                ],
              },
            ],
          },
        },
        content: [
          {
            type: "paragraph",
            content: [
              { type: "text", text: "text" },
              {
                type: "inlinePassthrough",
                attrs: { data: JSON.stringify({ footnoteReference: 1 }) },
              },
            ],
          },
        ],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.convertNotes("allFootnotesToEndnotes")).toBe(true);
    const extras = editor.state.doc.attrs.documentExtras as {
      footnotes: unknown[];
      endnotes: Array<{ id: number; children: unknown[] }>;
    };
    expect(extras.footnotes).toHaveLength(0);
    expect(extras.endnotes).toHaveLength(2);
    expect(extras.endnotes[1].id).toBe(2);
    const atomData = JSON.parse(String(editor.state.doc.nodeAt(5)?.attrs.data));
    expect(atomData).toEqual({ endnoteReference: 2 });
  });

  it("converts all endnotes to footnotes with updated IDs and node references", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        attrs: {
          documentExtras: {
            footnotes: [
              {
                id: 1,
                children: [
                  {
                    style: "FootnoteText",
                    children: [{ footnoteRef: true }, { text: "Note 1" }],
                  },
                ],
              },
            ],
            endnotes: [
              {
                id: 1,
                children: [
                  { style: "EndnoteText", children: [{ endnoteRef: true }, { text: "End 1" }] },
                ],
              },
            ],
          },
        },
        content: [
          {
            type: "paragraph",
            content: [
              { type: "text", text: "text" },
              {
                type: "inlinePassthrough",
                attrs: { data: JSON.stringify({ endnoteReference: 1 }) },
              },
            ],
          },
        ],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.convertNotes("allEndnotesToFootnotes")).toBe(true);
    const extras = editor.state.doc.attrs.documentExtras as {
      footnotes: Array<{ id: number; children: unknown[] }>;
      endnotes: unknown[];
    };
    expect(extras.endnotes).toHaveLength(0);
    expect(extras.footnotes).toHaveLength(2);
    expect(extras.footnotes[1].id).toBe(2);
    const atomData = JSON.parse(String(editor.state.doc.nodeAt(5)?.attrs.data));
    expect(atomData).toEqual({ footnoteReference: 2 });
  });

  it("swaps footnotes and endnotes", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        attrs: {
          documentExtras: {
            footnotes: [
              {
                id: 1,
                children: [
                  { style: "FootnoteText", children: [{ footnoteRef: true }, { text: "FN" }] },
                ],
              },
            ],
            endnotes: [
              {
                id: 2,
                children: [
                  { style: "EndnoteText", children: [{ endnoteRef: true }, { text: "EN" }] },
                ],
              },
            ],
          },
        },
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "inlinePassthrough",
                attrs: { data: JSON.stringify({ footnoteReference: 1 }) },
              },
              {
                type: "inlinePassthrough",
                attrs: { data: JSON.stringify({ endnoteReference: 2 }) },
              },
            ],
          },
        ],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.convertNotes("swapNotes")).toBe(true);
    const extras = editor.state.doc.attrs.documentExtras as {
      footnotes: Array<{ id: number; children: unknown[] }>;
      endnotes: Array<{ id: number; children: unknown[] }>;
    };
    expect(extras.footnotes[0].id).toBe(2);
    expect(extras.endnotes[0].id).toBe(1);
    const atom1 = JSON.parse(String(editor.state.doc.nodeAt(1)?.attrs.data));
    const atom2 = JSON.parse(String(editor.state.doc.nodeAt(2)?.attrs.data));
    expect(atom1).toEqual({ endnoteReference: 1 });
    expect(atom2).toEqual({ footnoteReference: 2 });
  });
});

describe("DialogCommands bookmarks", () => {
  it("adds, lists, and deletes bookmarks", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [{ type: "text", text: "Hello world" }],
          },
        ],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    editor.commands.setTextSelection({ from: 1, to: 6 });
    expect(dialogs.addBookmark("myBookmark")).toBe(true);

    const bms = dialogs.documentBookmarks();
    expect(bms).toHaveLength(1);
    expect(bms[0].name).toBe("myBookmark");
    expect(bms[0].from).toBe(2);

    expect(dialogs.deleteBookmark("myBookmark")).toBe(true);
    expect(dialogs.documentBookmarks()).toHaveLength(0);
  });

  it("rejects invalid bookmark names", () => {
    const editor = new Editor({
      element: null,
      extensions: [Document, Paragraph, Text, InlinePassthrough],
      content: {
        type: "doc",
        content: [{ type: "paragraph", content: [{ type: "text", text: "Test" }] }],
      },
    });
    const dialogs = new DialogCommands(host(editor));
    expect(dialogs.addBookmark("123invalid")).toBe(false);
    expect(dialogs.addBookmark("has space")).toBe(false);
  });
});

describe("Citation and Bibliography formatters", () => {
  const source: BibliographySource = {
    tag: "Turing1936",
    sourceType: "Book",
    title: "On Computable Numbers",
    year: "1936",
    publisher: "London Mathematical Society",
    author: {
      authors: [{ first: "Alan", last: "Turing" }],
    },
  };

  const multiAuthorSource: BibliographySource = {
    tag: "KnuthEtAl",
    sourceType: "Book",
    title: "Concrete Mathematics",
    year: "1994",
    publisher: "Addison-Wesley",
    author: {
      authors: [
        { first: "Ronald", last: "Graham" },
        { first: "Donald", last: "Knuth" },
        { first: "Oren", last: "Patashnik" },
      ],
    },
  };

  it("formats in-text citations per style", () => {
    expect(formatInTextCitation(source, "APA")).toBe("(Turing, 1936)");
    expect(formatInTextCitation(source, "MLA")).toBe("(Turing)");
    expect(formatInTextCitation(source, "CHICAGO")).toBe("(Turing 1936)");
    expect(formatInTextCitation(source, "IEEE", 3)).toBe("[3]");

    expect(formatInTextCitation(multiAuthorSource, "APA")).toBe("(Graham et al., 1994)");
    expect(formatInTextCitation(multiAuthorSource, "MLA")).toBe("(Graham et al.)");
    expect(formatInTextCitation(multiAuthorSource, "CHICAGO")).toBe("(Graham et al. 1994)");
    expect(formatInTextCitation(multiAuthorSource, "IEEE", 1)).toBe("[1]");
  });

  it("formats bibliography entries per style", () => {
    const apa = formatBibliographyEntry(source, "APA", 1);
    expect(apa).toBe("Turing, A. (1936). On Computable Numbers. London Mathematical Society.");

    const mla = formatBibliographyEntry(source, "MLA", 1);
    expect(mla).toBe("Turing, Alan. On Computable Numbers. London Mathematical Society, 1936.");

    const chicago = formatBibliographyEntry(source, "CHICAGO", 1);
    expect(chicago).toBe("Turing, Alan. 1936. On Computable Numbers. London Mathematical Society.");

    const ieee = formatBibliographyEntry(source, "IEEE", 1);
    expect(ieee).toBe("[1] A. Turing, On Computable Numbers. London Mathematical Society, 1936.");
  });
});
