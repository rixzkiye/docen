// @vitest-environment node
// Quick Parts insertion / F3 against a real headless editor — the pre-flight
// matrix: document start/end/middle, table cells, table blocks, marks/lists,
// single-transaction undo/redo, and large blocks.
import {
  Bold,
  Document,
  Paragraph,
  selectionSlicePayload,
  Table,
  TableCell,
  TableRow,
} from "@docen/docx";
import {
  Editor,
  Node as TextNode,
  type Editor as EditorType,
  type JSONContent,
} from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import { describe, expect, it } from "vitest";

import { withBlocks, type BuildingBlock, type BuildingBlockSlice } from "../building-blocks";
import { parseSlicePayload } from "../building-blocks";
import { CellSelection } from "../canvas/cell-selection";
import { DocumentCommands } from "./commands";

// Tiptap's schema needs the plain text node (same trick as the sibling specs).
const Text = TextNode.create({ name: "text", group: "inline" });

const EXTENSIONS = [
  Document,
  Paragraph,
  Text,
  Table,
  TableRow,
  TableCell,
  Bold,
  DocumentCommands,
  UndoRedo,
];

const para = (text: string, marks?: JSONContent["marks"]): JSONContent => ({
  type: "paragraph",
  ...(text ? { content: [{ type: "text", text, ...(marks ? { marks } : {}) }] } : {}),
});

const makeBlock = (
  id: string,
  name: string,
  content: JSONContent[],
  openStart = 0,
  openEnd = 0,
): BuildingBlock => ({
  id,
  name,
  gallery: "custQuickParts",
  category: "General",
  description: "",
  savedAt: 0,
  insertMode: "content",
  content: { openStart, openEnd, content },
});

const setBlocks = (editor: EditorType, blocks: BuildingBlock[]): void => {
  editor.commands.command(({ state, dispatch }) => {
    dispatch?.(
      state.tr.setDocAttribute(
        "documentExtras",
        withBlocks(state.doc.attrs.documentExtras, blocks),
      ),
    );
    return true;
  });
};

const build = (blocks: BuildingBlock[] = [], content?: JSONContent): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: EXTENSIONS,
    content: content ?? { type: "doc", content: [para("")] },
  });
  // element:null skips mount(), and with it plugin installation — register the
  // extension manager's plugins by hand (the bridge does the same) so undo
  // history and keymaps are live.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  setBlocks(editor, blocks);
  return editor;
};

const textOf = (editor: EditorType): string =>
  editor.state.doc.textBetween(0, editor.state.doc.content.size, "\n", "\n");

describe("insert-building-block", () => {
  it("inserts a paragraph block at the document start", () => {
    const editor = build([makeBlock("a", "Signature", [para("Best regards,")])], {
      type: "doc",
      content: [para("Hello world")],
    });
    editor.commands.setTextSelection(1);
    expect(editor.commands["insert-building-block"]("a")).toBe(true);
    expect(textOf(editor)).toBe("Best regards,\nHello world");
  });

  it("inserts at the end and mid-paragraph", () => {
    const blocks = [makeBlock("a", "Sig", [para("BB")])];
    const atEnd = build(blocks, { type: "doc", content: [para("Hello")] });
    atEnd.commands.setTextSelection(6);
    expect(atEnd.commands["insert-building-block"]("a")).toBe(true);
    expect(textOf(atEnd)).toBe("Hello\nBB");

    const mid = build(blocks, { type: "doc", content: [para("Hello world")] });
    mid.commands.setTextSelection(6); // after "Hello"
    expect(mid.commands["insert-building-block"]("a")).toBe(true);
    expect(textOf(mid)).toBe("Hello\nBB\n world");
  });

  it("inserts inside a table cell", () => {
    const editor = build([makeBlock("a", "Sig", [para("Cell block")])], {
      type: "doc",
      content: [
        {
          type: "table",
          content: [
            {
              type: "tableRow",
              content: [{ type: "tableCell", content: [para("cell")] }],
            },
          ],
        },
      ],
    });
    editor.commands.setTextSelection(4); // inside the cell paragraph
    expect(editor.commands["insert-building-block"]("a")).toBe(true);
    expect(textOf(editor)).toContain("Cell block");
    expect(editor.state.doc.firstChild!.type.name).toBe("table");
  });

  it("inserts a table block into the body and into a cell (nested table)", () => {
    const tableBlock = makeBlock("t", "Grid", [
      {
        type: "table",
        content: [
          {
            type: "tableRow",
            content: [
              { type: "tableCell", content: [para("a")] },
              { type: "tableCell", content: [para("b")] },
            ],
          },
        ],
      },
    ]);
    const body = build([tableBlock], { type: "doc", content: [para("x")] });
    body.commands.setTextSelection(1);
    expect(body.commands["insert-building-block"]("t")).toBe(true);
    // At the paragraph start the table lands before the split remainder.
    expect(body.state.doc.child(0).type.name).toBe("table");
    expect(body.state.doc.child(1).textContent).toBe("x");

    const inCell = build([tableBlock], {
      type: "doc",
      content: [
        {
          type: "table",
          content: [
            { type: "tableRow", content: [{ type: "tableCell", content: [para("cell")] }] },
          ],
        },
      ],
    });
    inCell.commands.setTextSelection(4);
    expect(inCell.commands["insert-building-block"]("t")).toBe(true);
    const cell = inCell.state.doc.firstChild!.firstChild!.firstChild!;
    expect(cell.textContent).toContain("a");
    expect(cell.textContent).toContain("b");
  });

  it("keeps marks and list attrs from the saved slice", () => {
    const editor = build(
      [
        makeBlock("m", "Marked", [
          para("bold", [{ type: "bold" }]),
          {
            type: "paragraph",
            attrs: { bullet: { level: 0 } },
            content: [{ type: "text", text: "item" }],
          },
        ]),
      ],
      { type: "doc", content: [para("x")] },
    );
    editor.commands.setTextSelection(2);
    expect(editor.commands["insert-building-block"]("m")).toBe(true);
    expect(JSON.stringify(editor.getJSON())).toContain('"bold"');
    expect(JSON.stringify(editor.getJSON())).toContain('"bullet"');
  });

  it("merges a partial-paragraph slice (open depths) with the caret paragraph", () => {
    const slice: BuildingBlockSlice = {
      openStart: 1,
      openEnd: 1,
      content: [para("SELECTED")],
    };
    const editor = build(
      [makeBlock("p", "Partial", slice.content, slice.openStart, slice.openEnd)],
      { type: "doc", content: [para("Hello world")] },
    );
    editor.commands.setTextSelection(12); // end of "Hello world"
    expect(editor.commands["insert-building-block"]("p")).toBe(true);
    expect(textOf(editor)).toBe("Hello worldSELECTED");
  });

  it("declines unknown ids and corrupt content without touching the doc", () => {
    const editor = build([makeBlock("a", "A", [para("x")])]);
    const before = JSON.stringify(editor.getJSON());
    expect(editor.commands["insert-building-block"]("missing")).toBe(false);
    // A block whose content uses an unknown node type cannot fit the schema.
    setBlocks(editor, [makeBlock("bad", "Bad", [{ type: "nonsenseNode" }])]);
    expect(editor.commands["insert-building-block"]("bad")).toBe(false);
    expect(JSON.stringify(editor.getJSON().content)).toBe(
      JSON.stringify(JSON.parse(before).content),
    );
  });

  it("applies as exactly one transaction and round-trips through undo/redo", () => {
    const editor = build([makeBlock("a", "Signature", [para("Best regards,")])], {
      type: "doc",
      content: [para("Hello")],
    });
    const before = JSON.stringify(editor.getJSON());
    let docChanges = 0;
    editor.on("transaction", ({ transaction }) => {
      if (transaction.docChanged) docChanges++;
    });
    editor.commands.setTextSelection(6);
    expect(editor.commands["insert-building-block"]("a")).toBe(true);
    expect(docChanges).toBe(1);
    expect(textOf(editor)).toBe("Hello\nBest regards,");

    editor.commands.undo();
    expect(JSON.stringify(editor.getJSON())).toBe(before);
    editor.commands.redo();
    expect(textOf(editor)).toBe("Hello\nBest regards,");
  });

  it("inserts a large block as one transaction", () => {
    const many = Array.from({ length: 200 }, (_, i) => para(`line ${i}`));
    const editor = build([makeBlock("big", "Big", many)], { type: "doc", content: [para("x")] });
    let docChanges = 0;
    editor.on("transaction", ({ transaction }) => {
      if (transaction.docChanged) docChanges++;
    });
    editor.commands.setTextSelection(2);
    expect(editor.commands["insert-building-block"]("big")).toBe(true);
    expect(docChanges).toBe(1);
    expect(textOf(editor)).toContain("line 0");
    expect(textOf(editor)).toContain("line 199");
  });

  it("reads blocks saved on the doc attrs (not a private registry)", () => {
    const editor = build([], { type: "doc", content: [para("x")] });
    editor.commands.setTextSelection(1);
    expect(editor.commands["insert-building-block"]("late")).toBe(false);
    setBlocks(editor, [makeBlock("late", "Late", [para("now")])]);
    expect(editor.commands["insert-building-block"]("late")).toBe(true);
    expect(textOf(editor)).toContain("now");
  });
});

describe("autotext-f3", () => {
  const block = makeBlock("sig", "Signature", [para("Best regards,")]);

  it("replaces the typed name with the matching block", () => {
    const editor = build([block], {
      type: "doc",
      content: [para("Please sign: Signature")],
    });
    editor.commands.setTextSelection(editor.state.doc.content.size - 1); // end of paragraph
    let docChanges = 0;
    editor.on("transaction", ({ transaction }) => {
      if (transaction.docChanged) docChanges++;
    });
    expect(editor.commands["autotext-f3"]()).toBe(true);
    expect(docChanges).toBe(1);
    expect(textOf(editor)).toBe("Please sign: \nBest regards,");
  });

  it("matches case-insensitively and prefers the longest name", () => {
    const short = makeBlock("s", "Block", [para("short")]);
    const long = makeBlock("l", "My Block", [para("long")]);
    const editor = build([short, long], { type: "doc", content: [para("My Block")] });
    editor.commands.setTextSelection(9);
    expect(editor.commands["autotext-f3"]()).toBe(true);
    expect(textOf(editor)).toBe("long");
  });

  it("declines a partial name, a mid-word caret and no match", () => {
    const cases: Array<[string, number]> = [
      ["Signat", 7],
      ["xSignature", 10],
      ["Nothing here", 5],
    ];
    for (const [text, caret] of cases) {
      const editor = build([block], { type: "doc", content: [para(text)] });
      editor.commands.setTextSelection(caret);
      const before = JSON.stringify(editor.getJSON());
      expect(editor.commands["autotext-f3"]()).toBe(false);
      expect(JSON.stringify(editor.getJSON())).toBe(before);
    }
  });

  it("declines when no blocks exist", () => {
    const editor = build([], { type: "doc", content: [para("Signature")] });
    editor.commands.setTextSelection(10);
    expect(editor.commands["autotext-f3"]()).toBe(false);
  });

  it("undo restores the typed name in one step", () => {
    const editor = build([block], { type: "doc", content: [para("Signature")] });
    editor.commands.setTextSelection(10);
    const before = JSON.stringify(editor.getJSON());
    expect(editor.commands["autotext-f3"]()).toBe(true);
    editor.commands.undo();
    expect(JSON.stringify(editor.getJSON())).toBe(before);
  });
});

describe("save selection → insert round-trip (real slice capture)", () => {
  /** Capture the current selection the way the host's save dialog does. */
  const captured = (editor: EditorType, name: string): BuildingBlock => {
    const slice = parseSlicePayload(selectionSlicePayload(editor.state));
    expect(slice).not.toBeNull();
    return makeBlock("cap", name, slice!.content, slice!.openStart, slice!.openEnd);
  };

  it("saves a selection spanning two paragraphs and reinserts it", () => {
    const source = build([], { type: "doc", content: [para("Alpha one"), para("Beta two")] });
    // "one" through "Be" in the next paragraph — a slice open at both ends.
    source.commands.setTextSelection({ from: 7, to: 14 });
    const saved = captured(source, "Spanning");
    expect(saved.content.openStart).toBe(1);
    expect(saved.content.openEnd).toBe(1);

    const target = build([saved], { type: "doc", content: [para("")] });
    target.commands.setTextSelection(1);
    expect(target.commands["insert-building-block"]("cap")).toBe(true);
    expect(textOf(target)).toBe("one\nBe");
  });

  it("saves a whole table selection and reinserts it", () => {
    const table: JSONContent = {
      type: "table",
      content: [
        {
          type: "tableRow",
          content: [
            { type: "tableCell", content: [para("a")] },
            { type: "tableCell", content: [para("b")] },
          ],
        },
      ],
    };
    const source = build([], { type: "doc", content: [table] });
    source.commands.setNodeSelection(0);
    const saved = captured(source, "Whole table");
    expect(saved.content.content[0]?.type).toBe("table");

    const target = build([saved], { type: "doc", content: [para("")] });
    target.commands.setTextSelection(1);
    expect(target.commands["insert-building-block"]("cap")).toBe(true);
    expect(target.state.doc.firstChild!.type.name).toBe("table");
    expect(target.state.doc.firstChild!.textContent).toContain("a");
  });

  it("saves a selection spanning table cells sanely (no store corruption)", () => {
    const table: JSONContent = {
      type: "table",
      content: [
        {
          type: "tableRow",
          content: [
            { type: "tableCell", content: [para("a")] },
            { type: "tableCell", content: [para("b")] },
          ],
        },
      ],
    };
    const source = build([], { type: "doc", content: [table, para("after")] });
    const cells: number[] = [];
    source.state.doc.descendants((node, pos) => {
      if (node.type.name === "tableCell") cells.push(pos);
    });
    source.view.dispatch(
      source.state.tr.setSelection(CellSelection.create(source.state.doc, cells[0]!, cells[1]!)),
    );
    const slice = parseSlicePayload(selectionSlicePayload(source.state));
    expect(slice).not.toBeNull();
    const saved = makeBlock("cells", "Cells", slice!.content, slice!.openStart, slice!.openEnd);
    // The store projection must survive a cell-spanning slice…
    expect(() => withBlocks(undefined, [saved])).not.toThrow();

    // …and re-inserting it at a body caret must land the table content.
    const target = build([saved], { type: "doc", content: [para("")] });
    target.commands.setTextSelection(1);
    expect(target.commands["insert-building-block"]("cells")).toBe(true);
    expect(textOf(target)).toContain("a");
    expect(textOf(target)).toContain("b");
  });
});
