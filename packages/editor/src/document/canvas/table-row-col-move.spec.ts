// @vitest-environment happy-dom
import { Document, Paragraph, Table, TableCell, TableRow } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import { describe, expect, it } from "vitest";

import { DocumentCommands, WIRED_DISPATCH } from "../extensions/commands";
import { KEYBOARD_SHORTCUTS } from "../extensions/keymap";

const Text = TextNode.create({ name: "text", group: "inline" });

const EXTENSIONS = [
  Document,
  Paragraph,
  Text,
  Table,
  TableRow,
  TableCell,
  DocumentCommands,
  UndoRedo,
];

function makeTableJSON(columnWidths: number[], cells: string[][]): any {
  return {
    type: "table",
    attrs: { columnWidths },
    content: cells.map((row) => ({
      type: "tableRow",
      content: row.map((text) => ({
        type: "tableCell",
        content: [{ type: "paragraph", content: text ? [{ type: "text", text }] : [] }],
      })),
    })),
  };
}

const build = (initialContent?: any[]): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: EXTENSIONS,
    content: {
      type: "doc",
      content: initialContent ?? [{ type: "paragraph" }],
    },
  });
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

function getTableRowsText(editor: EditorType): string[][] {
  const res: string[][] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === "tableRow") {
      const rowText: string[] = [];
      node.forEach((cell) => {
        rowText.push(cell.textContent);
      });
      res.push(rowText);
    }
  });
  return res;
}

function getTableColWidths(editor: EditorType): number[] {
  let widths: number[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === "table") {
      widths = node.attrs.columnWidths ?? [];
    }
  });
  return widths;
}

describe("W1.5 Table row/column move & quick insert", () => {
  describe("Registration and Keybindings", () => {
    it("registers move and insert-at commands in WIRED_DISPATCH", () => {
      const commands = [
        "move-row-up",
        "move-row-down",
        "move-row",
        "move-column",
        "insert-row-at",
        "insert-column-at",
      ];
      for (const cmd of commands) {
        expect(WIRED_DISPATCH.has(cmd)).toBe(true);
      }
    });

    it("binds Alt-Shift-ArrowUp and Alt-Shift-ArrowDown in KEYBOARD_SHORTCUTS", () => {
      expect(KEYBOARD_SHORTCUTS["Alt-Shift-ArrowUp"]).toBe("move-row-up");
      expect(KEYBOARD_SHORTCUTS["Alt-Shift-ArrowDown"]).toBe("move-row-down");
    });
  });

  describe("move-row-up and move-row-down (Alt+Shift+Up/Down)", () => {
    it("moves table row up when caret is in middle row, with 1-step undo", () => {
      const initial = makeTableJSON(
        [1000, 1000],
        [
          ["R0C0", "R0C1"],
          ["R1C0", "R1C1"],
          ["R2C0", "R2C1"],
        ],
      );
      const editor = build([initial]);

      // Place caret inside row 1 ("R1C0")
      let r1Pos = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.isText && node.text === "R1C0" && r1Pos < 0) {
          r1Pos = pos;
        }
      });
      expect(r1Pos).toBeGreaterThan(0);
      editor.commands.setTextSelection(r1Pos);

      expect((editor.commands as any)["move-row-up"]()).toBe(true);
      expect(getTableRowsText(editor)).toEqual([
        ["R1C0", "R1C1"],
        ["R0C0", "R0C1"],
        ["R2C0", "R2C1"],
      ]);

      // 1-step undo
      editor.commands.undo();
      expect(getTableRowsText(editor)).toEqual([
        ["R0C0", "R0C1"],
        ["R1C0", "R1C1"],
        ["R2C0", "R2C1"],
      ]);
    });

    it("does nothing (returns false) when moving row-up at row 0", () => {
      const initial = makeTableJSON(
        [1000, 1000],
        [
          ["R0C0", "R0C1"],
          ["R1C0", "R1C1"],
        ],
      );
      const editor = build([initial]);

      let r0Pos = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.isText && node.text === "R0C0" && r0Pos < 0) {
          r0Pos = pos;
        }
      });
      editor.commands.setTextSelection(r0Pos);

      expect((editor.commands as any)["move-row-up"]()).toBe(false);
      expect(getTableRowsText(editor)).toEqual([
        ["R0C0", "R0C1"],
        ["R1C0", "R1C1"],
      ]);
    });

    it("moves table row down when caret is in middle row, with 1-step undo", () => {
      const initial = makeTableJSON(
        [1000, 1000],
        [
          ["R0C0", "R0C1"],
          ["R1C0", "R1C1"],
          ["R2C0", "R2C1"],
        ],
      );
      const editor = build([initial]);

      let r1Pos = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.isText && node.text === "R1C0" && r1Pos < 0) {
          r1Pos = pos;
        }
      });
      editor.commands.setTextSelection(r1Pos);

      expect((editor.commands as any)["move-row-down"]()).toBe(true);
      expect(getTableRowsText(editor)).toEqual([
        ["R0C0", "R0C1"],
        ["R2C0", "R2C1"],
        ["R1C0", "R1C1"],
      ]);

      // 1-step undo
      editor.commands.undo();
      expect(getTableRowsText(editor)).toEqual([
        ["R0C0", "R0C1"],
        ["R1C0", "R1C1"],
        ["R2C0", "R2C1"],
      ]);
    });

    it("does nothing (returns false) when moving row-down at bottom row", () => {
      const initial = makeTableJSON(
        [1000, 1000],
        [
          ["R0C0", "R0C1"],
          ["R1C0", "R1C1"],
        ],
      );
      const editor = build([initial]);

      let r1Pos = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.isText && node.text === "R1C0" && r1Pos < 0) {
          r1Pos = pos;
        }
      });
      editor.commands.setTextSelection(r1Pos);

      expect((editor.commands as any)["move-row-down"]()).toBe(false);
      expect(getTableRowsText(editor)).toEqual([
        ["R0C0", "R0C1"],
        ["R1C0", "R1C1"],
      ]);
    });

    it("moves paragraph outline blocks up and down when outside table", () => {
      const editor = build([
        { type: "paragraph", content: [{ type: "text", text: "First paragraph" }] },
        { type: "paragraph", content: [{ type: "text", text: "Second paragraph" }] },
        { type: "paragraph", content: [{ type: "text", text: "Third paragraph" }] },
      ]);

      // Place caret in "Second paragraph"
      let p2Pos = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.isText && node.text === "Second paragraph" && p2Pos < 0) {
          p2Pos = pos;
        }
      });
      editor.commands.setTextSelection(p2Pos);

      expect((editor.commands as any)["move-row-up"]()).toBe(true);
      const textsAfterUp: string[] = [];
      editor.state.doc.forEach((node) => textsAfterUp.push(node.textContent));
      expect(textsAfterUp).toEqual(["Second paragraph", "First paragraph", "Third paragraph"]);

      expect((editor.commands as any)["move-row-down"]()).toBe(true);
      const textsAfterDown: string[] = [];
      editor.state.doc.forEach((node) => textsAfterDown.push(node.textContent));
      expect(textsAfterDown).toEqual(["First paragraph", "Second paragraph", "Third paragraph"]);
    });
  });

  describe("move-row (drag-and-drop row reorder)", () => {
    it("moves row from top to bottom (index 0 to 3)", () => {
      const initial = makeTableJSON(
        [1000, 1000],
        [
          ["R0", "0"],
          ["R1", "1"],
          ["R2", "2"],
        ],
      );
      const editor = build([initial]);

      editor.commands.setTextSelection(3); // Inside table
      expect(
        (editor.commands as any)["move-row"]({
          fromIndex: 0,
          toIndex: 3,
        }),
      ).toBe(true);

      expect(getTableRowsText(editor)).toEqual([
        ["R1", "1"],
        ["R2", "2"],
        ["R0", "0"],
      ]);

      // 1-step undo
      editor.commands.undo();
      expect(getTableRowsText(editor)).toEqual([
        ["R0", "0"],
        ["R1", "1"],
        ["R2", "2"],
      ]);
    });

    it("moves row from bottom to top (index 2 to 0)", () => {
      const initial = makeTableJSON(
        [1000, 1000],
        [
          ["R0", "0"],
          ["R1", "1"],
          ["R2", "2"],
        ],
      );
      const editor = build([initial]);

      editor.commands.setTextSelection(3);
      expect(
        (editor.commands as any)["move-row"]({
          fromIndex: 2,
          toIndex: 0,
        }),
      ).toBe(true);

      expect(getTableRowsText(editor)).toEqual([
        ["R2", "2"],
        ["R0", "0"],
        ["R1", "1"],
      ]);
    });
  });

  describe("move-column (drag-and-drop column reorder)", () => {
    it("reorders columns and columnWidths", () => {
      const initial = makeTableJSON(
        [1000, 2000, 3000],
        [
          ["C0", "C1", "C2"],
          ["D0", "D1", "D2"],
        ],
      );
      const editor = build([initial]);

      editor.commands.setTextSelection(3); // Inside table
      expect(
        (editor.commands as any)["move-column"]({
          fromIndex: 0,
          toIndex: 3,
        }),
      ).toBe(true);

      expect(getTableRowsText(editor)).toEqual([
        ["C1", "C2", "C0"],
        ["D1", "D2", "D0"],
      ]);
      expect(getTableColWidths(editor)).toEqual([2000, 3000, 1000]);

      // 1-step undo
      editor.commands.undo();
      expect(getTableRowsText(editor)).toEqual([
        ["C0", "C1", "C2"],
        ["D0", "D1", "D2"],
      ]);
      expect(getTableColWidths(editor)).toEqual([1000, 2000, 3000]);
    });
  });

  describe("insert-row-at and insert-column-at (hover quick-insert +)", () => {
    it("inserts an empty row at top boundary (index 0)", () => {
      const initial = makeTableJSON(
        [1000, 1000],
        [
          ["R0C0", "R0C1"],
          ["R1C0", "R1C1"],
        ],
      );
      const editor = build([initial]);

      editor.commands.setTextSelection(3);
      expect(
        (editor.commands as any)["insert-row-at"]({
          index: 0,
        }),
      ).toBe(true);

      const rows = getTableRowsText(editor);
      expect(rows.length).toBe(3);
      expect(rows[0]).toEqual(["", ""]);
      expect(rows[1]).toEqual(["R0C0", "R0C1"]);
      expect(rows[2]).toEqual(["R1C0", "R1C1"]);

      editor.commands.undo();
      expect(getTableRowsText(editor).length).toBe(2);
    });

    it("inserts an empty row between existing rows (index 1)", () => {
      const initial = makeTableJSON(
        [1000, 1000],
        [
          ["R0C0", "R0C1"],
          ["R1C0", "R1C1"],
        ],
      );
      const editor = build([initial]);

      editor.commands.setTextSelection(3);
      expect(
        (editor.commands as any)["insert-row-at"]({
          index: 1,
        }),
      ).toBe(true);

      const rows = getTableRowsText(editor);
      expect(rows.length).toBe(3);
      expect(rows[0]).toEqual(["R0C0", "R0C1"]);
      expect(rows[1]).toEqual(["", ""]);
      expect(rows[2]).toEqual(["R1C0", "R1C1"]);
    });

    it("inserts an empty column between existing columns (index 1)", () => {
      const initial = makeTableJSON(
        [1000, 2000],
        [
          ["C0", "C1"],
          ["D0", "D1"],
        ],
      );
      const editor = build([initial]);

      editor.commands.setTextSelection(3);
      expect(
        (editor.commands as any)["insert-column-at"]({
          index: 1,
        }),
      ).toBe(true);

      const rows = getTableRowsText(editor);
      expect(rows[0]).toEqual(["C0", "", "C1"]);
      expect(rows[1]).toEqual(["D0", "", "D1"]);
      expect(getTableColWidths(editor).length).toBe(3);

      editor.commands.undo();
      expect(getTableRowsText(editor)[0]).toEqual(["C0", "C1"]);
    });
  });
});
