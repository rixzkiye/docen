// @vitest-environment happy-dom
import { Document, Paragraph, Table, TableCell, TableRow } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import { describe, expect, it } from "vitest";

import { DocumentCommands, WIRED_DISPATCH } from "../extensions/commands";
import { CaretMap } from "./caret-map";

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

function makeTableJSON(
  columnWidths: number[],
  cells: string[][] = [
    ["A", "B"],
    ["C", "D"],
  ],
): any {
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

const build = (initialTable?: any): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: EXTENSIONS,
    content: {
      type: "doc",
      content: initialTable ? [initialTable] : [{ type: "paragraph" }],
    },
  });
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

function tablesOf(editor: EditorType) {
  const tables: any[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === "table") {
      tables.push(node);
      return false;
    }
    return true;
  });
  return tables;
}

describe("W1.4 Draw Table & Table Eraser tools", () => {
  describe("Registration & Tool commands", () => {
    it("registers draw-table and table-eraser in WIRED_DISPATCH", () => {
      expect(WIRED_DISPATCH.has("draw-table")).toBe(true);
      expect(WIRED_DISPATCH.has("table-eraser")).toBe(true);
      expect(WIRED_DISPATCH.has("draw-table-stroke")).toBe(true);
      expect(WIRED_DISPATCH.has("table-eraser-click")).toBe(true);
    });
  });

  describe("Draw Table tool (draw-table-stroke)", () => {
    it("creates min 1x1 table when dragging outside a table in one undo step", () => {
      const editor = build();
      expect(tablesOf(editor).length).toBe(0);

      // Drag small box (100px x 40px)
      expect(
        (editor.commands as any)["draw-table-stroke"]({
          widthPx: 100,
          heightPx: 40,
          inTable: false,
        }),
      ).toBe(true);

      const tables = tablesOf(editor);
      expect(tables.length).toBe(1);
      const table = tables[0]!;
      expect(table.childCount).toBe(1); // 1 row
      expect(table.child(0).childCount).toBe(1); // 1 col

      // Undo step
      editor.commands.undo();
      expect(tablesOf(editor).length).toBe(0);
    });

    it("grows table grid dimensions based on dragged stroke size", () => {
      const editor = build();

      // Drag large box (300px x 150px) => 2 cols, 2 rows
      expect(
        (editor.commands as any)["draw-table-stroke"]({
          widthPx: 300,
          heightPx: 150,
          inTable: false,
        }),
      ).toBe(true);

      const table = tablesOf(editor)[0]!;
      expect(table.childCount).toBe(2); // 2 rows
      expect(table.child(0).childCount).toBe(2); // 2 cols
      expect(table.attrs.columnWidths).toEqual([2250, 2250]); // 4500 twip / 2
    });

    it("splits row below when drawing horizontal stroke inside existing cell", () => {
      const editor = build(makeTableJSON([1000, 1000], [["A", "B"]]));
      editor.commands.setTextSelection(3); // inside cell "A"

      const beforeTable = tablesOf(editor)[0]!;
      expect(beforeTable.childCount).toBe(1);

      // Horizontal stroke: dx = 60, dy = 5
      expect(
        (editor.commands as any)["draw-table-stroke"]({
          dx: 60,
          dy: 5,
          inTable: true,
        }),
      ).toBe(true);

      const afterTable = tablesOf(editor)[0]!;
      expect(afterTable.childCount).toBe(2);

      // 1 undo step
      editor.commands.undo();
      expect(tablesOf(editor)[0]!.childCount).toBe(1);
    });

    it("splits col to right when drawing vertical stroke inside existing cell", () => {
      const editor = build(makeTableJSON([1000, 1000], [["A", "B"]]));
      editor.commands.setTextSelection(3); // inside cell "A"

      const beforeTable = tablesOf(editor)[0]!;
      expect(beforeTable.child(0).childCount).toBe(2);

      // Vertical stroke: dx = 5, dy = 60
      expect(
        (editor.commands as any)["draw-table-stroke"]({
          dx: 5,
          dy: 60,
          inTable: true,
        }),
      ).toBe(true);

      const afterTable = tablesOf(editor)[0]!;
      expect(afterTable.child(0).childCount).toBe(3);

      // 1 undo step
      editor.commands.undo();
      expect(tablesOf(editor)[0]!.child(0).childCount).toBe(2);
    });
  });

  describe("Table Eraser tool (table-eraser-click)", () => {
    it("merges adjacent cells when clicking an interior border in one undo step", () => {
      const editor = build(makeTableJSON([1000, 1000], [["A", "B"]]));
      // Cell A: pos 2, Cell B: pos 7
      let posA = -1;
      let posB = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === "tableCell") {
          if (posA < 0) posA = pos;
          else if (posB < 0) posB = pos;
        }
      });
      expect(posA).toBeGreaterThan(0);
      expect(posB).toBeGreaterThan(0);

      // Interior border between cell A and cell B has 2 sides (right of A and left of B)
      expect(
        (editor.commands as any)["table-eraser-click"]({
          sides: [
            { pos: posA, side: "right" },
            { pos: posB, side: "left" },
          ],
        }),
      ).toBe(true);

      const table = tablesOf(editor)[0]!;
      // Row folded into 1 cell with columnSpan = 2
      expect(table.child(0).childCount).toBe(1);
      expect(table.child(0).child(0).attrs.columnSpan).toBe(2);

      // 1 undo step restores the 2 cells
      editor.commands.undo();
      const restored = tablesOf(editor)[0]!;
      expect(restored.child(0).childCount).toBe(2);
    });

    it("erases outer border when clicking an exterior table border in one undo step", () => {
      const editor = build(makeTableJSON([1000, 1000], [["A", "B"]]));
      let posA = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === "tableCell" && posA < 0) {
          posA = pos;
        }
      });

      // Exterior left border of cell A
      expect(
        (editor.commands as any)["table-eraser-click"]({
          sides: [{ pos: posA, side: "left" }],
        }),
      ).toBe(true);

      const cell = tablesOf(editor)[0]!.child(0).child(0);
      expect(cell.attrs.borders?.left?.style).toBe("nil");

      // 1 undo step restores border
      editor.commands.undo();
      const restoredCell = tablesOf(editor)[0]!.child(0).child(0);
      expect(restoredCell.attrs.borders?.left).toBeUndefined();
    });
  });

  describe("CaretMap cellAtPoint helper", () => {
    it("finds cell position and rect covering point coordinates", () => {
      const { doc } = build().state;
      const map = new CaretMap([], doc, () => ({ contentLeftPx: 0, contentTopPx: 0 }));
      (map.cellBoxes as any).set(10, [{ page: 0, xPx: 50, yPx: 50, widthPx: 100, heightPx: 50 }]);

      const found = map.cellAtPoint(0, 75, 75);
      expect(found).not.toBeNull();
      expect(found?.pos).toBe(10);

      const outside = map.cellAtPoint(0, 200, 200);
      expect(outside).toBeNull();
    });
  });
});
