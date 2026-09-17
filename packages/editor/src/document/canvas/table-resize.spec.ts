// @vitest-environment node
import { Document, Paragraph, Table, TableCell, TableRow } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import { describe, expect, it, vi } from "vitest";

import { DocumentCommands } from "../extensions/commands";
import { spanOf } from "./cell-selection";

// Canvas stub for CaretMap
const fakeCtx = {
  set font(_v: string) {},
  measureText: (text: string) => ({ width: text.length * 10 }),
} as unknown as CanvasRenderingContext2D;
vi.stubGlobal("document", {
  createElement: (tag: string) => (tag === "canvas" ? { getContext: () => fakeCtx } : {}),
} as unknown as Document);

const { CaretMap } = await import("./caret-map");

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

function makeTableJSON(columnWidths: number[], cells: string[][] = [["A", "B"]]): any {
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

describe("W1.1 table border interaction & resize", () => {
  describe("CaretMap tableBorderHitAt", () => {
    it("detects column borders within 3px tolerance and rejects outside tolerance", () => {
      const { doc } = build().state;
      // Construct a mock CaretMap with a tableZone
      const map = new CaretMap([], doc, () => ({ contentLeftPx: 0, contentTopPx: 0 }));
      (map.tableZones as any[]).push({
        page: 0,
        xPx: 50,
        yPx: 50,
        widthPx: 200,
        heightPx: 100,
        colEdges: [0, 100, 200],
        rowEdges: [0, 50, 100],
      });
      // Mock a cell box for cell pos mapping
      (map.cellBoxes as any).set(10, [{ page: 0, xPx: 50, yPx: 50, widthPx: 100, heightPx: 50 }]);

      // Border 0: x = 50. Within 3px:
      const hitLeft = map.tableBorderHitAt(0, 52, 75, 3);
      expect(hitLeft).not.toBeNull();
      expect(hitLeft?.kind).toBe("col");
      expect(hitLeft?.index).toBe(0);
      expect(hitLeft?.cellPos).toBe(10);

      // Border 1: x = 150. Within 3px:
      const hitMid = map.tableBorderHitAt(0, 151, 75, 3);
      expect(hitMid).not.toBeNull();
      expect(hitMid?.kind).toBe("col");
      expect(hitMid?.index).toBe(1);

      // Border 2: x = 250. Within 3px:
      const hitRight = map.tableBorderHitAt(0, 248, 75, 3);
      expect(hitRight).not.toBeNull();
      expect(hitRight?.kind).toBe("col");
      expect(hitRight?.index).toBe(2);

      // Outside 3px tolerance (x = 160 is 10px away from 150):
      const miss = map.tableBorderHitAt(0, 160, 75, 3);
      expect(miss).toBeNull();
    });

    it("detects row borders within 3px tolerance", () => {
      const { doc } = build().state;
      const map = new CaretMap([], doc, () => ({ contentLeftPx: 0, contentTopPx: 0 }));
      (map.tableZones as any[]).push({
        page: 0,
        xPx: 50,
        yPx: 50,
        widthPx: 200,
        heightPx: 100,
        colEdges: [0, 100, 200],
        rowEdges: [0, 50, 100],
      });
      (map.cellBoxes as any).set(10, [{ page: 0, xPx: 50, yPx: 50, widthPx: 100, heightPx: 50 }]);

      // Row border 1: y = 50 + 50 = 100.
      const hitRow = map.tableBorderHitAt(0, 120, 102, 3);
      expect(hitRow).not.toBeNull();
      expect(hitRow?.kind).toBe("row");
      expect(hitRow?.index).toBe(1);
      expect(hitRow?.cellPos).toBe(10);
    });
  });

  describe("AutoFit commands & tcW update", () => {
    it("autofit-contents updates columnWidths AND tcW attrs on all cells", () => {
      const editor = build(makeTableJSON([2000, 2000], [["甲乙丙", ""]]));
      editor.commands.setTextSelection(3);

      expect(editor.commands["autofit-contents"]()).toBe(true);
      const table = tablesOf(editor)[0]!;
      expect(table.attrs.columnWidths).toEqual([840, 720]);

      // Verify tcW (cell width) updated
      const row = table.child(0);
      expect(row.child(0).attrs.width).toEqual({ value: 840, type: "dxa" });
      expect(row.child(1).attrs.width).toEqual({ value: 720, type: "dxa" });

      // Single undo step restores previous widths
      editor.commands.undo();
      const restored = tablesOf(editor)[0]!;
      expect(restored.attrs.columnWidths).toEqual([2000, 2000]);
    });

    it("autofit-contents with targetCol updates only that column", () => {
      const editor = build(makeTableJSON([2000, 2000], [["甲乙丙", "丁"]]));
      editor.commands.setTextSelection(3);

      // Target column 0 only
      expect(editor.commands["autofit-contents"](0)).toBe(true);
      const table = tablesOf(editor)[0]!;
      // Col 0 shrunk to content (840), Col 1 preserved at 2000
      expect(table.attrs.columnWidths).toEqual([840, 2000]);
    });

    it("autofit-window rescales grid and stamps tcW in one undo step", () => {
      const editor = build(makeTableJSON([720, 1440]));
      editor.commands.setTextSelection(3);

      expect(editor.commands["autofit-window"]("1440")).toBe(true);
      const table = tablesOf(editor)[0]!;
      expect(table.attrs.columnWidths).toEqual([480, 960]);
      expect(table.child(0).child(0).attrs.width).toEqual({ value: 480, type: "dxa" });
      expect(table.child(0).child(1).attrs.width).toEqual({ value: 960, type: "dxa" });

      editor.commands.undo();
      expect(tablesOf(editor)[0]!.attrs.columnWidths).toEqual([720, 1440]);
    });
  });

  describe("Column and Row drag resize logic", () => {
    it("Shift-proportional neighbor resize preserves total table width", () => {
      const editor = build(makeTableJSON([1000, 1000]));
      const table = tablesOf(editor)[0]!;

      // Simulate Shift+drag resize on boundary between col 0 and col 1 (+200 twip)
      const deltaTwip = 200;
      const initialWidths = [...table.attrs.columnWidths];
      const nextWidths = [...initialWidths];
      const c = 0;
      const maxGrow = nextWidths[c + 1] - 360;
      const maxShrink = nextWidths[c] - 360;
      const clamped = Math.max(-maxShrink, Math.min(maxGrow, deltaTwip));
      nextWidths[c] += clamped;
      nextWidths[c + 1] -= clamped;

      expect(nextWidths).toEqual([1200, 800]);
      expect(nextWidths[0] + nextWidths[1]).toBe(initialWidths[0] + initialWidths[1]);

      // Dispatch single transaction
      const tr = editor.state.tr.setNodeMarkup(0, undefined, {
        ...table.attrs,
        columnWidths: nextWidths,
      });
      let curRowPos = 1;
      for (let r = 0; r < table.childCount; r++) {
        const row = table.child(r);
        let curCellPos = curRowPos + 1;
        let col = 0;
        for (let ci = 0; ci < row.childCount; ci++) {
          const cell = row.child(ci);
          const span = spanOf(cell);
          let cellWidthTwip = 0;
          for (let i = 0; i < span && col + i < nextWidths.length; i++) {
            cellWidthTwip += nextWidths[col + i]!;
          }
          tr.setNodeMarkup(curCellPos, undefined, {
            ...cell.attrs,
            width: { value: cellWidthTwip, type: "dxa" },
          });
          col += span;
          curCellPos += cell.nodeSize;
        }
        curRowPos += row.nodeSize;
      }
      editor.view.dispatch(tr);

      const updated = tablesOf(editor)[0]!;
      expect(updated.attrs.columnWidths).toEqual([1200, 800]);
      expect(updated.child(0).child(0).attrs.width).toEqual({ value: 1200, type: "dxa" });
      expect(updated.child(0).child(1).attrs.width).toEqual({ value: 800, type: "dxa" });

      // 1-step undo
      editor.commands.undo();
      expect(tablesOf(editor)[0]!.attrs.columnWidths).toEqual([1000, 1000]);
    });

    it("Row height resize stamps atLeast rule in one undo step", () => {
      const editor = build(makeTableJSON([1000, 1000]));
      const table = tablesOf(editor)[0]!;

      // Set height for row 0
      const tr = editor.state.tr.setNodeMarkup(1, undefined, {
        ...table.child(0).attrs,
        height: { value: 720, rule: "atLeast" },
      });
      editor.view.dispatch(tr);

      expect(tablesOf(editor)[0]!.child(0).attrs.height).toEqual({
        value: 720,
        rule: "atLeast",
      });

      editor.commands.undo();
      expect(tablesOf(editor)[0]!.child(0).attrs.height).toBeNull();
    });
  });
});
