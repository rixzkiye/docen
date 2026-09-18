// @vitest-environment happy-dom
import { Document, Paragraph, Table, TableCell, TableRow } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import { describe, expect, it } from "vitest";

import type { RibbonMenuItem } from "../../ui/addin/types";
import { appendMenuItems } from "../../ui/components/ribbon/command-helpers";
import { DocumentCommands, WIRED_DISPATCH } from "../extensions/commands";

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
    ["10", "20"],
    ["30", ""],
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

describe("W1.3 table context menu & submenus", () => {
  describe("Submenu rendering in command-helpers", () => {
    it("renders submenu hierarchy into fluent-menu-list", () => {
      const list = document.createElement("div");
      const selected: RibbonMenuItem[] = [];

      const menuItems: RibbonMenuItem[] = [
        {
          text: "Insert",
          items: [
            { text: "Insert Column Left", event: "insert-column-left" },
            { text: "Insert Column Right", event: "insert-column-right" },
            { text: "-" },
            { text: "Insert Row Above", event: "insert-row-above" },
            { text: "Insert Row Below", event: "insert-row-below" },
            { text: "-" },
            { text: "Insert Cells...", event: "insert-cells" },
          ],
        },
        { text: "Merge Cells", event: "merge-cells" },
      ];

      appendMenuItems(list, menuItems, (item) => selected.push(item));

      // Parent items
      const parentItems = list.children;
      expect(parentItems.length).toBe(2);

      const insertItem = parentItems[0] as HTMLElement;
      expect(insertItem.hasAttribute("data-has-submenu")).toBe(true);

      // Submenu list
      const subList = insertItem.querySelector("fluent-menu-list");
      expect(subList).not.toBeNull();
      expect(subList?.getAttribute("slot")).toBe("submenu");

      const subChildren = subList?.children ?? [];
      // 5 items + 2 dividers = 7 items
      expect(subChildren.length).toBe(7);

      const mergeItem = parentItems[1] as HTMLElement;
      expect(mergeItem.hasAttribute("data-has-submenu")).toBe(false);
      expect(mergeItem.textContent).toBe("Merge Cells");
    });
  });

  describe("Table context menu commands dispatch", () => {
    it("registers all required table context commands in WIRED_DISPATCH", () => {
      const expectedCommands = [
        "insert-column-left",
        "insert-column-right",
        "insert-row-above",
        "insert-row-below",
        "insert-cell",
        "insert-cells",
        "delete-cell",
        "delete-cells",
        "delete-column",
        "delete-row",
        "delete-table",
        "select-table",
        "select-table-row",
        "select-table-cell",
        "select-table-column",
        "merge-cells",
        "split-cell",
        "split-table",
        "autofit-contents",
        "autofit-window",
        "fixed-column-width",
        "align-cell",
        "distribute-rows",
        "distribute-columns",
        "text-direction",
        "cell-margins",
        "repeat-header-rows",
        "table-formula",
        "formula",
      ];

      for (const cmd of expectedCommands) {
        expect(WIRED_DISPATCH.has(cmd), `command ${cmd} should be in WIRED_DISPATCH`).toBe(true);
      }
    });

    it("inserts a cell in current row using insert-cell command", () => {
      const editor = build(makeTableJSON([1000, 1000], [["A", "B"]]));
      editor.commands.setTextSelection(3); // Inside cell "A"

      const beforeTable = tablesOf(editor)[0]!;
      expect(beforeTable.child(0).childCount).toBe(2);

      expect(editor.commands["insert-cell"]()).toBe(true);

      const afterTable = tablesOf(editor)[0]!;
      expect(afterTable.child(0).childCount).toBe(3);
    });

    it("deletes a cell using delete-cell command", () => {
      const editor = build(makeTableJSON([1000, 1000, 1000], [["A", "B", "C"]]));
      editor.commands.setTextSelection(3); // Inside cell "A"

      expect(editor.commands["delete-cell"]()).toBe(true);

      const afterTable = tablesOf(editor)[0]!;
      expect(afterTable.child(0).childCount).toBe(2);
      expect(afterTable.child(0).child(0).textContent).toBe("B");
    });

    it("calculates table formula sum from above cells", () => {
      // 2x2 table: Row 0 has "10", "20"; Row 1 has cell 0: "30", cell 1: ""
      const editor = build(
        makeTableJSON(
          [1000, 1000],
          [
            ["10", "20"],
            ["30", ""],
          ],
        ),
      );
      // Find position of the empty cell (row 1, col 1)
      let targetPos = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === "paragraph" && node.textContent === "" && targetPos < 0) {
          targetPos = pos + 1;
        }
      });
      expect(targetPos).toBeGreaterThan(0);
      editor.commands.setTextSelection(targetPos);

      expect(editor.commands["table-formula"]()).toBe(true);

      const table = tablesOf(editor)[0]!;
      expect(table.child(1).child(1).textContent).toBe("20");
    });

    it("calculates table formula sum from left cells when no cells above have numbers", () => {
      // 1x3 table: "15", "25", ""
      const editor = build(makeTableJSON([1000, 1000, 1000], [["15", "25", ""]]));
      let targetPos = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === "paragraph" && node.textContent === "" && targetPos < 0) {
          targetPos = pos + 1;
        }
      });
      expect(targetPos).toBeGreaterThan(0);
      editor.commands.setTextSelection(targetPos);

      expect(editor.commands["table-formula"]()).toBe(true);

      const table = tablesOf(editor)[0]!;
      expect(table.child(0).child(2).textContent).toBe("40");
    });
  });
});
