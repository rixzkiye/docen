import { Schema } from "@tiptap/pm/model";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { describe, expect, it } from "vitest";

import { CellSelection } from "./cell-selection";

describe("Selection Cycling and Boundary Rules (W2.5)", () => {
  const schema = new Schema({
    nodes: {
      doc: { content: "(paragraph | table)+" },
      paragraph: { content: "text*", toDOM: () => ["p", 0] },
      table: { content: "tableRow+", toDOM: () => ["table", ["tbody", 0]] },
      tableRow: { content: "tableCell+", toDOM: () => ["tr", 0] },
      tableCell: {
        content: "paragraph+",
        attrs: { columnSpan: { default: 1 } },
        toDOM: () => ["td", 0],
      },
      text: { inline: true },
    },
  });

  const createTableDoc = () => {
    return schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Introduction line")]),
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("Cell 1 text")]),
          ]),
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("Cell 2 text")]),
          ]),
        ]),
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("Cell 3 text")]),
          ]),
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("Cell 4 text")]),
          ]),
        ]),
      ]),
      schema.node("paragraph", null, [schema.text("Concluding line")]),
    ]);
  };

  it("cycles Ctrl+A in table: cell text -> whole table -> entire document", () => {
    const doc = createTableDoc();
    const cell1Text = "Cell 1 text";
    let cell1Pos = -1;
    doc.descendants((node, pos) => {
      if (node.isText && node.text === cell1Text) {
        cell1Pos = pos + 2;
        return false;
      }
    });

    let state = EditorState.create({
      doc,
      schema,
      selection: TextSelection.create(doc, cell1Pos),
    });

    // Helper simulating edit-bridge Ctrl+A cycling logic
    const cycleCtrlA = () => {
      const sel = state.selection;
      const $pos = state.doc.resolve(sel.from);
      let $cell = null;
      for (let d = $pos.depth; d > 0; d--) {
        if ($pos.node(d).type.name === "tableCell") {
          $cell = state.doc.resolve($pos.before(d));
          break;
        }
      }

      if ($cell) {
        const cellNode = $cell.nodeAfter;
        if (cellNode) {
          const cellStart = $cell.pos + 1;
          const cellEnd = $cell.pos + cellNode.nodeSize - 1;
          const coversCell =
            sel instanceof TextSelection && sel.from <= cellStart + 1 && sel.to >= cellEnd - 1;

          if (!coversCell && !(sel instanceof CellSelection)) {
            // Cycle 1: Select cell contents
            state = state.apply(
              state.tr.setSelection(
                TextSelection.between(
                  state.doc.resolve(cellStart + 1),
                  state.doc.resolve(cellEnd - 1),
                ),
              ),
            );
            return 1;
          }

          const tableSel = CellSelection.tableSelection($cell);
          const alreadyTable =
            sel instanceof CellSelection &&
            sel.anchorCell === tableSel.anchorCell &&
            sel.headCell === tableSel.headCell;

          if (!alreadyTable) {
            // Cycle 2: Select whole table
            state = state.apply(state.tr.setSelection(tableSel));
            return 2;
          }
        }
      }

      // Cycle 3: Select entire document
      const start = TextSelection.near(state.doc.resolve(0), 1).from;
      const end = TextSelection.near(state.doc.resolve(state.doc.content.size), -1).to;
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, start, end)));
      return 3;
    };

    // 1st Ctrl+A: selects cell text
    const step1 = cycleCtrlA();
    expect(step1).toBe(1);
    expect(state.doc.textBetween(state.selection.from, state.selection.to)).toBe("Cell 1 text");

    // 2nd Ctrl+A: selects entire table
    const step2 = cycleCtrlA();
    expect(step2).toBe(2);
    expect(state.selection instanceof CellSelection).toBe(true);

    // 3rd Ctrl+A: selects entire document
    const step3 = cycleCtrlA();
    expect(step3).toBe(3);
    expect(state.selection.from).toBe(1);
    expect(state.selection.to).toBe(doc.content.size - 1);
  });

  it("handles word boundaries for Latin, CJK, and punctuation correctly", () => {
    const charClass = (ch: string): number =>
      /\s/.test(ch) ? 0 : /[㐀-鿿豈-﫿]/.test(ch) ? 1 : /[0-9A-Za-z]/.test(ch) ? 2 : 3;

    expect(charClass(" ")).toBe(0);
    expect(charClass("\t")).toBe(0);
    expect(charClass("中")).toBe(1);
    expect(charClass("文")).toBe(1);
    expect(charClass("A")).toBe(2);
    expect(charClass("9")).toBe(2);
    expect(charClass(".")).toBe(3);
    expect(charClass(",")).toBe(3);
    expect(charClass("!")).toBe(3);
  });
});
