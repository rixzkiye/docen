import { Schema } from "@tiptap/pm/model";
import { EditorState } from "@tiptap/pm/state";
import { describe, expect, it } from "vitest";

import { computeBlockRanges, copyBlockRanges, deleteBlockRanges } from "./block-selection";

describe("Block Selection (W2.3 Alt+drag rectangular selection)", () => {
  const schema = new Schema({
    nodes: {
      doc: { content: "paragraph+" },
      paragraph: { content: "text*", toDOM: () => ["p", 0] },
      text: { inline: true },
    },
  });

  it("computes rectangular ranges from line layout using posAtPoint", () => {
    const mockCaretMap = {
      lines: [
        {
          page: 0,
          yPx: 100,
          line: { heightPx: 20 },
        },
        {
          page: 0,
          yPx: 125,
          line: { heightPx: 20 },
        },
        {
          page: 0,
          yPx: 150,
          line: { heightPx: 20 },
        },
        {
          page: 1, // different page
          yPx: 100,
          line: { heightPx: 20 },
        },
      ],
      posAtPoint: (page: number, x: number, y: number) => {
        if (page !== 0) return null;
        // Mock pos based on line y and x
        const lineIdx = y < 120 ? 0 : y < 145 ? 1 : 2;
        const base = 1 + lineIdx * 30;
        const col = Math.round((x - 50) / 10);
        return base + Math.max(0, col);
      },
    };

    const ranges = computeBlockRanges(mockCaretMap, 0, 80, 95, 150, 155);
    expect(ranges.length).toBe(3);
    // Line 0: [1 + 3, 1 + 10] = [4, 11]
    expect(ranges[0]).toEqual({ from: 4, to: 11 });
    // Line 1: [31 + 3, 31 + 10] = [34, 41]
    expect(ranges[1]).toEqual({ from: 34, to: 41 });
    // Line 2: [61 + 3, 61 + 10] = [64, 71]
    expect(ranges[2]).toEqual({ from: 64, to: 71 });
  });

  it("copies and deletes rectangular block text", () => {
    const p1 = "ColumnA  ColumnB  ColumnC";
    const p2 = "Alpha1   Beta1    Gamma1";
    const p3 = "Alpha2   Beta2    Gamma2";
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text(p1)]),
      schema.node("paragraph", null, [schema.text(p2)]),
      schema.node("paragraph", null, [schema.text(p3)]),
    ]);
    let state = EditorState.create({ doc, schema });
    const mockView = {
      dispatch: (tr: any) => {
        state = state.apply(tr);
      },
    };
    const mockEditor = {
      get state() {
        return state;
      },
      get view() {
        return mockView;
      },
    } as any;

    // p1 starts at pos 1:
    const from1 = 1 + p1.indexOf("ColumnB");
    const to1 = from1 + "ColumnB".length;
    // p2 starts at 1 + p1.length + 2:
    const base2 = 1 + p1.length + 2;
    const from2 = base2 + p2.indexOf("Beta1");
    const to2 = from2 + "Beta1".length;
    // p3 starts at base2 + p2.length + 2:
    const base3 = base2 + p2.length + 2;
    const from3 = base3 + p3.indexOf("Beta2");
    const to3 = from3 + "Beta2".length;

    const ranges = [
      { from: from1, to: to1 },
      { from: from2, to: to2 },
      { from: from3, to: to3 },
    ];

    const copied = copyBlockRanges(mockEditor, ranges);
    expect(copied).toBe("ColumnB\nBeta1\nBeta2");

    deleteBlockRanges(mockEditor, ranges);
    expect(state.doc.textContent).toBe("ColumnA    ColumnCAlpha1       Gamma1Alpha2       Gamma2");
  });
});
