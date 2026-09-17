import type { StylesOptions, TableCellOptions, TableOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { activeConditionalTypes, resolveTableLook, resolveTableStyle } from "../style-cascade";
import type { ProjectContext } from "./project/context";
import { projectTable } from "./project/table";

function makeCtx(styles?: StylesOptions): ProjectContext {
  return {
    styles,
    characterStyles: new Map(),
    numberings: new Map(),
    listCounters: new Map(),
    openComments: new Set(),
    footnoteOrdinals: new Map(),
    endnoteOrdinals: new Map(),
    defaultTabStopPx: 48,
    revisionAuthorColors: new Map(),
  };
}

describe("W1.2 TableLook & Conditional Table Styles Cascade", () => {
  describe("resolveTableLook", () => {
    it("returns Word defaults when look is empty or undefined", () => {
      const def = resolveTableLook(undefined);
      expect(def).toEqual({
        firstRow: true,
        lastRow: false,
        bandRow: true,
        firstCol: false,
        lastCol: false,
        bandCol: false,
      });
    });

    it("decodes Word hex bitmask val (e.g. 04A0)", () => {
      // 0x04A0: firstRow (0x20) ON, firstCol (0x80) ON, noHBand (0x200) OFF => bandRow ON, noVBand (0x400) ON => bandCol OFF
      const look = resolveTableLook({ val: "04A0" });
      expect(look.firstRow).toBe(true);
      expect(look.firstCol).toBe(true);
      expect(look.lastRow).toBe(false);
      expect(look.lastCol).toBe(false);
      expect(look.bandRow).toBe(true);
      expect(look.bandCol).toBe(false);
    });

    it("decodes bitmask with noHBand enabled (bandRow disabled)", () => {
      // 0x0220: firstRow ON (0x20), noHBand ON (0x200 => bandRow false)
      const look = resolveTableLook({ val: "0220" });
      expect(look.firstRow).toBe(true);
      expect(look.bandRow).toBe(false);
    });

    it("explicit booleans override val bitmask", () => {
      const look = resolveTableLook({ val: "04A0", firstRow: false, bandCol: true });
      expect(look.firstRow).toBe(false);
      expect(look.bandCol).toBe(true);
    });
  });

  describe("activeConditionalTypes", () => {
    const look = {
      firstRow: true,
      lastRow: true,
      firstCol: true,
      lastCol: true,
      bandRow: true,
      bandCol: true,
    };

    it("detects corner cells (nwCell, neCell, swCell, seCell)", () => {
      // 3x3 table
      const totalRows = 3;
      const totalCols = 3;

      const nw = activeConditionalTypes({ rowIndex: 0, colIndex: 0, totalRows, totalCols }, look);
      expect(nw.has("nwCell")).toBe(true);
      expect(nw.has("firstRow")).toBe(true);
      expect(nw.has("firstCol")).toBe(true);

      const ne = activeConditionalTypes({ rowIndex: 0, colIndex: 2, totalRows, totalCols }, look);
      expect(ne.has("neCell")).toBe(true);
      expect(ne.has("firstRow")).toBe(true);
      expect(ne.has("lastCol")).toBe(true);

      const sw = activeConditionalTypes({ rowIndex: 2, colIndex: 0, totalRows, totalCols }, look);
      expect(sw.has("swCell")).toBe(true);
      expect(sw.has("lastRow")).toBe(true);
      expect(sw.has("firstCol")).toBe(true);

      const se = activeConditionalTypes({ rowIndex: 2, colIndex: 2, totalRows, totalCols }, look);
      expect(se.has("seCell")).toBe(true);
      expect(se.has("lastRow")).toBe(true);
      expect(se.has("lastCol")).toBe(true);
    });

    it("exempts firstRow and lastRow from horizontal banding when firstRow/lastRow active", () => {
      const typesRow0 = activeConditionalTypes(
        { rowIndex: 0, colIndex: 1, totalRows: 5, totalCols: 3 },
        look,
      );
      expect(typesRow0.has("firstRow")).toBe(true);
      expect(typesRow0.has("band1Horz")).toBe(false);

      const typesRow1 = activeConditionalTypes(
        { rowIndex: 1, colIndex: 1, totalRows: 5, totalCols: 3 },
        look,
      );
      expect(typesRow1.has("band1Horz")).toBe(true);

      const typesRow2 = activeConditionalTypes(
        { rowIndex: 2, colIndex: 1, totalRows: 5, totalCols: 3 },
        look,
      );
      expect(typesRow2.has("band2Horz")).toBe(true);

      const typesRow4 = activeConditionalTypes(
        { rowIndex: 4, colIndex: 1, totalRows: 5, totalCols: 3 },
        look,
      );
      expect(typesRow4.has("lastRow")).toBe(true);
      expect(typesRow4.has("band1Horz")).toBe(false);
      expect(typesRow4.has("band2Horz")).toBe(false);
    });

    it("starts horizontal banding from row 0 when firstRow is false", () => {
      const lookNoHeader = { ...look, firstRow: false };
      const row0 = activeConditionalTypes(
        { rowIndex: 0, colIndex: 1, totalRows: 4, totalCols: 3 },
        lookNoHeader,
      );
      expect(row0.has("band1Horz")).toBe(true);
      expect(row0.has("firstRow")).toBe(false);

      const row1 = activeConditionalTypes(
        { rowIndex: 1, colIndex: 1, totalRows: 4, totalCols: 3 },
        lookNoHeader,
      );
      expect(row1.has("band2Horz")).toBe(true);
    });

    it("respects styleRowBandSize for multi-row bands", () => {
      // bandSize 2: rows 1,2 => band1Horz; rows 3,4 => band2Horz
      const row1 = activeConditionalTypes(
        { rowIndex: 1, colIndex: 1, totalRows: 6, totalCols: 3 },
        look,
        2,
      );
      const row2 = activeConditionalTypes(
        { rowIndex: 2, colIndex: 1, totalRows: 6, totalCols: 3 },
        look,
        2,
      );
      const row3 = activeConditionalTypes(
        { rowIndex: 3, colIndex: 1, totalRows: 6, totalCols: 3 },
        look,
        2,
      );
      expect(row1.has("band1Horz")).toBe(true);
      expect(row2.has("band1Horz")).toBe(true);
      expect(row3.has("band2Horz")).toBe(true);
    });
  });

  describe("resolveTableStyle inheritance", () => {
    it("merges conditionalFormats along the basedOn chain", () => {
      const styles: StylesOptions["tableStyles"] = [
        {
          id: "BaseStyle",
          conditionalFormats: [
            {
              type: "firstRow",
              cell: { shading: { fill: "111111", type: "clear" } },
              run: { bold: true },
            },
            {
              type: "band1Horz",
              cell: { shading: { fill: "EEEEEE", type: "clear" } },
            },
          ],
        },
        {
          id: "ChildStyle",
          basedOn: "BaseStyle",
          conditionalFormats: [
            {
              type: "firstRow",
              cell: { shading: { fill: "222222", type: "clear" } }, // overrides fill, keeps bold
            },
          ],
        },
      ];

      const resolved = resolveTableStyle(styles, "ChildStyle");
      expect(resolved).toBeDefined();
      const firstRow = resolved!.conditionalFormats.get("firstRow");
      expect(firstRow?.cell?.shading?.fill).toBe("222222");
      expect(firstRow?.run?.bold).toBe(true);
      expect(resolved!.conditionalFormats.get("band1Horz")?.cell?.shading?.fill).toBe("EEEEEE");
    });

    it("resolves built-in preset fallback when not in tableStyles", () => {
      const resolved = resolveTableStyle([], "grid-table");
      expect(resolved).toBeDefined();
      expect(resolved!.conditionalFormats.get("firstRow")?.cell?.shading?.fill).toBe("4472C4");
      expect(resolved!.conditionalFormats.get("band1Horz")?.cell?.shading?.fill).toBe("D9E2F3");
    });
  });

  describe("projectTable with tblLook conditional formatting", () => {
    const testStyles: StylesOptions = {
      tableStyles: [
        {
          id: "TestBanded",
          table: {
            borders: {
              top: { style: "single", size: 4, color: "000000" },
              bottom: { style: "single", size: 4, color: "000000" },
              insideHorizontal: { style: "single", size: 4, color: "CCCCCC" },
            },
          },
          conditionalFormats: [
            {
              type: "firstRow",
              cell: { shading: { fill: "305496", type: "clear" } },
              run: { bold: true, color: "FFFFFF" },
            },
            {
              type: "band1Horz",
              cell: { shading: { fill: "D9E1F2", type: "clear" } },
            },
            {
              type: "lastRow",
              cell: {
                shading: { fill: "F2F2F2", type: "clear" },
                borders: { top: { style: "double", size: 8, color: "305496" } },
              },
            },
            {
              type: "nwCell",
              cell: { shading: { fill: "1F3864", type: "clear" } },
            },
          ],
        },
      ],
    };

    const makeTable = (look?: TableOptions["tableLook"]): TableOptions => ({
      style: "TestBanded",
      tableLook: look,
      rows: [
        {
          tableHeader: true,
          cells: [
            { children: [{ paragraph: { children: ["H1"] } }] },
            { children: [{ paragraph: { children: ["H2"] } }] },
          ],
        },
        {
          cells: [
            { children: [{ paragraph: { children: ["B1"] } }] },
            { children: [{ paragraph: { children: ["B2"] } }] },
          ],
        },
        {
          cells: [
            { children: [{ paragraph: { children: ["B3"] } }] },
            { children: [{ paragraph: { children: ["B4"] } }] },
          ],
        },
        {
          cells: [
            { children: [{ paragraph: { children: ["T1"] } }] },
            { children: [{ paragraph: { children: ["T2"] } }] },
          ],
        },
      ],
    });

    it("projects firstRow, nwCell, band1Horz, and lastRow fills according to default tblLook", () => {
      const ctx = makeCtx(testStyles);
      const table = projectTable(
        makeTable({ firstRow: true, bandRow: true, lastRow: true, firstCol: true }),
        ctx,
      );

      // Row 0: Header row
      // Cell (0,0) is nwCell because both firstRow and firstCol are active
      expect(table.rows[0].cells[0].fill).toBe("1F3864");
      // Cell (0,1) is firstRow
      expect(table.rows[0].cells[1].fill).toBe("305496");

      // Row 1: First body row => band1Horz (offset 0 from header)
      expect(table.rows[1].cells[0].fill).toBe("D9E1F2");
      expect(table.rows[1].cells[1].fill).toBe("D9E1F2");

      // Row 2: Second body row => band2Horz (not defined in style => undefined)
      expect(table.rows[2].cells[0].fill).toBeUndefined();
      expect(table.rows[2].cells[1].fill).toBeUndefined();

      // Row 3: Total row => lastRow
      expect(table.rows[3].cells[0].fill).toBe("F2F2F2");
      expect(table.rows[3].cells[0].borders?.top?.style).toBe("double");
    });

    it("runs in header row inherit bold and white color from firstRow conditional formatting", () => {
      const ctx = makeCtx(testStyles);
      const table = projectTable(makeTable({ firstRow: true }), ctx);

      const headerCell = table.rows[0].cells[1];
      const para = headerCell.blocks[0];
      if (para.kind !== "paragraph" || !para.defaultTextStyle)
        throw new Error("expected paragraph with style");
      expect(para.defaultTextStyle.bold).toBe(true);
      expect(para.defaultTextStyle.color).toBe("FFFFFF");

      // Body row should NOT be bold or white
      const bodyCell = table.rows[1].cells[0];
      const bodyPara = bodyCell.blocks[0];
      if (bodyPara.kind !== "paragraph" || !bodyPara.defaultTextStyle)
        throw new Error("expected paragraph");
      expect(bodyPara.defaultTextStyle.bold).toBeUndefined();
    });

    it("live toggle: disabling firstRow shifts banding to row 0 and removes header styling", () => {
      const ctx = makeCtx(testStyles);
      const table = projectTable(makeTable({ firstRow: false, bandRow: true }), ctx);

      // Row 0 now gets band1Horz instead of firstRow
      expect(table.rows[0].cells[0].fill).toBe("D9E1F2");
      const para = table.rows[0].cells[0].blocks[0];
      if (para.kind !== "paragraph" || !para.defaultTextStyle)
        throw new Error("expected paragraph");
      expect(para.defaultTextStyle.bold).toBeUndefined();

      // Row 1 now gets band2Horz (undefined)
      expect(table.rows[1].cells[0].fill).toBeUndefined();
    });

    it("direct cell shading overrides conditional formatting fill", () => {
      const ctx = makeCtx(testStyles);
      const t = makeTable({ firstRow: true });
      // Direct shading on cell (0, 1)
      const row0 = t.rows?.[0];
      if (row0 && "cells" in row0 && Array.isArray(row0.cells)) {
        (row0.cells[1] as TableCellOptions).shading = { fill: "FFFF00", type: "clear" };
      }

      const table = projectTable(t, ctx);
      expect(table.rows[0].cells[1].fill).toBe("FFFF00");
    });
  });
});
