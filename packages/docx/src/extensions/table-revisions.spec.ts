import type { DocumentOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";
import type { ProjectContext } from "../layout/project/context";
import { projectTable } from "../layout/project/table";

describe("W9.2 Table Property Revisions (tblPrChange, trPrChange, tcPrChange)", () => {
  it("round-trips tblPrChange, trPrChange, and tcPrChange through resolve and compile", () => {
    const doc: DocumentOptions = {
      sections: [
        {
          children: [
            {
              table: {
                revision: { id: 1, author: "Reviewer", date: "2026-09-18T08:00:00Z" },
                rows: [
                  {
                    revision: { id: 2, author: "Reviewer", date: "2026-09-18T08:00:00Z" },
                    cells: [
                      {
                        revision: { id: 3, author: "Reviewer", date: "2026-09-18T08:00:00Z" },
                        children: [{ paragraph: { text: "Cell content" } }],
                      },
                    ],
                  },
                ],
              },
            },
          ],
        },
      ],
    };

    const resolved = resolveDocument(doc, docxExtensions);
    const tableNode = resolved.content?.[0];
    expect(tableNode?.type).toBe("table");
    expect(tableNode?.attrs?.tblPrChange).toMatchObject({ id: 1, author: "Reviewer" });

    const rowNode = tableNode?.content?.[0];
    expect(rowNode?.type).toBe("tableRow");
    expect(rowNode?.attrs?.trPrChange).toMatchObject({ id: 2, author: "Reviewer" });

    const cellNode = rowNode?.content?.[0];
    expect(cellNode?.type).toBe("tableCell");
    expect(cellNode?.attrs?.tcPrChange).toMatchObject({ id: 3, author: "Reviewer" });

    const compiled = compileDocument(resolved, docxExtensions);
    const compiledTable = (compiled.sections[0]!.children[0] as { table: any }).table;
    expect(compiledTable.revision?.id).toBe(1);
    expect(compiledTable.rows[0].revision?.id).toBe(2);
    expect(compiledTable.rows[0].cells[0].revision?.id).toBe(3);
  });

  it("projects revision balloon anchors on table, row, and cell revisions when balloons enabled", () => {
    const ctx: ProjectContext = {
      styles: undefined,
      characterStyles: new Map(),
      numberings: { byNumId: new Map(), abstractNumberings: new Map() } as any,
      listCounters: new Map(),
      openComments: new Set(),
      footnoteOrdinals: new Map(),
      endnoteOrdinals: new Map(),
      revisionAuthorColors: new Map(),
      stateful: { hit: false },
      markup: { view: "all", balloons: "all" },
    };

    const tableOptions = {
      tblPrChange: { id: 101, author: "TableEditor", date: "2026-09-18T08:00:00Z" },
      rows: [
        {
          trPrChange: { id: 102, author: "RowEditor", date: "2026-09-18T08:00:00Z" },
          cells: [
            {
              tcPrChange: { id: 103, author: "CellEditor", date: "2026-09-18T08:00:00Z" },
              children: [{ paragraph: { text: "Hello Table" } }],
            },
          ],
        },
      ],
    };

    const layout = projectTable(tableOptions as any, ctx);
    expect(layout.rows).toHaveLength(1);
    const firstCell = layout.rows[0]?.cells[0];
    const firstBlock = firstCell?.blocks[0] as {
      kind: string;
      balloons?: { id: number; kind: string; label: string }[];
    };

    expect(firstBlock.kind).toBe("paragraph");
    expect(firstBlock.balloons).toBeDefined();
    const balloonIds = firstBlock.balloons?.map((b) => b.id);
    expect(balloonIds).toContain(101); // Table revision balloon
    expect(balloonIds).toContain(102); // Row revision balloon
    expect(balloonIds).toContain(103); // Cell revision balloon
  });
});
