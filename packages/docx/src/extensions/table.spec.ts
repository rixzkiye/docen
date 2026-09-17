import { unzipSync } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import { describe, it, expect } from "vitest";

import { compileDocument, docxExtensions, generateDOCXSync, resolveDocument } from "../index";
import { parseDocx, renderDocx } from "./table";

// float round-trips byte-faithful through the table's renderDocx/parseDocx
// attrs passthrough (SKIP_KEYS drops only rows/columnWidthsRevision) — this is
// what lets a floating table keep its anchor across DOCX→JSON→DOCX even when
// the v1 renderer degrades (absolute anchor, center) and emits no CSS.
describe("table float round-trip", () => {
  it("preserves a full float anchor verbatim", () => {
    const float = {
      horizontalAnchor: "text",
      verticalAnchor: "text",
      absoluteHorizontalPosition: 720,
      absoluteVerticalPosition: 360,
      relativeHorizontalPosition: "right",
      relativeVerticalPosition: "top",
      leftFromText: 180,
      rightFromText: 180,
      topFromText: 200,
      bottomFromText: 200,
      overlap: "neverOverlap",
    };
    const back = parseDocx({ rows: [], ...renderDocx({ type: "table", attrs: { float } }) });
    expect(back.float).toEqual(float);
  });
});

describe("table rowSpan resolve", () => {
  /**
   * TableCellPropertiesOptions.rowSpan is office-open's writer-side shorthand:
   * the writer expands it into a vMerge restart plus continuation cells, but
   * the parser only ever emits the expanded form. resolve must materialize
   * the same expansion — carrying rowSpan through silently dropped the merge.
   */
  it("materializes rowSpan as vMerge restart + continuation cells", () => {
    const doc: DocumentOptions = {
      sections: [
        {
          children: [
            {
              table: {
                rows: [
                  {
                    cells: [
                      { children: [{ paragraph: { text: "r1c1" } }], rowSpan: 2 },
                      { children: [{ paragraph: { text: "r1c2" } }] },
                    ],
                  },
                  { cells: [{ children: [{ paragraph: { text: "r2c2" } }] }] },
                ],
              },
            },
          ],
        },
      ],
    };
    const json = resolveDocument(doc, docxExtensions);
    const table = json.content?.find((n) => n.type === "table");
    const rows = (table?.content ?? []) as {
      content?: { attrs?: Record<string, unknown> }[];
    }[];
    expect(rows[0]?.content?.[0]?.attrs?.verticalMerge).toBe("restart");
    expect(rows[0]?.content?.[0]?.attrs?.rowSpan).toBeUndefined();
    // The continuation lands at the same grid column in the next row.
    expect(rows[1]?.content?.[0]?.attrs?.verticalMerge).toBe("continue");
    expect(rows[1]?.content).toHaveLength(2);

    const compiled = compileDocument(json, docxExtensions);
    const cell = compiled.sections[0]?.children[0] as {
      table: { rows: { cells: { verticalMerge?: string }[] }[] };
    };
    expect(cell.table.rows[0]?.cells[0]?.verticalMerge).toBe("restart");
    expect(cell.table.rows[1]?.cells[0]?.verticalMerge).toBe("continue");
    expect(cell.table.rows[1]?.cells[1]).toMatchObject({ children: [{ paragraph: "r2c2" }] });

    const xml = new TextDecoder().decode(
      unzipSync(generateDOCXSync(json, { prepare: false }) as Uint8Array)["word/document.xml"],
    );
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toMatch(/<w:vMerge(\s+w:val="continue")?\/>/);
  });

  it("places the continuation after a leading column span", () => {
    const doc: DocumentOptions = {
      sections: [
        {
          children: [
            {
              table: {
                rows: [
                  {
                    cells: [
                      {
                        children: [{ paragraph: { text: "wide" } }],
                        columnSpan: 2,
                      },
                      { children: [{ paragraph: { text: "span" } }], rowSpan: 2 },
                    ],
                  },
                  { cells: [{ children: [{ paragraph: { text: "below wide" } }], columnSpan: 2 }] },
                ],
              },
            },
          ],
        },
      ],
    };
    const json = resolveDocument(doc, docxExtensions);
    const rows = (json.content?.find((n) => n.type === "table")?.content ?? []) as {
      content?: { attrs?: Record<string, unknown> }[];
    }[];
    // The continuation belongs to column 2, after the 2-wide cell.
    expect(rows[0]?.content?.[1]?.attrs?.verticalMerge).toBe("restart");
    expect(rows[1]?.content?.[1]?.attrs?.verticalMerge).toBe("continue");
  });
});
