// @vitest-environment happy-dom
import {
  decodePassthroughData,
  docxExtensions,
  encodePassthroughData,
  generateDOCXSync,
  parseDOCXSync,
  type JSONContent,
} from "@docen/docx";
import { Editor, type Editor as EditorType } from "@docen/docx/core";
import { unzipSync } from "@office-open/core";
import { describe, expect, it, vi } from "vitest";

import { DocumentCommands } from "../extensions/commands";
import {
  createBlankExcelWorkbookBytes,
  downloadOleObject,
  getQuickTableBuildingBlocks,
  getQuickTableJson,
  type QuickTableId,
} from "../quick-tables";

function buildEditor(content?: JSONContent): EditorType {
  return new Editor({
    element: null,
    extensions: [...docxExtensions, DocumentCommands],
    content: content ?? {
      type: "doc",
      content: [{ type: "paragraph", content: [{ type: "text", text: "Target line" }] }],
    },
  });
}

describe("Quick Tables presets and generator", () => {
  const presets: QuickTableId[] = [
    "calendar1",
    "calendar2",
    "matrix",
    "tabularList",
    "doubleTable",
    "subheadings",
  ];

  for (const id of presets) {
    it(`generates valid table structure for preset '${id}'`, () => {
      const json = getQuickTableJson(id);
      expect(json).not.toBeNull();
      expect(json!.type).toBe("table");
      expect(Array.isArray(json!.content)).toBe(true);
      expect(json!.content!.length).toBeGreaterThan(0);

      // Verify each row contains table cells
      for (const row of json!.content!) {
        expect(row.type).toBe("tableRow");
        expect(Array.isArray(row.content)).toBe(true);
        expect(row.content!.length).toBeGreaterThan(0);
        for (const cell of row.content!) {
          expect(cell.type).toBe("tableCell");
          expect(Array.isArray(cell.content)).toBe(true);
        }
      }
    });
  }

  it("exports building block entries for Built-In Building Blocks Organizer", () => {
    const blocks = getQuickTableBuildingBlocks();
    expect(blocks).toHaveLength(6);
    for (const b of blocks) {
      expect(b.gallery).toBe("custTables");
      expect(b.category).toBe("Built-In");
      expect(b.content.content[0]?.type).toBe("table");
    }
  });

  it("generates a valid minimal OOXML Excel spreadsheet bundle", () => {
    const bytes = createBlankExcelWorkbookBytes();
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(0);

    const parts = unzipSync(bytes);
    expect(parts["[Content_Types].xml"]).toBeDefined();
    expect(parts["_rels/.rels"]).toBeDefined();
    expect(parts["xl/workbook.xml"]).toBeDefined();
    expect(parts["xl/_rels/workbook.xml.rels"]).toBeDefined();
    expect(parts["xl/worksheets/sheet1.xml"]).toBeDefined();

    const dec = new TextDecoder();
    expect(dec.decode(parts["[Content_Types].xml"])).toContain("spreadsheetml.sheet.main+xml");
    expect(dec.decode(parts["xl/workbook.xml"])).toContain('<sheet name="Sheet1"');
  });
});

describe("Editor commands: insert-quick-table & insert-excel", () => {
  it("executes 'insert-quick-table' command and replaces selection with table", () => {
    const editor = buildEditor();
    expect(editor.commands["insert-quick-table"]("calendar1")).toBe(true);

    let foundTable = false;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "table") {
        foundTable = true;
        expect(node.childCount).toBeGreaterThan(1);
        return false;
      }
      return true;
    });
    expect(foundTable).toBe(true);
  });

  it("executes 'insert-quick-table' with matrix preset", () => {
    const editor = buildEditor();
    expect(editor.commands["insert-quick-table"]("matrix")).toBe(true);

    let cellCount = 0;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tableCell") cellCount++;
    });
    // Matrix is 4 rows x 4 cols = 16 cells
    expect(cellCount).toBe(16);
  });

  it("inserts quick table via insert-building-block", () => {
    const editor = buildEditor();
    expect(editor.commands["insert-building-block"]("quick-table-doubleTable")).toBe(true);

    let found = false;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "table") {
        found = true;
        return false;
      }
      return true;
    });
    expect(found).toBe(true);
  });

  it("executes 'insert-excel' command and embeds minimal Excel workbook", () => {
    const editor = buildEditor();
    expect(editor.commands["insert-excel"]()).toBe(true);

    let foundInlinePassthrough = false;
    let passthroughData: string | undefined;

    editor.state.doc.descendants((node) => {
      if (node.type.name === "inlinePassthrough") {
        foundInlinePassthrough = true;
        passthroughData = node.attrs?.data;
        return false;
      }
      return true;
    });

    expect(foundInlinePassthrough).toBe(true);
    expect(passthroughData).toBeDefined();

    const decoded = decodePassthroughData<{
      object?: {
        width?: number;
        height?: number;
        embed?: {
          data?: Uint8Array;
          progId?: string;
          fileName?: string;
          relationshipType?: string;
        };
      };
    }>(passthroughData!);

    expect(decoded?.object).toBeDefined();
    expect(decoded!.object!.embed?.progId).toBe("Excel.Sheet.12");
    expect(decoded!.object!.embed?.fileName).toBe("Microsoft_Excel_Worksheet1.xlsx");
    expect(decoded!.object!.embed?.relationshipType).toBe("oleObject");

    const zipParts = unzipSync(decoded!.object!.embed!.data!);
    expect(zipParts["xl/workbook.xml"]).toBeDefined();
  });
});

describe("Editability Gen 1-3 round-trip for embedded Excel OLE object", () => {
  it("preserves embedded Excel workbook binary and <w:object> across 3 generations", () => {
    const origBytes = createBlankExcelWorkbookBytes();
    const docJson: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "inlinePassthrough",
              attrs: {
                data: encodePassthroughData({
                  object: {
                    width: 360,
                    height: 180,
                    embed: {
                      data: origBytes,
                      progId: "Excel.Sheet.12",
                      fileName: "Microsoft_Excel_Worksheet.xlsx",
                      relationshipType: "oleObject",
                    },
                  },
                }),
              },
            },
          ],
        },
      ],
    };

    // Gen 1: generate DOCX
    const gen1Docx = generateDOCXSync(docJson, { prepare: false }) as Uint8Array;
    const gen1Zip = unzipSync(gen1Docx);

    // Verify Gen 1 OOXML parts
    const dec = new TextDecoder();
    const docXml = dec.decode(gen1Zip["word/document.xml"]!);
    expect(docXml).toContain("w:object");
    expect(docXml).toContain('ProgID="Excel.Sheet.12"');

    const relsXml = dec.decode(gen1Zip["word/_rels/document.xml.rels"]!);
    expect(relsXml).toContain("oleObject");

    const embeddingEntry = Object.keys(gen1Zip).find((name) => name.startsWith("word/embeddings/"));
    expect(embeddingEntry).toBeDefined();
    expect(gen1Zip[embeddingEntry!]).toEqual(origBytes);

    // Gen 2: parse DOCX back to doc JSON
    const gen2Json = parseDOCXSync(gen1Docx);
    let gen2Object: any;
    for (const block of gen2Json.content ?? []) {
      for (const child of block.content ?? []) {
        if (child.type === "inlinePassthrough" && child.attrs?.data) {
          const decoded = decodePassthroughData<any>(child.attrs.data);
          if (decoded?.object) {
            gen2Object = decoded.object;
            break;
          }
        }
      }
    }
    expect(gen2Object).toBeDefined();
    expect(gen2Object.embed?.progId).toBe("Excel.Sheet.12");
    expect(new Uint8Array(gen2Object.embed?.data)).toEqual(origBytes);

    // Gen 3: re-generate DOCX from Gen 2 parsed JSON
    const gen3Docx = generateDOCXSync(gen2Json, { prepare: false }) as Uint8Array;
    const gen3Zip = unzipSync(gen3Docx);
    const gen3DocXml = dec.decode(gen3Zip["word/document.xml"]!);
    expect(gen3DocXml).toContain("w:object");
    expect(gen3DocXml).toContain('ProgID="Excel.Sheet.12"');

    const gen3Embedding = Object.keys(gen3Zip).find((name) => name.startsWith("word/embeddings/"));
    expect(gen3Embedding).toBeDefined();
    expect(gen3Zip[gen3Embedding!]).toEqual(origBytes);
  });
});

describe("downloadOleObject helper", () => {
  it("triggers browser download with proper mime and file name", () => {
    const bytes = new Uint8Array([1, 2, 3, 4]);
    const clickedSpy = vi.fn();
    const originalCreate = document.createElement.bind(document);

    vi.spyOn(document, "createElement").mockImplementation((tag: string) => {
      const el = originalCreate(tag);
      if (tag === "a") {
        el.click = clickedSpy;
      }
      return el;
    });

    downloadOleObject(bytes, "test_sheet.xlsx");
    expect(clickedSpy).toHaveBeenCalled();
  });
});
