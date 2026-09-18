import { generateDocumentSync } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { encodePassthroughData } from "../extensions/passthrough";
import {
  detectUnsupportedContent,
  parseDOCXSync,
  resolveDocument,
  type UnsupportedContentKind,
} from "../index";

/**
 * Item-14 detection oracle: real bytes/model → resolve → report. The detector
 * must flag exactly the uneditable container content (altChunk, subDoc,
 * SmartArt, OLE, raw/custom XML, content parts) and stay quiet for metadata
 * that has visible semantics of its own (comment range markers, bookmarks,
 * proofErr, symbol runs).
 */

function atom(type: "passthrough" | "inlinePassthrough", tag: string, value: unknown): JSONContent {
  return { type, attrs: { data: encodePassthroughData({ [tag]: value }) } };
}

function kinds(json: JSONContent): UnsupportedContentKind[] {
  return detectUnsupportedContent(json).items.map((i) => i.kind);
}

describe("detectUnsupportedContent", () => {
  it("reports block-level altChunk/rawXml/customXml and inline containers", () => {
    const json: JSONContent = {
      type: "doc",
      content: [
        atom("passthrough", "altChunk", { data: "PGgxPk<" }),
        atom("passthrough", "rawXml", "<w:p/>"),
        atom("passthrough", "customXml", { element: "CX", children: [] }),
        {
          type: "paragraph",
          content: [
            atom("inlinePassthrough", "smartArt", { nodes: [] }),
            atom("inlinePassthrough", "object", { shapeId: "_x0000_s1026" }),
            atom("inlinePassthrough", "subDoc", { data: "x" }),
            atom("inlinePassthrough", "rawXml", "<w:fldSimple/>"),
            atom("inlinePassthrough", "customXml", { element: "CX", children: [] }),
            atom("inlinePassthrough", "contentPart", { referenceId: "rId1" }),
            { type: "text", text: "editable" },
          ],
        },
      ],
    };
    const report = detectUnsupportedContent(json);
    expect(report.items).toEqual([
      { kind: "altChunk", count: 1 },
      { kind: "subDoc", count: 1 },
      { kind: "smartArt", count: 1 },
      { kind: "oleObject", count: 1 },
      { kind: "rawXml", count: 2 },
      { kind: "customXml", count: 2 },
      { kind: "contentPart", count: 1 },
    ]);
    expect(report.total).toBe(9);
  });

  it("stays quiet for marker metadata with visible semantics", () => {
    const json: JSONContent = {
      type: "doc",
      content: [
        { type: "paragraph", content: [atom("inlinePassthrough", "bookmarkStart", { id: 1 })] },
        { type: "paragraph", content: [atom("inlinePassthrough", "proofErr", "spellStart")] },
        {
          type: "paragraph",
          content: [atom("inlinePassthrough", "commentRangeStart", { id: 1 })],
        },
        { type: "paragraph", content: [atom("inlinePassthrough", "symbolRun", { char: "F0A7" })] },
        { type: "paragraph", content: [atom("passthrough", "bookmarkEnd", { id: 1 })] },
      ],
    };
    expect(detectUnsupportedContent(json)).toEqual({ items: [], total: 0 });
  });

  it("flags content that survives a real generate → parse byte trip", () => {
    const bytes = generateDocumentSync({
      sections: [
        {
          children: [
            { altChunk: { data: "PGgxPk<", contentType: "text/html", extension: "html" } },
          ],
        },
      ],
    });
    const json = parseDOCXSync(new Uint8Array(bytes as Buffer));
    const report = detectUnsupportedContent(json);
    expect(report.total).toBe(1);
    expect(report.items[0]).toEqual({ kind: "altChunk", count: 1 });
  });

  it("flags an OLE object resolved from the model", () => {
    const json = resolveDocument({
      sections: [
        {
          children: [
            { paragraph: { children: [{ object: { shapeId: "_x0000_s1026" } } as never] } },
          ],
        },
      ],
    });
    expect(kinds(json)).toContain("oleObject");
  });
});
