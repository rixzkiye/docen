import {
  Document,
  generateDOCXSync,
  InlinePassthrough,
  Link,
  Paragraph,
  Tab,
  TocField,
} from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vitest";

import { collectFieldPages } from "./field-pages";

/**
 * Item 17's editor side: the save path hands the builder a field-index → page
 * map derived from the live canvas pagination, so PAGE/NUMPAGES caches match
 * the document the user is about to save. The index contract (document order
 * of simple/complex field atoms, form fields and unmapped atoms included in
 * the count) must line up with the docx pass — this spec drives both legs end
 * to end.
 */

const Text = TextNode.create({ name: "text", group: "inline" });

const field = (instruction: string, cached: string): Record<string, unknown> => ({
  type: "inlinePassthrough",
  attrs: { data: JSON.stringify({ simpleField: { instruction, cachedValue: cached } }) },
});

const build = (doc: Record<string, unknown>): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Tab, Link, TocField, InlinePassthrough],
    content: doc,
  });
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

const documentXml = (bytes: Uint8Array): string =>
  new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]);

describe("collectFieldPages", () => {
  it("numbers field atoms in document order and applies section offsets", () => {
    const editor = build({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [field(" PAGE ", "9")],
        },
        { type: "paragraph", content: [{ type: "text", text: "middle" }] },
        {
          type: "paragraph",
          content: [field(" PAGE ", "9"), { type: "text", text: " / " }, field(" NUMPAGES ", "9")],
        },
      ],
    });
    // Physical pages: first field on page 0, the pair on page 2; the second
    // section restarts at 1 (offset -2 on its pages).
    const pages = collectFieldPages(editor.state.doc, {
      sectionOfPage: [0, 0, 1],
      pageOffsets: [0, -2],
      physicalPageOf: (pos) => (pos < 40 ? 0 : 2),
    });
    expect([...pages.entries()]).toEqual([
      [0, 1],
      [1, 1],
      [2, 1],
    ]);
  });

  it("skips form fields and unmapped positions without shifting later keys", () => {
    const editor = build({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ formField: { name: "ff" } }) },
            },
            field(" PAGE ", "9"),
          ],
        },
      ],
    });
    const pages = collectFieldPages(editor.state.doc, {
      sectionOfPage: [0],
      pageOffsets: [0],
      physicalPageOf: (pos) => (pos > 100 ? null : 0),
    });
    expect([...pages.entries()]).toEqual([[0, 1]]);
  });
});

describe("save-path page context feeds the builder", () => {
  it("writes the live page numbers into PAGE/NUMPAGES caches", () => {
    const editor = build({
      type: "doc",
      content: [
        { type: "paragraph", content: [field(" PAGE ", "9")] },
        { type: "paragraph", content: [{ type: "text", text: "spacer" }] },
        { type: "paragraph", content: [field(" PAGE ", "9")] },
        { type: "paragraph", content: [field(" NUMPAGES ", "9")] },
      ],
    });
    let pos = 0;
    const fieldPositions: number[] = [];
    editor.state.doc.descendants((node, at) => {
      if (node.type.name === "inlinePassthrough") {
        fieldPositions.push(at);
        pos = at;
      }
      return true;
    });
    // First field physical page 0, the rest physical page 2 (0-based) — three
    // pages total, one section (no restart).
    const split = fieldPositions[1]!;
    const pages = collectFieldPages(editor.state.doc, {
      sectionOfPage: [0, 0, 0],
      pageOffsets: [0],
      physicalPageOf: (at) => (at < split ? 0 : 2),
    });
    expect(pos).toBeGreaterThan(split);
    const bytes = generateDOCXSync(editor.getJSON(), {
      prepare: false,
      fields: { pageCount: 3, pageOf: ({ index }) => pages.get(index) },
    }) as Uint8Array;
    const xml = documentXml(bytes);
    expect(xml).toContain(">1<");
    expect(xml).toContain(">3<");
    expect(xml).not.toContain(">9<");
  });
});
