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

import { TocCommands } from "../extensions/toc";
import { collectBookmarkPages, collectFieldPages, collectTocTargetPages } from "./field-pages";

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

  it("maps bookmark names to their start page for PAGEREF", () => {
    const editor = build({
      type: "doc",
      content: [
        { type: "paragraph", content: [field(" PAGEREF _Ref1 \\h ", "9")] },
        {
          type: "paragraph",
          content: [
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ bookmarkStart: { id: 1, name: "_Ref1" } }) },
            },
            { type: "text", text: "target" },
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ bookmarkEnd: { id: 1 } }) },
            },
          ],
        },
      ],
    });
    let bookmarkPos = 0;
    editor.state.doc.descendants((node, at) => {
      if (node.type.name === "inlinePassthrough" && node.attrs?.data?.includes("bookmarkStart")) {
        bookmarkPos = at;
      }
      return true;
    });
    const view = {
      sectionOfPage: [0, 0, 1],
      pageOffsets: [0, 0],
      physicalPageOf: (pos: number) => (pos < bookmarkPos ? 0 : 2),
    };
    const bookmarks = collectBookmarkPages(editor.state.doc, view);
    expect([...bookmarks.entries()]).toEqual([["_Ref1", 3]]);
    const pages = collectFieldPages(editor.state.doc, view);
    const bytes = generateDOCXSync(editor.getJSON(), {
      prepare: false,
      fields: {
        pageCount: 3,
        pageOf: ({ index, bookmark }) =>
          bookmark != null ? bookmarks.get(bookmark) : pages.get(index),
      },
    }) as Uint8Array;
    // The PAGEREF shows the target's page (3), not the field's page (1).
    const xml = documentXml(bytes);
    expect(xml).toContain(">3<");
    expect(xml).not.toContain(">9<");
  });
});

describe("collectTocTargetPages", () => {
  it("maps heading/caption candidates in document order and feeds TOC caches", () => {
    const editor = build({
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { heading: "Heading1" },
          content: [{ type: "text", text: "One" }],
        },
        { type: "paragraph", content: [{ type: "text", text: "spacer" }] },
        {
          type: "paragraph",
          attrs: { heading: "Heading2" },
          content: [{ type: "text", text: "Two" }],
        },
        {
          type: "paragraph",
          attrs: { style: "Caption" },
          content: [
            { type: "text", text: "Figure " },
            {
              type: "inlinePassthrough",
              attrs: {
                data: JSON.stringify({
                  simpleField: { instruction: " SEQ Figure \\* ARABIC ", cachedValue: "1" },
                }),
              },
            },
            { type: "text", text: ": One" },
          ],
        },
        {
          type: "tocField",
          attrs: { options: { headingStyleRange: "1-3" } },
          content: [{ type: "paragraph" }],
        },
        {
          type: "tocField",
          attrs: { options: { captionLabelIncludingNumbers: "Figure" } },
          content: [{ type: "paragraph" }],
        },
      ],
    });
    let secondHeading = 0;
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === "paragraph" && node.attrs.heading === "Heading2") secondHeading = pos;
      return true;
    });
    const view = {
      sectionOfPage: [0, 0, 0],
      pageOffsets: [0],
      physicalPageOf: (pos: number) => (pos < secondHeading ? 0 : 2),
    };
    const toc = collectTocTargetPages(editor.state.doc, view);
    // The Caption paragraph is a heading candidate too (any styled paragraph
    // is) — the fill filters it out by level, but the index sequence includes it.
    expect([...toc.headingPages.entries()]).toEqual([
      [0, 1],
      [1, 3],
      [2, 3],
    ]);
    expect([...toc.captionPages.entries()]).toEqual([[0, 3]]);
    const bytes = generateDOCXSync(editor.getJSON(), {
      prepare: false,
      fields: {
        pageCount: 3,
        tocPageOf: ({ index, kind }) =>
          kind === "heading" ? toc.headingPages.get(index) : toc.captionPages.get(index),
      },
    }) as Uint8Array;
    const xml = documentXml(bytes);
    // Both empty-cache TOCs fill with cached numbers from the live map.
    const tocSlice = xml.slice(xml.indexOf("<w:sdt>"), xml.indexOf("</w:sdt>"));
    expect(tocSlice).toContain(">One<");
    expect(tocSlice).toContain(">3<");
    // The table of figures' `\c` entry carries the caption's page too.
    const tofIndex = xml.lastIndexOf("<w:sdt>");
    const tofSlice = xml.slice(tofIndex, xml.indexOf("</w:sdt>", tofIndex));
    expect(tofSlice).toContain("Figure 1: One");
    expect(tofSlice).toContain('\\c "Figure"');
    expect(tofSlice).toContain(">3<");
    editor.destroy();
  });
});

describe("editor TOC commands compile to Word switches", () => {
  it("emits \\t, \\b and \\c (not the dialog-only option keys) into the package", () => {
    const seed = (data: object) => ({
      type: "inlinePassthrough",
      attrs: { data: JSON.stringify(data) },
    });
    // The TOC commands live in their own extension; this editor needs them.
    const editor = new Editor({
      element: null,
      extensions: [
        Document,
        Paragraph,
        TextNode.create({ name: "text", group: "inline" }),
        Tab,
        Link,
        TocField,
        InlinePassthrough,
        TocCommands,
      ],
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { heading: "Heading1" },
            content: [{ type: "text", text: "Intro" }],
          },
          {
            type: "paragraph",
            content: [seed({ bookmarkStart: { id: 7, name: "Scope" } })],
          },
          {
            type: "paragraph",
            attrs: { style: "SpecialTitle" },
            content: [{ type: "text", text: "Scoped Custom" }],
          },
          {
            type: "paragraph",
            content: [seed({ bookmarkEnd: { id: 7 } })],
          },
          {
            type: "paragraph",
            attrs: { style: "Caption" },
            content: [
              { type: "text", text: "Figure " },
              seed({ simpleField: { instruction: " SEQ Figure \\* ARABIC ", cachedValue: "1" } }),
              { type: "text", text: ": chart" },
            ],
          },
        ],
      },
    });
    for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
    editor.commands.setTextSelection(1);
    expect(
      editor.commands.toc(() => 1, undefined, { styles: "SpecialTitle,1", bookmark: "Scope" }),
    ).toBe(true);
    expect(editor.commands["table-of-figures"](() => 1, undefined, "Figure")).toBe(true);

    const bytes = generateDOCXSync(editor.getJSON(), { prepare: false }) as Uint8Array;
    const xml = documentXml(bytes);
    expect(xml).toContain(`\\t "SpecialTitle,1"`);
    expect(xml).toContain(`\\b "Scope"`);
    expect(xml).toContain(`\\c "Figure"`);
    // The scoped TOC cached entry only covers the bookmark's custom-style
    // paragraph; the figure table lists the caption.
    expect(xml).toContain("Scoped Custom");
    expect(xml).toContain("Figure 1: chart");
    editor.destroy();
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
