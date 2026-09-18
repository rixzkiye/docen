import { docxExtensions } from "@docen/docx";
import { Editor } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import {
  demoteHeading,
  demoteHeadingAtCaret,
  filterOutlineByLevel,
  findHeadingSectionRange,
  moveBlockDown,
  moveBlockUp,
  moveHeadingSection,
  promoteHeading,
  promoteHeadingAtCaret,
  selectSimilarFormatting,
  setHeadingLevel,
} from "./outline";

describe("Outline View Commands (D4b)", () => {
  it("filters outline tree items by level", () => {
    const items = [
      {
        id: "1",
        title: "Heading 1",
        level: 1,
        children: [
          { id: "2", title: "Heading 2", level: 2 },
          { id: "3", title: "Heading 3", level: 3 },
        ],
      },
      { id: "4", title: "Another Heading 1", level: 1 },
    ];

    const level1Only = filterOutlineByLevel(items, 1);
    expect(level1Only.length).toBe(2);
    expect(level1Only[0]?.children).toBeUndefined();

    const upToLevel2 = filterOutlineByLevel(items, 2);
    expect(upToLevel2[0]?.children?.length).toBe(1);
    expect(upToLevel2[0]?.children?.[0]?.title).toBe("Heading 2");
  });

  it("promotes and demotes headings at caret", () => {
    const editor = new Editor({
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { style: "Heading2" },
            content: [{ type: "text", text: "Section A" }],
          },
        ],
      },
    });

    editor.commands.setTextSelection(3);
    const promoted = promoteHeadingAtCaret(editor);
    expect(promoted).toBe(true);
    expect(editor.getJSON().content?.[0]?.attrs?.style).toBe("Heading1");

    const demoted = demoteHeadingAtCaret(editor);
    expect(demoted).toBe(true);
    expect(editor.getJSON().content?.[0]?.attrs?.style).toBe("Heading2");
  });

  it("moves blocks up and down", () => {
    const editor = new Editor({
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          { type: "paragraph", content: [{ type: "text", text: "Para 1" }] },
          { type: "paragraph", content: [{ type: "text", text: "Para 2" }] },
        ],
      },
    });

    // Caret at Para 1
    editor.commands.setTextSelection(2);
    expect(moveBlockUp(editor)).toBe(false); // cannot move top block up
    expect(moveBlockDown(editor)).toBe(true);

    const json = editor.getJSON() as Record<string, any>;
    expect(json.content?.[0]?.content?.[0]?.text).toBe("Para 2");
    expect(json.content?.[1]?.content?.[0]?.text).toBe("Para 1");
  });

  it("selects similar formatting across paragraphs", () => {
    const editor = new Editor({
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              { type: "text", marks: [{ type: "bold" }], text: "Bold One" },
              { type: "text", text: " Plain" },
            ],
          },
          {
            type: "paragraph",
            content: [{ type: "text", marks: [{ type: "bold" }], text: "Bold Two" }],
          },
        ],
      },
    });

    editor.commands.setTextSelection(2); // inside "Bold One"
    const matched = selectSimilarFormatting(editor);
    expect(matched).toBe(true);
  });

  it("finds heading section ranges correctly", () => {
    const editor = new Editor({
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { heading: "Heading1" },
            content: [{ type: "text", text: "H1" }],
          },
          { type: "paragraph", content: [{ type: "text", text: "Body 1" }] },
          {
            type: "paragraph",
            attrs: { heading: "Heading2" },
            content: [{ type: "text", text: "H2" }],
          },
          { type: "paragraph", content: [{ type: "text", text: "Body 2" }] },
          {
            type: "paragraph",
            attrs: { heading: "Heading1" },
            content: [{ type: "text", text: "Next H1" }],
          },
        ],
      },
    });

    const doc = editor.state.doc;
    // Section for H1 at index 0 should span indices 0..3 (H1, Body 1, H2, Body 2)
    const s0 = findHeadingSectionRange(doc, 0);
    expect(s0.startIndex).toBe(0);
    expect(s0.endIndex).toBe(4);

    // Section for H2 at index 2 should span indices 2..3 (H2, Body 2)
    const s2 = findHeadingSectionRange(doc, 2);
    expect(s2.startIndex).toBe(2);
    expect(s2.endIndex).toBe(4);

    // Section for Next H1 at index 4 spans index 4..5
    const s4 = findHeadingSectionRange(doc, 4);
    expect(s4.startIndex).toBe(4);
    expect(s4.endIndex).toBe(5);
  });

  it("moves entire heading sections before or after", () => {
    const editor = new Editor({
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { heading: "Heading1" },
            content: [{ type: "text", text: "Sec A" }],
          },
          { type: "paragraph", content: [{ type: "text", text: "Body A" }] },
          {
            type: "paragraph",
            attrs: { heading: "Heading1" },
            content: [{ type: "text", text: "Sec B" }],
          },
          { type: "paragraph", content: [{ type: "text", text: "Body B" }] },
        ],
      },
    });

    // Move Sec A after Sec B
    const moved = moveHeadingSection(editor, 0, 2, "after");
    expect(moved).toBe(true);

    expect(editor.state.doc.child(0).textContent).toBe("Sec B");
    expect(editor.state.doc.child(1).textContent).toBe("Body B");
    expect(editor.state.doc.child(2).textContent).toBe("Sec A");
    expect(editor.state.doc.child(3).textContent).toBe("Body A");
  });

  it("supports promoteHeading and demoteHeading by blockIndex", () => {
    const editor = new Editor({
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { heading: "Heading2" },
            content: [{ type: "text", text: "Title" }],
          },
        ],
      },
    });

    promoteHeading(editor, 0);
    expect(editor.getJSON().content?.[0]?.attrs?.heading).toBe("Heading1");

    demoteHeading(editor, 0);
    expect(editor.getJSON().content?.[0]?.attrs?.heading).toBe("Heading2");
  });

  it("supports setHeadingLevel", () => {
    const editor = new Editor({
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [{ type: "paragraph", content: [{ type: "text", text: "Text" }] }],
      },
    });

    setHeadingLevel(editor, 3, 0);
    expect(editor.getJSON().content?.[0]?.attrs?.heading).toBe("Heading3");

    setHeadingLevel(editor, 0, 0);
    expect(editor.getJSON().content?.[0]?.attrs?.heading).toBeNull();
  });
});
