import { docxExtensions } from "@docen/docx";
import { Editor } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import {
  demoteHeadingAtCaret,
  filterOutlineByLevel,
  moveBlockDown,
  moveBlockUp,
  promoteHeadingAtCaret,
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
});
