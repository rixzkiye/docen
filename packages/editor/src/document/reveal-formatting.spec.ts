// @vitest-environment happy-dom
// formattingInfoOf — the snapshot the host feeds the Reveal Formatting pane and
// captures as the compare-to-selection reference.
import { docxExtensions } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { formattingInfoOf } from "./reveal-formatting";

const content = {
  type: "doc",
  attrs: {
    sectionProperties: {
      pageSize: { width: 11906, height: 16838, orientation: "landscape" },
      pageMargin: { top: 1440, left: 1440 },
    },
  },
  content: [
    { type: "paragraph", content: [{ type: "text", text: "Alpha plain" }] },
    {
      type: "paragraph",
      content: [{ type: "text", marks: [{ type: "bold" }], text: "Beta bold" }],
    },
  ],
};

function makeEditor(): Editor {
  return new Editor({ element: null, extensions: docxExtensions, content: content as never });
}

function posIn(editor: Editor, text: string): number {
  let pos = -1;
  editor.state.doc.descendants((node, p) => {
    if (node.isText && node.text === text) pos = p + 1;
  });
  return pos;
}

describe("formattingInfoOf", () => {
  it("reads font marks and the live section geometry", () => {
    const editor = makeEditor();
    editor.commands.setTextSelection(posIn(editor, "Alpha plain"));
    const plain = formattingInfoOf(editor);
    expect(plain.font?.family).toBe("Calibri");
    expect(plain.font?.bold).toBe(false);
    expect(plain.section?.orientation).toBe("Landscape");
    expect(plain.section?.margins).toBe('Top: 1.0", Left: 1.0"');
    expect(plain.section?.paperSize).toBe('8.3" × 11.7"');

    editor.commands.setTextSelection(posIn(editor, "Beta bold"));
    expect(formattingInfoOf(editor).font?.bold).toBe(true);
  });
});
