// @vitest-environment happy-dom
/**
 * Section-targeted sectPr writes — the fixed vertical ruler's margin drags
 * must land on the section at the pane top (never the caret's) and preserve
 * every margin side the drag does not touch.
 */
import { docxExtensions } from "@docen/docx";
import { Editor } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { SectionCommands } from "./sections";

function setup() {
  const editor = new Editor({
    element: null,
    extensions: docxExtensions,
    content: {
      type: "doc",
      attrs: {
        sectionProperties: {
          pageMargin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        },
      },
      content: [
        {
          type: "paragraph",
          attrs: { sectionProperties: { pageMargin: { top: 720, bottom: 720 } } },
          content: [{ type: "text", text: "First section" }],
        },
        { type: "paragraph", content: [{ type: "text", text: "Second section" }] },
      ],
    },
  });
  const commands = new SectionCommands({
    editor: () => editor,
    bridge: () => undefined,
    element: () => document.body,
    flow: () => undefined,
  });
  return { editor, commands };
}

describe("SectionCommands — section-indexed sectPr writes (vertical ruler)", () => {
  it("maps a section index to its own sectPr position", () => {
    const { commands } = setup();
    // The first section-carrying paragraph closes section 0; the final section
    // is the body-level sectPr.
    expect(commands.sectionSectPrPosAt(0)).toBe(0);
    expect(commands.sectionSectPrPosAt(1)).toBeNull();
    expect(commands.sectionSectPrPosAt(2)).toBeNull();
  });

  it("writes a margin into the dragged section, not the caret's", () => {
    const { editor, commands } = setup();
    // The caret sits in the SECOND section (body sectPr)…
    editor.commands.setTextSelection(editor.state.doc.content.size - 1);
    expect(commands.sectionSectPrPos()).toBeNull();

    // …but the ruler drag opened on section 0.
    commands.setSectionMargin(0, "top", 2160);
    const first = editor.state.doc.child(0).attrs.sectionProperties;
    expect(first.pageMargin.top).toBe(2160);
    // Untouched sides survive the merge.
    expect(first.pageMargin.bottom).toBe(720);

    commands.setSectionMargin(1, "bottom", 1080);
    const body = editor.state.doc.attrs.sectionProperties;
    expect(body.pageMargin.bottom).toBe(1080);
    expect(body.pageMargin.top).toBe(1440);
    expect(body.pageMargin.left).toBe(1440);
    // The other section stayed where its own drag left it.
    expect(editor.state.doc.child(0).attrs.sectionProperties.pageMargin.top).toBe(2160);
  });

  it("clamps a drag past the page edge to a non-negative margin", () => {
    const { editor, commands } = setup();
    commands.setSectionMargin(0, "top", -300);
    expect(editor.state.doc.child(0).attrs.sectionProperties.pageMargin.top).toBe(0);
  });
});
