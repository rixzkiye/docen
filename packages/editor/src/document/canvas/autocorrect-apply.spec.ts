// @vitest-environment node
// The bridge's fix application split out and exercised against a real Tiptap
// document — ordinal superscripts and hyperlink markings are observable only
// once the marks land in the schema-backed doc, so this spec is the bridge's
// transaction-level oracle (the bridge itself is DOM-bound and has no spec).
import { Document, Paragraph } from "@docen/docx";
import { Editor, Mark, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { applyAutocorrect, autocorrectOf, DEFAULT_AUTOCORRECT, hyperlinkFix } from "./autocorrect";

// The plain text node + minimal marks the engine addresses by name.
const Text = TextNode.create({ name: "text", group: "inline" });
const Superscript = Mark.create({ name: "superscript" });
const Link = Mark.create({
  name: "link",
  addAttributes: () => ({ href: { default: null } }),
});
const TextStyle = Mark.create({
  name: "textStyle",
  addAttributes: () => ({ style: { default: null } }),
});

const build = (text: string): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Superscript, Link, TextStyle],
    content: {
      type: "doc",
      content: [{ type: "paragraph", content: text ? [{ type: "text", text }] : undefined }],
    },
  });
  // Caret at the paragraph's end (paragraph content starts at 1).
  editor.commands.setTextSelection(1 + text.length);
  return editor;
};

/** The paragraph's runs — the independent oracle (text + mark names). */
const runsOf = (editor: EditorType): Array<{ text: string; marks: string[] }> => {
  const runs: Array<{ text: string; marks: string[] }> = [];
  editor.state.doc.firstChild?.forEach((node) => {
    runs.push({ text: node.text ?? "", marks: node.marks.map((mark) => mark.type.name) });
  });
  return runs;
};

describe("applyAutocorrect", () => {
  it("replaces the partial suffix and superscripts the full one", () => {
    const editor = build("1s");
    const { from, to } = editor.state.selection;
    const fix = autocorrectOf("t", "1s", { ...DEFAULT_AUTOCORRECT, ordinalSuperscript: true });
    expect(fix).not.toBeNull();
    if (!fix) return;
    editor.view.dispatch(applyAutocorrect(editor.state.tr, from - fix.back, to, fix));
    expect(runsOf(editor)).toEqual([
      { text: "1", marks: [] },
      { text: "st", marks: ["superscript"] },
    ]);
  });

  it("linkifies the URL behind the typed space", () => {
    const url = "https://example.com";
    const editor = build(url);
    const { from, to } = editor.state.selection;
    const fix = autocorrectOf(" ", url);
    expect(fix).not.toBeNull();
    if (!fix) return;
    editor.view.dispatch(applyAutocorrect(editor.state.tr, from, to, fix));
    const runs = runsOf(editor);
    expect(runs).toHaveLength(2);
    expect(runs[0]!.text).toBe(url);
    expect(runs[1]).toEqual({ text: " ", marks: [] });
    expect(runs[0]!.marks).toEqual(expect.arrayContaining(["link", "textStyle"]));
    const link = editor.state.doc.firstChild?.firstChild?.marks.find(
      (mark) => mark.type.name === "link",
    );
    expect(link?.attrs.href).toBe(url);
  });

  it("stamps link marks over a URL with no text change (the Enter leg)", () => {
    const url = "www.example.com";
    const editor = build(url);
    const { from, to } = editor.state.selection;
    const fix = hyperlinkFix(url);
    expect(fix).not.toBeNull();
    if (!fix) return;
    editor.view.dispatch(applyAutocorrect(editor.state.tr, from, to, fix));
    expect(runsOf(editor)).toEqual([
      { text: url, marks: expect.arrayContaining(["link", "textStyle"]) },
    ]);
  });
});
