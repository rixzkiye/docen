// @vitest-environment happy-dom
// Text Effects end-to-end on the editor side: the gallery presets, the
// apply/clear commands on a real editor, the custom-dialog patch, and the
// DOCX round-trip of the w14 raw XML.
import { docxExtensions, generateDOCXSync, parseDOCXSync, type JSONContent } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { DocumentCommands, stampStyleRunPatches } from "./extensions/commands";
import { applyEffectsPatch, applyPreset, presetXml, textEffectThemeXml } from "./text-effects";

function makeEditor(text = "Effect"): Editor {
  return new Editor({
    element: null,
    extensions: [...docxExtensions, DocumentCommands],
    content: {
      type: "doc",
      content: [{ type: "paragraph", content: [{ type: "text", text }] }],
    },
  });
}

const rawXmlOf = (editor: Editor): string | null | undefined => {
  const attrs = editor.getAttributes("textStyle") as { w14RawXml?: unknown };
  return attrs.w14RawXml as string | null | undefined;
};

const marksOf = (json: JSONContent): Array<{ type?: string; attrs?: Record<string, unknown> }> => {
  const text = (json.content?.[0]?.content ?? [])[0] as JSONContent;
  return (text.marks ?? []) as never;
};

describe("Text Effects presets", () => {
  it("serializes a gallery preset into Word's w14 XML", () => {
    const outline = presetXml("outline:2")!;
    expect(outline).toContain('<w14:textOutline w14:w="12700"');
    expect(outline).toContain('<w14:srgbClr w14:val="2E75B5"/>');
    expect(presetXml("rotation:4")).toContain('<w14:camera w14:prst="obliqueTopLeft"/>');
    expect(presetXml("bevel:3")).toContain(
      '<w14:bevelT w14:w="50483" w14:h="50483" w14:prst="cross"/>',
    );
  });

  it("declines unknown preset tokens", () => {
    expect(presetXml("outline:99")).toBeUndefined();
    expect(applyPreset(null, "nope:1")).toBeUndefined();
    expect(applyPreset(null, "outline")).toBeUndefined();
    expect(applyPreset(null, "textFill:1")).toBeUndefined();
  });

  it("clears one family with :0 while preserving the others", () => {
    let xml = applyPreset(null, "outline:1")!;
    xml = applyPreset(xml, "glow:1")!;
    expect(xml).toContain("w14:textOutline");
    expect(xml).toContain("w14:glow");
    const cleared = applyPreset(xml, "outline:0")!;
    expect(cleared).not.toContain("w14:textOutline");
    expect(cleared).toContain("w14:glow");
  });

  it("applies a dialog patch and clears the null entries", () => {
    let xml = applyEffectsPatch(null, {
      outline: { color: "FF0000", widthPx: 1 } as never,
      glow: { color: "00B050", opacity: 0.5, radiusPx: 4 } as never,
    })!;
    expect(xml).toContain("w14:textOutline");
    expect(xml).toContain("w14:glow");
    xml = applyEffectsPatch(xml, { outline: null })!;
    expect(xml).not.toContain("w14:textOutline");
    expect(xml).toContain("w14:glow");
  });

  it("builds the Design tab's document theme XML", () => {
    expect(textEffectThemeXml("shadow")).toContain("w14:shadow");
    expect(textEffectThemeXml("none")).toBeNull();
    expect(textEffectThemeXml("bogus")).toBeUndefined();
  });

  it("stamps the Design theme onto the heading styles and clears it again", () => {
    const xml = textEffectThemeXml("glow")!;
    const patch = { w14RawXml: xml };
    const styles = stampStyleRunPatches(
      {},
      { title: patch, heading1: patch, heading2: patch, heading3: patch },
    );
    const defaults = (styles as { default: Record<string, { run?: Record<string, unknown> }> })
      .default;
    expect(defaults.title?.run?.w14RawXml).toContain("w14:glow");
    expect(defaults.heading3?.run?.w14RawXml).toContain("w14:glow");
    const cleared = stampStyleRunPatches(styles, {
      title: { w14RawXml: null },
      heading1: { w14RawXml: null },
      heading2: { w14RawXml: null },
      heading3: { w14RawXml: null },
    });
    const clearedDefaults = (
      cleared as { default: Record<string, { run?: Record<string, unknown> }> }
    ).default;
    expect(clearedDefaults.title?.run?.w14RawXml).toBeUndefined();
    expect(clearedDefaults.heading1?.run?.w14RawXml).toBeUndefined();
  });
});

describe("text-effects command on a real editor", () => {
  it("applies a preset to the selection and clears it again", () => {
    const editor = makeEditor();
    editor.commands.selectAll();
    expect(editor.commands["text-effects"]("glow:1")).toBe(true);
    expect(rawXmlOf(editor)).toContain("w14:glow");
    expect(editor.commands["text-effects"]("glow:0")).toBe(true);
    expect(rawXmlOf(editor) ?? null).toBeNull();
  });

  it("declines the options entry and unknown tokens without touching the marks", () => {
    const editor = makeEditor();
    editor.commands.selectAll();
    expect(editor.commands["text-effects"]("options")).toBe(false);
    expect(editor.commands["text-effects"]("outline:options")).toBe(false);
    expect(editor.commands["text-effects"]("outline:42")).toBe(false);
    expect(rawXmlOf(editor)).toBeUndefined();
  });

  it("patches all dialog sections at once", () => {
    const editor = makeEditor();
    editor.commands.selectAll();
    expect(
      editor.commands["text-effects-apply"](
        JSON.stringify({
          outline: { color: "C00000", widthPx: 2, dash: "dash" },
          bevel: { top: { widthPx: 4, heightPx: 3, preset: "softRound" } },
          rotation: null,
        }),
      ),
    ).toBe(true);
    const xml = rawXmlOf(editor)!;
    expect(xml).toContain("w14:textOutline");
    expect(xml).toContain("w14:prstDash");
    expect(xml).toContain("w14:props3d");
    expect(editor.commands["text-effects-apply"]("not json")).toBe(false);
  });

  it("round-trips the applied effects through DOCX byte-for-byte", () => {
    const editor = makeEditor();
    editor.commands.selectAll();
    editor.commands["text-effects"]("outline:3");
    editor.commands["text-effects"]("shadow:1");
    const before = rawXmlOf(editor)!;
    const parsed = parseDOCXSync(generateDOCXSync(editor.getJSON()) as Uint8Array);
    const mark = marksOf(parsed).find((m) => m.type === "textStyle");
    expect(mark?.attrs?.w14RawXml).toBe(before);
  });
});
