// @vitest-environment node
import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vitest";

import { Editor, Node as TextNode } from "../core";
import { generateDOCXSync, parseDOCXSync, compileDocument } from "../index";
import type { JSONContent } from "../index";
import { projectDocumentOptions } from "../layout";
import { Document } from "./document";
import { Paragraph } from "./paragraph";
import { SectionBreak } from "./section-break";

// The section-break command contract: w:type belongs to the section WHOSE
// sectPr carries it and says how that section starts relative to the previous
// one (ECMA-376 §17.6.22) — so the type of an inserted break goes on the
// FOLLOWING section's sectPr, never on the section being closed.
const Text = TextNode.create({ name: "text", group: "inline" });
const extensions = [Document, Paragraph, Text, SectionBreak];

const para = (text: string, attrs?: Record<string, unknown>): JSONContent => ({
  type: "paragraph",
  ...(attrs ? { attrs } : {}),
  content: [{ type: "text", text }],
});

const editorWith = (content: JSONContent[], attrs?: Record<string, unknown>): Editor => {
  const editor = new Editor({
    element: null,
    extensions,
    content: { type: "doc", ...(attrs ? { attrs } : {}), content },
  });
  // element:null skips mount() and with it plugin installation — the
  // production bridge registers the extension plugins by hand (edit-bridge.ts).
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

const documentXmlOf = (json: JSONContent): string =>
  new TextDecoder().decode(
    unzipSync(generateDOCXSync(json, { prepare: false }) as Uint8Array)["word/document.xml"],
  );

/** Every sectPr slice in body order — each closes (or, for the final one,
 *  ends) one section, so the Nth slice is section N's properties. */
const sectPrsOf = (xml: string): string[] =>
  [...xml.matchAll(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => m[0]!);

/** A sectPr's OWN w:type element — `toContain("w:type")` would match the
 *  unrelated `w:docGrid w:type="lines"` attribute. */
const breakTypeOf = (sectPr: string): string | undefined =>
  /<w:type\b[^>]*w:val="([^"]+)"/.exec(sectPr)?.[1];

/** The projected mark label at the end of section `index` (the painter maps
 *  these to Word's "Section Break (…)" labels; `true` is Next Page). */
const markAt = (json: JSONContent, index = 0): unknown => {
  const { sections } = projectDocumentOptions(compileDocument(json));
  const blocks = sections[index]?.blocks ?? [];
  const last = blocks[blocks.length - 1];
  return last?.kind === "paragraph" ? last.sectionEnd : undefined;
};

describe("setSectionBreak command w:type placement", () => {
  it("types the following (final) section's body-level sectPr on a fresh document", () => {
    const editor = editorWith([para("A"), para("B")]);
    editor.commands.setSectionBreak({ type: "oddPage" });

    const json = editor.getJSON();
    const content = json.content ?? [];
    expect(content).toHaveLength(3);
    // The closed section's paragraph keeps the boundary marker but claims no
    // type — stamping oddPage here would make section A start on parity.
    expect(content[0]?.attrs?.sectionProperties).toEqual({});
    // The fresh empty paragraph is the new section's first paragraph and
    // holds the caret.
    expect(content[1]?.content ?? []).toEqual([]);
    expect(editor.state.selection.$from.parent).toBe(editor.state.doc.child(1));
    // No later boundary exists: the new (final) section's sectPr is
    // body-level.
    expect(json.attrs?.sectionProperties).toEqual({ type: "oddPage" });
  });

  it("types the next existing boundary paragraph when the new section is not final", () => {
    const editor = editorWith([para("A"), para("B", { sectionProperties: {} }), para("C")]);
    editor.commands.setSectionBreak({ type: "continuous" });

    const content = editor.getJSON().content ?? [];
    expect(content).toHaveLength(4);
    expect(content[0]?.attrs?.sectionProperties).toEqual({});
    expect(content[1]?.content ?? []).toEqual([]);
    // B closes the inserted section — the type belongs to it, not to A.
    expect(content[2]?.attrs?.sectionProperties).toEqual({ type: "continuous" });
    // The final section's body-level sectPr stays untouched.
    expect(editor.getJSON().attrs?.sectionProperties ?? null).toBeNull();
  });

  it("clears an existing type for nextPage (the OOXML default) while both sections keep their geometry", () => {
    const pageSize = { width: 11906, height: 16838 };
    const editor = editorWith([para("A"), para("B")], {
      sectionProperties: { type: "oddPage", pageSize },
    });
    editor.commands.setSectionBreak();

    const json = editor.getJSON();
    // The closed section keeps a copy of its own sectPr — the split must not
    // wipe the section's page setup (it keeps its original start type too).
    expect(json.content?.[0]?.attrs?.sectionProperties).toEqual({ type: "oddPage", pageSize });
    const body = json.attrs?.sectionProperties as Record<string, unknown>;
    expect("type" in body).toBe(false);
    // The rest of the following section's sectPr survives the clear.
    expect(body.pageSize).toEqual(pageSize);
  });

  it("leaves an absent body sectPr absent for a plain nextPage break", () => {
    const editor = editorWith([para("A"), para("B")]);
    editor.commands.setSectionBreak();

    const json = editor.getJSON();
    expect(json.content?.[0]?.attrs?.sectionProperties).toEqual({});
    expect(json.attrs?.sectionProperties ?? null).toBeNull();
  });
});

describe("split geometry preservation", () => {
  it("copies the final section's page geometry onto the closing paragraph", () => {
    const pageSize = { width: 12240, height: 15840 };
    const pageMargin = { left: 1800, right: 1800 };
    const editor = editorWith([para("A"), para("B")], {
      sectionProperties: { pageSize, pageMargin },
    });
    editor.commands.setSectionBreak({ type: "oddPage" });

    const json = editor.getJSON();
    // The first section keeps its own geometry after the split…
    expect(json.content?.[0]?.attrs?.sectionProperties).toEqual({ pageSize, pageMargin });
    // …and the new final section keeps the geometry plus the new type.
    expect(json.attrs?.sectionProperties).toEqual({ pageSize, pageMargin, type: "oddPage" });

    const { sections } = compileDocument(json);
    const propsOf = (i: number) => {
      const props = sections[i]?.properties;
      const size =
        typeof props?.pageSize === "object" && props.pageSize ? props.pageSize : undefined;
      const margin =
        typeof props?.pageMargin === "object" && props.pageMargin ? props.pageMargin : undefined;
      return { size, margin, type: props?.type };
    };
    expect(propsOf(0).size?.width).toBe(12240);
    expect(propsOf(0).margin?.left).toBe(1800);
    expect(propsOf(1).size?.width).toBe(12240);
    expect(propsOf(1).type).toBe("oddPage");
  });

  it("copies a later boundary's sectPr in the mid-document case", () => {
    const pageSize = { width: 12240, height: 15840 };
    const pageMargin = { left: 1800 };
    const editor = editorWith([
      para("A"),
      para("B", { sectionProperties: { pageSize, pageMargin } }),
      para("C"),
    ]);
    editor.commands.setSectionBreak({ type: "continuous" });

    const content = editor.getJSON().content ?? [];
    expect(content[0]?.attrs?.sectionProperties).toEqual({ pageSize, pageMargin });
    expect(content[2]?.attrs?.sectionProperties).toEqual({
      pageSize,
      pageMargin,
      type: "continuous",
    });
  });
});

describe("section break w:type round-trip", () => {
  it("writes the type into the following sectPr and re-parses to the same section types", () => {
    const editor = editorWith([para("A"), para("B")]);
    editor.commands.setSectionBreak({ type: "oddPage" });
    const json = editor.getJSON();

    const sectPrs = sectPrsOf(documentXmlOf(json));
    expect(sectPrs).toHaveLength(2);
    expect(breakTypeOf(sectPrs[0]!)).toBeUndefined();
    expect(breakTypeOf(sectPrs[1]!)).toBe("oddPage");

    const reparsed = parseDOCXSync(generateDOCXSync(json, { prepare: false }));
    const types = compileDocument(reparsed).sections.map((s) => s.properties?.type ?? "nextPage");
    expect(types).toEqual(["nextPage", "oddPage"]);
  });

  it("places a mid-document break's type in the next boundary section's sectPr", () => {
    const editor = editorWith([para("A"), para("B", { sectionProperties: {} }), para("C")]);
    editor.commands.setSectionBreak({ type: "continuous" });
    const json = editor.getJSON();

    const sectPrs = sectPrsOf(documentXmlOf(json));
    expect(sectPrs).toHaveLength(3);
    expect(breakTypeOf(sectPrs[0]!)).toBeUndefined();
    expect(breakTypeOf(sectPrs[1]!)).toBe("continuous");
    expect(breakTypeOf(sectPrs[2]!)).toBeUndefined();

    const reparsed = parseDOCXSync(generateDOCXSync(json, { prepare: false }));
    const types = compileDocument(reparsed).sections.map((s) => s.properties?.type ?? "nextPage");
    expect(types).toEqual(["nextPage", "continuous", "nextPage"]);
  });
});

describe("section break mark label follows the following section", () => {
  it("labels an oddPage break on the paragraph before it", () => {
    const editor = editorWith([para("A"), para("B")]);
    editor.commands.setSectionBreak({ type: "oddPage" });
    expect(markAt(editor.getJSON())).toBe("oddPage");
  });

  it("labels a continuous break on the paragraph before it", () => {
    const editor = editorWith([para("A"), para("B")]);
    editor.commands.setSectionBreak({ type: "continuous" });
    expect(markAt(editor.getJSON())).toBe("continuous");
  });

  it("collapses nextPage to the painter's default label", () => {
    const editor = editorWith([para("A"), para("B")]);
    editor.commands.setSectionBreak();
    expect(markAt(editor.getJSON())).toBe(true);
  });

  it("labels a mid-document break from the boundary paragraph it types", () => {
    const editor = editorWith([para("A"), para("B", { sectionProperties: {} }), para("C")]);
    editor.commands.setSectionBreak({ type: "evenPage" });
    expect(markAt(editor.getJSON())).toBe("evenPage");
  });
});
