import { Bold, Document, InlinePassthrough, Italic, Paragraph, Tab } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { IndexCommands, parseTaInstruction, parseXeInstruction } from "./index-commands";

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

/** A body paragraph whose runs end with the given inline atoms. */
const body = (...inline: Record<string, unknown>[]): Record<string, unknown> => ({
  type: "paragraph",
  content: inline,
});

const text = (value: string): Record<string, unknown> => ({ type: "text", text: value });

/** Word's invisible index marker: a cached-less XE fldSimple in the inline
 *  passthrough atom. */
const xe = (entry: string, flags = ""): Record<string, unknown> => ({
  type: "inlinePassthrough",
  attrs: {
    data: JSON.stringify({
      simpleField: { instruction: flags ? `XE "${entry}" ${flags}` : `XE "${entry}"` },
    }),
  },
});

const ta = (
  longCitation: string,
  category = 1,
  opts?: { short?: string; bold?: boolean; italic?: boolean },
): Record<string, unknown> => {
  const parts = [
    `TA \\l "${longCitation}"`,
    `\\s "${opts?.short ?? longCitation}"`,
    `\\c ${category}`,
  ];
  if (opts?.bold) parts.push("\\b");
  if (opts?.italic) parts.push("\\i");
  return {
    type: "inlinePassthrough",
    attrs: { data: JSON.stringify({ simpleField: { instruction: parts.join(" ") } }) },
  };
};

const docOf = (...blocks: Record<string, unknown>[]): Record<string, unknown> => ({
  type: "doc",
  content: blocks,
});

/** A headless editor with exactly the schema the index workflow touches. */
const build = (doc: Record<string, unknown>): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Tab, Bold, Italic, InlinePassthrough, IndexCommands],
    content: doc,
  });
  // element:null skips Tiptap's mount (and with it plugin installation) — the
  // same gap the canvas edit bridge patches by registering the sorted list.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

interface IndexEntryInfo {
  style: unknown;
  indent: unknown;
  tabStops: unknown;
  text: string;
  pages: string;
}

/** Flatten the document's Index-styled paragraphs into testable shapes. */
const entriesOf = (editor: EditorType): IndexEntryInfo[] => {
  const out: IndexEntryInfo[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== "paragraph") return true;
    const style = node.attrs.style;
    if (typeof style !== "string" || !/^Index\d$/.test(style)) return true;
    let entry = "";
    let pages = "";
    let inPages = false;
    node.content.forEach((child) => {
      if (child.type.name === "text") {
        if (inPages) pages += child.text ?? "";
        else entry += child.text ?? "";
      } else if (child.type.name === "tab") {
        inPages = true;
      }
    });
    out.push({
      style,
      indent: node.attrs.indent,
      tabStops: node.attrs.tabStops,
      text: entry,
      pages,
    });
    return true;
  });
  return out;
};

describe("insert-index command", () => {
  it("builds Index1/Index2 entries from XE fields with page numbers", () => {
    // Paragraph content starts at the paragraph pos + 1: 甲乙@1, 丙丁@6.
    const editor = build(
      docOf(body(text("甲乙"), xe("甲项")), body(text("丙丁"), xe("乙项:子项"))),
    );
    editor.commands.setTextSelection(1);
    const pages = new Map([
      [1, 2],
      [6, 5],
    ]);
    expect(editor.commands["insert-index"]((pos) => pages.get(pos) ?? 1)).toBe(true);
    const entries = entriesOf(editor);
    // `乙项:子项` nests: its main line has no page (no XE of its own — Word's
    // shape), the sub line carries it.
    expect(entries).toHaveLength(3);
    expect(entries[0]).toMatchObject({ style: "Index1", text: "甲项", pages: "2" });
    expect(entries[1]).toMatchObject({ style: "Index1", text: "乙项", pages: "" });
    // Word's built-in index indent: 220 twips per level past the first.
    expect(entries[2]).toMatchObject({
      style: "Index2",
      text: "子项",
      indent: { left: 220 },
      pages: "5",
    });
    // Every entry carries the dotted right tab stop.
    for (const entry of entries) {
      expect(entry.tabStops).toEqual([{ type: "right", position: 9350, leader: "dot" }]);
    }
    // The Index style definitions join the document styles.
    const styles = editor.state.doc.attrs.styles as {
      paragraphStyles?: { id?: string }[];
    };
    const ids = (styles.paragraphStyles ?? []).map((style) => style.id);
    expect(ids).toContain("Index1");
    expect(ids).toContain("Index2");
    editor.destroy();
  });

  it("sorts mains by the document collation and dedupes pages", () => {
    const editor = build(
      docOf(
        body(text("一"), xe("乙项"), xe("乙项")),
        body(text("二"), xe("甲项")),
        body(text("三"), xe("丙项")),
      ),
    );
    editor.commands.setTextSelection(1);
    // Pinyin collation: 丙(bing) < 甲(jia) < 乙(yi); the double 乙项 marker on
    // one page lands once.
    expect(editor.commands["insert-index"](() => 1)).toBe(true);
    expect(entriesOf(editor).map((entry) => entry.text)).toEqual(["丙项", "甲项", "乙项"]);
    expect(entriesOf(editor)[2]).toMatchObject({ pages: "1" });
    editor.destroy();
  });

  it("fails with no XE fields and inserts nothing", () => {
    const editor = build(docOf(body(text("正文"))));
    editor.commands.setTextSelection(1);
    expect(editor.commands["insert-index"]()).toBe(false);
    expect(editor.state.doc.childCount).toBe(1);
    editor.destroy();
  });
});

describe("update-index command", () => {
  it("rebuilds the entry block in place from the current XE fields", () => {
    const editor = build(docOf(body(text("甲乙"), xe("甲项")), body(text("丙丁"), xe("乙项"))));
    editor.commands.setTextSelection(1);
    editor.commands["insert-index"]();
    expect(entriesOf(editor)).toHaveLength(2);
    // Rename the XE inside the first body paragraph, then rebuild: the entry
    // follows and the stale Index paragraphs are gone (two remain, not four).
    editor.state.doc.descendants((node, at) => {
      if (node.type.name !== "inlinePassthrough") return true;
      const data = JSON.parse(String(node.attrs.data)) as {
        simpleField?: { instruction?: string };
      };
      if (data.simpleField?.instruction !== 'XE "甲项"') return true;
      editor.commands.command(({ state, dispatch }) => {
        const tr = state.tr.setNodeMarkup(at, undefined, {
          data: JSON.stringify({ simpleField: { instruction: 'XE "丙项"' } }),
        });
        dispatch?.(tr);
        return true;
      });
      return false;
    });
    expect(editor.commands["update-index"]()).toBe(true);
    expect(entriesOf(editor).map((entry) => entry.text)).toEqual(["丙项", "乙项"]);
    editor.destroy();
  });

  it("reports false when the doc has no index block", () => {
    const editor = build(docOf(body(text("甲乙"), xe("甲项"))));
    expect(editor.commands["update-index"]()).toBe(false);
    editor.destroy();
  });
});

describe("instruction parsers", () => {
  it("parses XE instructions with \\b, \\i, and \\t switches", () => {
    expect(parseXeInstruction('XE "Alpha"')).toEqual({
      entry: "Alpha",
      bold: false,
      italic: false,
      crossReference: undefined,
    });
    expect(parseXeInstruction('XE "Beta:Sub" \\b \\i')).toEqual({
      entry: "Beta:Sub",
      bold: true,
      italic: true,
      crossReference: undefined,
    });
    expect(parseXeInstruction('XE "Gamma" \\t "See Delta"')).toEqual({
      entry: "Gamma",
      bold: false,
      italic: false,
      crossReference: "See Delta",
    });
  });

  it("parses TA instructions with \\l, \\s, \\c, and formatting", () => {
    expect(parseTaInstruction('TA \\l "Roe v. Wade, 410 U.S. 113" \\s "Roe" \\c 1 \\b')).toEqual({
      longCitation: "Roe v. Wade, 410 U.S. 113",
      shortCitation: "Roe",
      category: 1,
      bold: true,
      italic: false,
    });
    expect(parseTaInstruction('TA \\l "17 U.S.C. § 101" \\c 2')).toEqual({
      longCitation: "17 U.S.C. § 101",
      shortCitation: undefined,
      category: 2,
      bold: false,
      italic: false,
    });
  });
});

describe("mark-entry and mark-entry-all commands", () => {
  it("mark-entry inserts an XE passthrough node at the selection", () => {
    const editor = build(docOf(body(text("Hello world"))));
    editor.commands.setTextSelection(1);
    expect(
      editor.commands["mark-entry"]({
        entry: "Greeting",
        subentry: "English",
        bold: true,
      }),
    ).toBe(true);

    let foundInstruction = "";
    editor.state.doc.descendants((node) => {
      if (node.type.name === "inlinePassthrough") {
        const data = JSON.parse(String(node.attrs.data)) as {
          simpleField?: { instruction?: string };
        };
        foundInstruction = data.simpleField?.instruction ?? "";
      }
      return true;
    });
    expect(foundInstruction).toBe('XE "Greeting:English" \\b');
    editor.destroy();
  });

  it("mark-entry-all marks every occurrence in the document", () => {
    const editor = build(docOf(body(text("Apple and orange")), body(text("Banana and Apple pie"))));
    expect(
      editor.commands["mark-entry-all"]({
        text: "Apple",
        bold: true,
      }),
    ).toBe(true);

    let count = 0;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "inlinePassthrough") {
        const data = JSON.parse(String(node.attrs.data)) as {
          simpleField?: { instruction?: string };
        };
        if (data.simpleField?.instruction === 'XE "Apple" \\b') count++;
      }
      return true;
    });
    expect(count).toBe(2);
    editor.destroy();
  });
});

describe("index formatting and cross-references", () => {
  it("renders bold page numbers and cross references", () => {
    const editor = build(
      docOf(
        body(text("Alpha text"), xe("Alpha", "\\b")),
        body(text("Beta text"), xe("Beta", '\\t "See Gamma"')),
      ),
    );
    editor.commands.setTextSelection(1);
    expect(editor.commands["insert-index"](() => 3)).toBe(true);

    const entries = entriesOf(editor);
    expect(entries).toHaveLength(2);
    expect(entries[0]).toMatchObject({ style: "Index1", text: "Alpha", pages: "3" });
    expect(entries[1]).toMatchObject({ style: "Index1", text: "Beta", pages: "See Gamma" });

    // Verify the bold mark is present on Alpha's page number
    let alphaPageHasBold = false;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "paragraph" && node.attrs.style === "Index1") {
        if (node.textContent.includes("Alpha")) {
          node.descendants((child) => {
            if (child.isText && child.text === "3") {
              alphaPageHasBold = child.marks.some((m) => m.type.name === "bold");
            }
            return true;
          });
        }
      }
      return true;
    });
    expect(alphaPageHasBold).toBe(true);
    editor.destroy();
  });
});

interface ToaEntryInfo {
  style: string;
  text: string;
  pages?: string;
  tabStops?: unknown;
}

const toaEntriesOf = (editor: EditorType): ToaEntryInfo[] => {
  const out: ToaEntryInfo[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== "paragraph") return true;
    const style = String(node.attrs.style ?? "");
    if (style !== "TOAHeading" && style !== "TableOfAuthorities") return true;
    let entry = "";
    let pages = "";
    let inPages = false;
    node.content.forEach((child) => {
      if (child.type.name === "text") {
        if (inPages) pages += child.text ?? "";
        else entry += child.text ?? "";
      } else if (child.type.name === "tab") {
        inPages = true;
      }
    });
    out.push({
      style,
      text: entry,
      ...(inPages ? { pages } : {}),
      tabStops: node.attrs.tabStops,
    });
    return true;
  });
  return out;
};

describe("insert-toa and update-toa commands", () => {
  it("builds Table of Authorities grouped by category with live page numbers", () => {
    const editor = build(
      docOf(
        body(text("Case 1"), ta("Smith v. Jones, 100 F.3d 1", 1, { short: "Smith" })),
        body(text("Statute 1"), ta("17 U.S.C. § 106", 2)),
        body(text("Case 2"), ta("Brown v. Board, 347 U.S. 483", 1, { short: "Brown" })),
      ),
    );
    editor.commands.setTextSelection(1);
    const pages = new Map([
      [1, 3],
      [10, 10],
      [22, 5],
    ]);
    expect(editor.commands["insert-toa"]((pos) => pages.get(pos) ?? 1)).toBe(true);

    const entries = toaEntriesOf(editor);
    // Heading Cases, Brown, Smith, Heading Statutes, 17 U.S.C. § 106
    expect(entries).toHaveLength(5);
    expect(entries[0]).toMatchObject({ style: "TOAHeading", text: "Cases" });
    expect(entries[1]).toMatchObject({
      style: "TableOfAuthorities",
      text: "Brown v. Board, 347 U.S. 483",
      pages: "5",
    });
    expect(entries[2]).toMatchObject({
      style: "TableOfAuthorities",
      text: "Smith v. Jones, 100 F.3d 1",
      pages: "3",
    });
    expect(entries[3]).toMatchObject({ style: "TOAHeading", text: "Statutes" });
    expect(entries[4]).toMatchObject({
      style: "TableOfAuthorities",
      text: "17 U.S.C. § 106",
      pages: "10",
    });

    // Check style stamping
    const styles = editor.state.doc.attrs.styles as {
      paragraphStyles?: { id?: string }[];
    };
    const ids = (styles.paragraphStyles ?? []).map((s) => s.id);
    expect(ids).toContain("TOAHeading");
    expect(ids).toContain("TableOfAuthorities");
    editor.destroy();
  });

  it("substitutes passim when citation appears on 5 or more pages", () => {
    const editor = build(
      docOf(
        body(text("P1"), ta("Omnipresent Case", 1)),
        body(text("P2"), ta("Omnipresent Case", 1)),
        body(text("P3"), ta("Omnipresent Case", 1)),
        body(text("P4"), ta("Omnipresent Case", 1)),
        body(text("P5"), ta("Omnipresent Case", 1)),
      ),
    );
    editor.commands.setTextSelection(1);
    let p = 1;
    expect(editor.commands["insert-toa"](() => p++)).toBe(true);

    const entries = toaEntriesOf(editor);
    expect(entries[1]).toMatchObject({
      style: "TableOfAuthorities",
      text: "Omnipresent Case",
      pages: "passim",
    });
    editor.destroy();
  });

  it("update-toa rebuilds in place when citations change", () => {
    const editor = build(docOf(body(text("Case"), ta("Initial Case", 1))));
    editor.commands.setTextSelection(1);
    editor.commands["insert-toa"](() => 1);
    expect(toaEntriesOf(editor)).toHaveLength(2); // Heading + Entry

    // Rename citation
    editor.state.doc.descendants((node, at) => {
      if (node.type.name !== "inlinePassthrough") return true;
      editor.commands.command(({ state, dispatch }) => {
        const tr = state.tr.setNodeMarkup(at, undefined, {
          data: JSON.stringify({
            simpleField: { instruction: 'TA \\l "Revised Case" \\s "Revised" \\c 1' },
          }),
        });
        dispatch?.(tr);
        return true;
      });
      return false;
    });

    expect(editor.commands["update-toa"](() => 2)).toBe(true);
    const updated = toaEntriesOf(editor);
    expect(updated).toHaveLength(2);
    expect(updated[1]).toMatchObject({
      style: "TableOfAuthorities",
      text: "Revised Case",
      pages: "2",
    });
    editor.destroy();
  });
});
