import { Document, InlinePassthrough, Link, Paragraph, Tab, TocField } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { TocCommands } from "./toc";

// Tiptap's schema needs the plain text node; the engine builds the same shape
// internally (tiptapNodeExtensions) but does not export it standalone.
const Text = TextNode.create({ name: "text", group: "inline" });

const heading = (level: number, text: string): Record<string, unknown> => ({
  type: "paragraph",
  attrs: { heading: `Heading${level}` },
  content: [{ type: "text", text }],
});

const docOf = (...blocks: Record<string, unknown>[]): Record<string, unknown> => ({
  type: "doc",
  content: blocks,
});

/** The [from, to) text range of the heading paragraph named `text` — entry
 *  paragraphs carry a TOC style and no heading attr, so they never match. */
const headingTextRange = (editor: EditorType, text: string): [number, number] => {
  let range: [number, number] = [0, 0];
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === "paragraph" && node.attrs.heading && node.textContent === text) {
      range = [pos + 1, pos + 1 + text.length];
      return false;
    }
    return true;
  });
  return range;
};

/** A headless editor with exactly the schema the TOC workflow touches. */
const build = (doc: Record<string, unknown>): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, Tab, Link, TocField, InlinePassthrough, TocCommands],
    content: doc,
  });
  // element:null skips Tiptap's mount (and with it plugin installation) — the
  // same gap the canvas edit bridge patches by registering the sorted list.
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

interface EntryInfo {
  style: unknown;
  indent: unknown;
  tabStops: unknown;
  text: string;
  linkHref: unknown;
  page: string;
}

/** Flatten the first tocField's entries into testable shapes. */
const entriesOf = (editor: EditorType): EntryInfo[] => {
  const out: EntryInfo[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== "tocField") return true;
    node.forEach((entry) => {
      let text = "";
      let linkHref: unknown = null;
      let page = "";
      let inPage = false;
      entry.content.forEach((child) => {
        if (child.type.name === "text") {
          if (inPage) page += child.text ?? "";
          else text += child.text ?? "";
          const link = child.marks.find((m) => m.type.name === "link");
          if (link) linkHref = link.attrs.href;
        } else if (child.type.name === "tab") {
          inPage = true;
        }
      });
      out.push({
        style: entry.attrs.style,
        indent: entry.attrs.indent,
        tabStops: entry.attrs.tabStops,
        text,
        linkHref,
        page,
      });
    });
    return false;
  });
  return out;
};

describe("parsed DOCX TOC options (item 8)", () => {
  it("stamps \\t and \\b under office-open's option names on insert", () => {
    const editor = build(
      docOf(
        heading(1, "Outside"),
        {
          type: "paragraph",
          content: [
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ bookmarkStart: { id: 9, name: "Scope" } }) },
            },
            { type: "text", text: "Scope" },
          ],
        },
        {
          type: "paragraph",
          attrs: { style: "SpecialTitle" },
          content: [{ type: "text", text: "Custom Inside" }],
        },
        {
          type: "paragraph",
          content: [
            {
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify({ bookmarkEnd: { id: 9 } }) },
            },
          ],
        },
      ),
    );
    editor.commands.setTextSelection(1);
    expect(
      editor.commands.toc(undefined, undefined, {
        styles: "SpecialTitle,1",
        bookmark: "Scope",
      }),
    ).toBe(true);
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tocField") {
        expect(node.attrs.options).toEqual({
          headingStyleRange: "1-3",
          hyperlink: true,
          stylesWithLevels: [{ styleName: "SpecialTitle", level: 1 }],
          entriesFromBookmark: "Scope",
        });
      }
      return true;
    });
    // Only the bookmarked custom style joins the entries.
    expect(entriesOf(editor).map((e) => e.text)).toEqual(["Custom Inside"]);
    editor.destroy();
  });

  it("update-toc honors parsed stylesWithLevels and entriesFromBookmark", () => {
    const seed = (data: object) => ({
      type: "inlinePassthrough",
      attrs: { data: JSON.stringify(data) },
    });
    const editor = build(
      docOf(
        heading(1, "Before"),
        {
          type: "paragraph",
          content: [seed({ bookmarkStart: { id: 10, name: "Scope" } })],
        },
        {
          type: "paragraph",
          attrs: { style: "MyStyle" },
          content: [{ type: "text", text: "Custom Inside" }],
        },
        heading(2, "Inside Heading"),
        {
          type: "paragraph",
          content: [seed({ bookmarkEnd: { id: 10 } })],
        },
        heading(1, "After"),
        {
          type: "tocField",
          attrs: {
            options: {
              headingStyleRange: "1-3",
              entriesFromBookmark: "Scope",
              stylesWithLevels: [{ styleName: "MyStyle", level: 2 }],
            },
          },
          content: [{ type: "paragraph" }],
        },
      ),
    );
    expect(editor.commands["update-toc"]()).toBe(true);
    const entries = entriesOf(editor);
    expect(entries.map((e) => e.text)).toEqual(["Custom Inside", "Inside Heading"]);
    expect(entries.map((e) => e.style)).toEqual(["TOC2", "TOC2"]);
    editor.destroy();
  });
});

describe("toc command", () => {
  it("builds one entry per heading 1-3 with style, tab, link, and page", () => {
    const editor = build(
      docOf(heading(1, "Alpha"), heading(2, "Beta"), heading(3, "Gamma"), {
        type: "paragraph",
        content: [{ type: "text", text: "body" }],
      }),
    );
    editor.commands.setTextSelection(1);
    // Doc layout (paragraph content starts at the paragraph pos + 1):
    // Alpha@1, Beta@8, Gamma@14.
    const pages = new Map([
      [1, 2],
      [8, 5],
      [14, 9],
    ]);
    const pageOf = (pos: number): number => pages.get(pos) ?? 1;
    expect(editor.commands.toc(pageOf)).toBe(true);
    const entries = entriesOf(editor);
    expect(entries).toHaveLength(3);
    expect(entries[0]).toMatchObject({
      style: "TOC1",
      text: "Alpha",
      linkHref: "#_Toc1",
      page: "2",
    });
    // Word's built-in TOC indent: 220 twips per level past the first.
    expect(entries[1]).toMatchObject({
      style: "TOC2",
      text: "Beta",
      indent: { left: 220 },
      page: "5",
    });
    expect(entries[2]).toMatchObject({ style: "TOC3", text: "Gamma", indent: { left: 440 } });
    // Every entry carries the dotted right tab stop.
    for (const entry of entries) {
      expect(entry.tabStops).toEqual([{ type: "right", position: 9350, leader: "dot" }]);
    }
    editor.destroy();
  });

  it("skips headings past level 3 and empty headings", () => {
    const editor = build(
      docOf(heading(1, "One"), heading(4, "Deep"), {
        type: "paragraph",
        attrs: { heading: "Heading2" },
      }),
    );
    editor.commands.setTextSelection(1);
    editor.commands.toc();
    expect(entriesOf(editor)).toHaveLength(1);
    editor.destroy();
  });

  it("fails with no headings and inserts nothing", () => {
    const editor = build(docOf({ type: "paragraph" }));
    editor.commands.setTextSelection(1);
    expect(editor.commands.toc()).toBe(false);
    expect(editor.state.doc.firstChild?.type.name).toBe("paragraph");
    editor.destroy();
  });

  it("stamps the default field switches on the tocField", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    editor.commands.setTextSelection(1);
    editor.commands.toc();
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tocField") {
        expect(node.attrs.options).toEqual({ headingStyleRange: "1-3", hyperlink: true });
      }
      return true;
    });
    editor.destroy();
  });
});

describe("update-toc command", () => {
  it("rebuilds entries from current headings and keeps the field switches", () => {
    const editor = build(docOf(heading(1, "Alpha"), heading(2, "Beta")));
    editor.commands.setTextSelection(1);
    editor.commands.toc();
    // Rename a heading, then update: the entry text follows. The inserted TOC
    // sits before the headings.
    const [from, to] = headingTextRange(editor, "Alpha");
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.insertText("Renamed", from, to));
      return true;
    });
    expect(editor.commands["update-toc"]()).toBe(true);
    const entries = entriesOf(editor);
    expect(entries.map((e) => e.text)).toEqual(["Renamed", "Beta"]);
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tocField") {
        expect(node.attrs.options).toEqual({ headingStyleRange: "1-3", hyperlink: true });
      }
      return true;
    });
    editor.destroy();
  });

  it("reports false when the doc has no tocField", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    expect(editor.commands["update-toc"]()).toBe(false);
    editor.destroy();
  });
});

describe("custom toc insert options", () => {
  it("honors the level window, leader, and unaligned page numbers", () => {
    const editor = build(docOf(heading(1, "Alpha"), heading(3, "Deep")));
    editor.commands.setTextSelection(1);
    editor.commands.toc(() => 5, undefined, {
      headingRange: "1-2",
      leader: "hyphen",
      alignPageNumbers: false,
    });
    const entries = entriesOf(editor);
    // The \o window excludes level 3; the unaligned number trails the entry
    // text after a space (entriesOf only splits pages off at a tab).
    expect(entries.map((e) => e.text)).toEqual(["Alpha 5"]);
    // Unaligned numbers trail the text after a space — no right tab, no leader.
    expect(entries[0].page).toBe("");
    expect(entries[0].tabStops).toBeNull();
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tocField") {
        expect(node.attrs.options).toEqual({ headingStyleRange: "1-2", hyperlink: true });
      }
      return true;
    });
    editor.destroy();
  });

  it("omits the number run entirely when page numbers are off", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    editor.commands.setTextSelection(1);
    editor.commands.toc(() => 5, undefined, { showPageNumbers: false });
    const entries = entriesOf(editor);
    expect(entries[0]).toMatchObject({ style: "TOC1", text: "Alpha", page: "" });
    expect(entries[0].tabStops).toBeNull();
    editor.destroy();
  });
});

describe("update-toc-page command", () => {
  it("rewrites only the trailing numbers, keeping every entry's text", () => {
    const editor = build(docOf(heading(1, "Alpha"), heading(2, "Beta")));
    editor.commands.setTextSelection(1);
    editor.commands.toc(() => 2);
    expect(editor.commands["update-toc-page"](() => 7)).toBe(true);
    const entries = entriesOf(editor);
    expect(entries.map((e) => e.text)).toEqual(["Alpha", "Beta"]);
    expect(entries.map((e) => e.page)).toEqual(["7", "7"]);
    editor.destroy();
  });

  it("reports false and keeps the old numbers when no heading maps to a page", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    editor.commands.setTextSelection(1);
    editor.commands.toc(() => 2);
    expect(editor.commands["update-toc-page"](() => null)).toBe(false);
    expect(entriesOf(editor).map((e) => e.page)).toEqual(["2"]);
    editor.destroy();
  });
});

describe("remove-toc command", () => {
  it("deletes the whole TOC block but keeps the headings", () => {
    const editor = build(docOf(heading(1, "Alpha"), heading(2, "Beta")));
    editor.commands.setTextSelection(1);
    editor.commands.toc();
    expect(editor.commands["remove-toc"]()).toBe(true);
    expect(editor.state.doc.firstChild?.attrs.heading).toBe("Heading1");
    expect(entriesOf(editor)).toHaveLength(0);
    editor.destroy();
  });

  it("reports false when the doc has no tocField", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    expect(editor.commands["remove-toc"]()).toBe(false);
    editor.destroy();
  });
});

// ── Table of Figures (the \c caption switch) ──

const caption = (label: string, text: string): Record<string, unknown> => ({
  type: "paragraph",
  attrs: { style: "Caption" },
  content: [
    { type: "text", text: `${label} ` },
    {
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: { instruction: `SEQ ${label} * ARABIC`, cachedValue: "1" },
        }),
      },
    },
    { type: "text", text: `: ${text}` },
  ],
});

describe("table-of-figures command", () => {
  it("builds one entry per matching caption with the c switch stamped", () => {
    const editor = build(
      docOf(heading(1, "Intro"), caption("Figure", "Alpha chart"), caption("Table", "Beta grid")),
    );
    editor.commands.setTextSelection(1);
    expect(editor.commands["table-of-figures"](() => 4, undefined, "Figure")).toBe(true);
    const entries = entriesOf(editor);
    // Only the Figure caption counts — Table captions belong to another \c.
    expect(entries).toHaveLength(1);
    // The entry text resolves the SEQ atom's cached value — Word's table of
    // figures shows the number ("Figure 1: Alpha chart"), not "Figure : ...".
    expect(entries[0]).toMatchObject({ style: "TOC1", text: "Figure 1: Alpha chart", page: "4" });
    expect(entries[0].linkHref).toBeNull();
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tocField") {
        // Word's table-of-figures switch is \c (office-open's
        // captionLabelIncludingNumbers); \a drops the label+number on a Word
        // update and LibreOffice discards the table entirely.
        expect(node.attrs.options).toEqual({ captionLabelIncludingNumbers: "Figure" });
      }
      return true;
    });
    editor.destroy();
  });

  it("fails with no matching captions and inserts nothing", () => {
    const editor = build(docOf(caption("Table", "Beta grid")));
    editor.commands.setTextSelection(1);
    expect(editor.commands["table-of-figures"]()).toBe(false);
    expect(editor.state.doc.firstChild?.attrs.style).toBe("Caption");
    editor.destroy();
  });
});

describe("update-figures command", () => {
  it("rebuilds the figure table from current captions and keeps the label", () => {
    const editor = build(docOf(caption("Figure", "Alpha chart")));
    editor.commands.setTextSelection(1);
    editor.commands["table-of-figures"]();
    // Add a second caption, then update: the entry list follows.
    editor.commands.command(({ state, dispatch }) => {
      const node = state.schema.nodeFromJSON(caption("Figure", "Second chart"));
      dispatch?.(state.tr.insert(state.doc.content.size - 1, node));
      return true;
    });
    expect(editor.commands["update-figures"]()).toBe(true);
    expect(entriesOf(editor).map((e) => e.text)).toEqual([
      "Figure 1: Alpha chart",
      "Figure 1: Second chart",
    ]);
    editor.destroy();
  });

  it("reports false when the doc has no figure table", () => {
    const editor = build(docOf(heading(1, "Alpha")));
    editor.commands.setTextSelection(1);
    editor.commands.toc(); // a heading TOC is not a figure table
    expect(editor.commands["update-figures"]()).toBe(false);
    editor.destroy();
  });
});

describe("heading TOC vs figure table isolation", () => {
  it("update-toc and remove-toc ignore the \\c figure table", () => {
    const editor = build(docOf(caption("Figure", "Alpha chart")));
    editor.commands.setTextSelection(1);
    editor.commands["table-of-figures"]();
    // A figures-only document must not be rebuilt as (or deleted by) the
    // heading-TOC commands.
    expect(editor.commands["update-toc"]()).toBe(false);
    expect(editor.commands["remove-toc"]()).toBe(false);
    expect(entriesOf(editor)).toHaveLength(1);
    editor.destroy();
  });
});

describe("figure table with chapter-numbered captions", () => {
  // The real Insert Caption field code once "Include chapter number" is on —
  // the label regex must read past the \* and \s switches.
  const chapterCaption = (
    label: string,
    text: string,
    chapter: string,
  ): Record<string, unknown> => ({
    type: "paragraph",
    attrs: { style: "Caption" },
    content: [
      { type: "text", text: `${label} ` },
      {
        type: "inlinePassthrough",
        attrs: {
          data: JSON.stringify({
            simpleField: {
              instruction: `SEQ ${label} \\* ARABIC \\s 1`,
              cachedValue: `${chapter}-1`,
            },
          }),
        },
      },
      { type: "text", text: `: ${text}` },
    ],
  });

  it("keeps listing chapter-prefixed captions through insert and update", () => {
    const editor = build(
      docOf(
        heading(1, "第一章"),
        chapterCaption("Figure", "Alpha chart", "1"),
        heading(1, "第二章"),
        chapterCaption("Figure", "Beta chart", "2"),
        caption("Table", "Grid"),
      ),
    );
    editor.commands.setTextSelection(1);
    expect(editor.commands["table-of-figures"](() => 4, undefined, "Figure")).toBe(true);
    expect(entriesOf(editor).map((e) => e.text)).toEqual([
      "Figure 1-1: Alpha chart",
      "Figure 2-1: Beta chart",
    ]);

    // A new chapter-numbered caption joins the rebuilt table.
    editor.commands.command(({ state, dispatch }) => {
      const node = state.schema.nodeFromJSON(chapterCaption("Figure", "Gamma chart", "2"));
      dispatch?.(state.tr.insert(state.doc.content.size - 1, node));
      return true;
    });
    expect(editor.commands["update-figures"](() => 7)).toBe(true);
    const entries = entriesOf(editor);
    expect(entries.map((e) => e.text)).toEqual([
      "Figure 1-1: Alpha chart",
      "Figure 2-1: Beta chart",
      "Figure 2-1: Gamma chart",
    ]);
    expect(entries.map((e) => e.page)).toEqual(["7", "7", "7"]);
    editor.destroy();
  });

  it("stamps _Toc bookmarks on headings and links entries to them", () => {
    const editor = build(docOf(heading(1, "Chapter 1"), heading(2, "Section 1.1")));
    editor.commands.setTextSelection(1);
    expect(editor.commands.toc()).toBe(true);

    const entries = entriesOf(editor);
    expect(entries[0]!.linkHref).toBe("#_Toc1");
    expect(entries[1]!.linkHref).toBe("#_Toc2");

    // Headings in the document now carry matching _Toc bookmarkStart/bookmarkEnd atoms.
    const bookmarks: string[] = [];
    editor.state.doc.descendants((node) => {
      if (node.type.name === "inlinePassthrough") {
        try {
          const data = JSON.parse(String(node.attrs.data ?? "{}")) as {
            bookmarkStart?: { name?: string };
          };
          if (data.bookmarkStart?.name) bookmarks.push(data.bookmarkStart.name);
        } catch {
          /* skip */
        }
      }
      return true;
    });
    expect(bookmarks).toContain("_Toc1");
    expect(bookmarks).toContain("_Toc2");
    editor.destroy();
  });

  it("supports \\t custom styles switch to map custom paragraph styles to TOC levels", () => {
    const editor = build(
      docOf(
        {
          type: "paragraph",
          attrs: { style: "SpecialTitle" },
          content: [{ type: "text", text: "Special Header" }],
        },
        heading(1, "Normal Heading"),
      ),
    );
    editor.commands.setTextSelection(1);
    expect(
      editor.commands.toc(undefined, undefined, {
        headingRange: "1-3",
        styles: "SpecialTitle,1",
      }),
    ).toBe(true);

    const entries = entriesOf(editor);
    expect(entries).toHaveLength(2);
    expect(entries[0]).toMatchObject({ style: "TOC1", text: "Special Header" });
    expect(entries[1]).toMatchObject({ style: "TOC1", text: "Normal Heading" });
    editor.destroy();
  });

  it("update-toc-page matches headings via bookmark names even when heading text is renamed", () => {
    const editor = build(docOf(heading(1, "Original Title")));
    editor.commands.setTextSelection(1);
    expect(editor.commands.toc(() => 3)).toBe(true);
    expect(entriesOf(editor)[0]!.page).toBe("3");

    // Rename the heading text in the doc
    const [from, to] = headingTextRange(editor, "Original Title");
    editor.commands.command(({ state, dispatch }) => {
      dispatch?.(state.tr.insertText("Completely Different Title", from, to));
      return true;
    });

    // Update page numbers: since the bookmark _Toc1 still lives on the heading,
    // update-toc-page resolves the heading and updates its page number!
    expect(editor.commands["update-toc-page"](() => 9)).toBe(true);
    const entries = entriesOf(editor);
    expect(entries[0]!.page).toBe("9");
    editor.destroy();
  });

  it("supports \\h switch to disable hyperlinks in TOC entries", () => {
    const editor = build(docOf(heading(1, "Alpha"), heading(2, "Beta")));
    editor.commands.setTextSelection(1);
    expect(editor.commands.toc(undefined, undefined, { hyperlink: false })).toBe(true);

    const entries = entriesOf(editor);
    expect(entries).toHaveLength(2);
    expect(entries[0]!.linkHref).toBeNull();
    expect(entries[1]!.linkHref).toBeNull();
    editor.destroy();
  });

  it("supports \\b bookmark scope switch to limit TOC entries to a bookmark", () => {
    const seed = (data: object) => ({
      type: "inlinePassthrough",
      attrs: { data: JSON.stringify(data) },
    });
    const editor = build(
      docOf(
        heading(1, "Outside Before"),
        {
          type: "paragraph",
          content: [
            seed({ bookmarkStart: { id: 10, name: "SectionBookmark" } }),
            { type: "text", text: "Scope Start" },
          ],
        },
        heading(1, "Inside Bookmark 1"),
        heading(2, "Inside Bookmark 2"),
        {
          type: "paragraph",
          content: [{ type: "text", text: "Scope End" }, seed({ bookmarkEnd: { id: 10 } })],
        },
        heading(1, "Outside After"),
      ),
    );
    editor.commands.setTextSelection(1);
    expect(editor.commands.toc(undefined, undefined, { bookmark: "SectionBookmark" })).toBe(true);

    const entries = entriesOf(editor);
    expect(entries).toHaveLength(2);
    expect(entries[0]!.text).toBe("Inside Bookmark 1");
    expect(entries[1]!.text).toBe("Inside Bookmark 2");
    editor.destroy();
  });

  it("supports \\u outlineLevel switch to include/exclude outline level paragraphs", () => {
    const editorWithOutline = build(
      docOf({
        type: "paragraph",
        attrs: { outlineLevel: 0 },
        content: [{ type: "text", text: "Outline Level 1 Paragraph" }],
      }),
    );
    editorWithOutline.commands.setTextSelection(1);
    expect(
      editorWithOutline.commands.toc(undefined, undefined, {
        useAppliedParagraphOutlineLevel: true,
      }),
    ).toBe(true);
    expect(entriesOf(editorWithOutline)).toHaveLength(1);
    expect(entriesOf(editorWithOutline)[0]!.text).toBe("Outline Level 1 Paragraph");
    editorWithOutline.destroy();

    const editorWithoutOutline = build(
      docOf({
        type: "paragraph",
        attrs: { outlineLevel: 0 },
        content: [{ type: "text", text: "Outline Level 1 Paragraph" }],
      }),
    );
    editorWithoutOutline.commands.setTextSelection(1);
    expect(
      editorWithoutOutline.commands.toc(undefined, undefined, {
        useAppliedParagraphOutlineLevel: false,
      }),
    ).toBe(false);
    editorWithoutOutline.destroy();
  });
});
