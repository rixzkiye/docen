// @vitest-environment node
import { execFileSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import {
  Document,
  Image,
  InlinePassthrough,
  Link,
  Paragraph,
  Tab,
  Table,
  TableCell,
  TableRow,
} from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import type { FlowPage, LaidOutTable } from "@docen/layout";
import { describe, expect, it } from "vitest";

import { pagesToPdf, type PdfPageShot } from "./export-pdf";
import {
  buildPdfDestinations,
  buildPdfOutline,
  buildPdfPageLabels,
  buildPdfStructure,
  buildPdfStructElements,
  type PdfStructureView,
} from "./pdf-structure";

/**
 * The editor-side R10-P3 wiring: the export path derives bookmarks, page
 * labels, named destinations and tagged figures/tables from the live print
 * run, and the emitted PDF bytes carry them. These tests drive the derivation
 * against real Tiptap documents plus a caret-map-shaped view, then assert the
 * writer's bytes — the same contract the fidelity harness checks end to end.
 */

const Text = TextNode.create({ name: "text", group: "inline" });

const t = (text: string, marks?: Record<string, unknown>[]): Record<string, unknown> => ({
  type: "text",
  text,
  ...(marks ? { marks } : {}),
});

const bookmark = (name: string, id: number, end = false): Record<string, unknown> => ({
  type: "inlinePassthrough",
  attrs: { data: JSON.stringify(end ? { bookmarkEnd: { id } } : { bookmarkStart: { id, name } }) },
});

const build = (doc: Record<string, unknown>): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [
      Document,
      Paragraph,
      Text,
      Tab,
      Link,
      InlinePassthrough,
      Image,
      Table,
      TableRow,
      TableCell,
    ],
    content: doc,
  });
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

const table = (): Record<string, unknown> => ({
  type: "table",
  content: [
    {
      type: "tableRow",
      content: ["A", "B"].map((label) => ({
        type: "tableCell",
        content: [{ type: "paragraph", content: [t(label)] }],
      })),
    },
  ],
});

/** The showcase shape these tests derive from: two chapters (a level-2 child
 *  under the first), a bookmark + internal link, a table and a picture. */
const DOC = {
  type: "doc",
  content: [
    { type: "paragraph", attrs: { heading: "Heading1" }, content: [t("Chapter 1")] },
    { type: "paragraph", attrs: { heading: "Heading2" }, content: [t("Section 1.1")] },
    {
      type: "paragraph",
      content: [
        t("Body with a "),
        bookmark("Target", 1),
        t("bookmark"),
        bookmark("Target", 1, true),
        t("."),
      ],
    },
    {
      type: "paragraph",
      content: [t("Go to target", [{ type: "link", attrs: { href: "#Target" } }])],
    },
    { type: "paragraph", attrs: { heading: "Heading1" }, content: [t("Chapter 2")] },
    table(),
    {
      type: "paragraph",
      content: [
        {
          type: "image",
          attrs: { src: "data:image/png;base64,iVBORw0KGgo=", width: 10, height: 10, alt: "A dog" },
        },
      ],
    },
  ],
};

/** The laid table the struct-element pass attaches to its page. */
const TABLE_BLOCK: LaidOutTable = {
  kind: "table",
  widthPx: 300,
  columnWidthsPx: [150, 150],
  heightPx: 20,
  rows: [],
};

/** Page-local caret boxes for a top-level block: page + line top in px. */
interface BlockBox {
  from: number;
  to: number;
  page: number;
  yPx: number;
}

const TOP_LEVEL_PAGES = [0, 0, 1, 1, 2, 2, 3];

function viewFor(editor: EditorType, pages: FlowPage[]): PdfStructureView {
  const boxes: BlockBox[] = [];
  let pos = 0;
  editor.state.doc.forEach((node, offset) => {
    boxes.push({
      from: offset,
      to: offset + node.nodeSize,
      page: TOP_LEVEL_PAGES[boxes.length]!,
      yPx: 50 + boxes.length * 10,
    });
    pos += node.nodeSize;
  });
  expect(pos).toBe(editor.state.doc.content.size);
  return {
    doc: editor.state.doc,
    sections: [
      { pageNumbering: { format: "lowerRoman" }, flow: { pageHeightPx: 1056 } },
      { pageNumbering: { start: 5, format: "decimal" }, flow: { pageHeightPx: 1056 } },
    ],
    sectionOfPage: [0, 0, 1, 1],
    pages,
    boxOf: (at) => {
      const box = boxes.find((candidate) => at >= candidate.from && at <= candidate.to);
      return box ? { page: box.page, yPx: box.yPx, heightPx: 16 } : null;
    },
  };
}

const flowPage = (items: FlowPage["items"]): FlowPage => ({ items });

describe("buildPdfOutline", () => {
  it("nests heading bookmarks by level with their page and top", () => {
    const editor = build(DOC);
    const outline = buildPdfOutline(
      editor.state.doc,
      viewFor(editor, [flowPage([]), flowPage([]), flowPage([]), flowPage([])]),
    );
    expect(outline).toHaveLength(2);
    expect(outline[0]!.title).toBe("Chapter 1");
    expect(outline[0]!.dest).toEqual({ pageIndex: 0, top: (1056 - 50) * 0.75 });
    expect(outline[0]!.children!.map((item) => item.title)).toEqual(["Section 1.1"]);
    expect(outline[0]!.children![0]!.dest).toEqual({ pageIndex: 0, top: (1056 - 60) * 0.75 });
    expect(outline[1]!.title).toBe("Chapter 2");
    expect(outline[1]!.dest).toEqual({ pageIndex: 2, top: (1056 - 90) * 0.75 });
    editor.destroy();
  });

  it("skips headings whose page is unmapped and re-parents level jumps", () => {
    const editor = build({
      type: "doc",
      content: [
        { type: "paragraph", attrs: { heading: "Heading1" }, content: [t("One")] },
        { type: "paragraph", attrs: { heading: "Heading3" }, content: [t("Deep")] },
        { type: "paragraph", attrs: { heading: "Heading1" }, content: [t("Unmapped")] },
      ],
    });
    const view = viewFor(editor, [flowPage([])]);
    view.boxOf = (at) => (at < 10 ? { page: 0, yPx: 10, heightPx: 12 } : null);
    const outline = buildPdfOutline(editor.state.doc, view);
    expect(outline).toHaveLength(1);
    expect(outline[0]!.title).toBe("One");
    // The level-3 heading has no shallower sibling after it — it parented
    // under "One" instead of becoming a root of its own.
    expect(outline[0]!.children!.map((item) => item.title)).toEqual(["Deep"]);
    editor.destroy();
  });
});

describe("buildPdfPageLabels", () => {
  it("emits one range per section with the displayed start number", () => {
    const labels = buildPdfPageLabels(
      [
        { pageNumbering: { format: "lowerRoman" } },
        { pageNumbering: { start: 5, format: "decimal" } },
      ],
      [0, 0, 1, 1],
    );
    expect(labels).toEqual([
      { startPageIndex: 0, style: "romanLower" },
      { startPageIndex: 2, style: "decimal", startNumber: 5 },
    ]);
  });

  it("keeps one entry per page index when a continuous section shares a page", () => {
    const labels = buildPdfPageLabels(
      [{}, { pageNumbering: { start: 4 } }],
      // Section 1 opens on the same physical page the previous section's flow
      // ends on (a continuous break).
      [0, 1, 1],
    );
    expect(labels).toEqual([
      { startPageIndex: 0, style: "decimal" },
      { startPageIndex: 1, style: "decimal", startNumber: 4 },
    ]);
  });

  it("degrades unsupported w:numFmt tokens to decimal and keeps none as no style", () => {
    const labels = buildPdfPageLabels(
      [{ pageNumbering: { format: "chineseCounting" } }, { pageNumbering: { format: "none" } }],
      [0, 1],
    );
    expect(labels.map((range) => range.style)).toEqual(["decimal", "none"]);
  });

  it("carries prefix to PdfPageLabelRange when pageNumbering.prefix is present", () => {
    const labels = buildPdfPageLabels(
      [
        { pageNumbering: { prefix: "Intro-", format: "lowerRoman" } },
        { pageNumbering: { prefix: "Chapter 1-", start: 1, format: "decimal" } },
        { pageNumbering: { prefix: "Appendix ", format: "upperLetter", start: 1 } },
      ],
      [0, 1, 2],
    );
    expect(labels).toEqual([
      { startPageIndex: 0, style: "romanLower", prefix: "Intro-" },
      { startPageIndex: 1, style: "decimal", prefix: "Chapter 1-" },
      { startPageIndex: 2, style: "alphaUpper", prefix: "Appendix " },
    ]);
  });
});

describe("buildPdfDestinations", () => {
  it("maps every bookmark name to its page destination, first occurrence winning", () => {
    const editor = build(DOC);
    const view = viewFor(editor, [flowPage([]), flowPage([]), flowPage([]), flowPage([])]);
    const destinations = buildPdfDestinations(editor.state.doc, view);
    // The bookmark rides top-level block 2 (page 1, yPx 70).
    expect(destinations).toEqual({ Target: { pageIndex: 1, y: (1056 - 70) * 0.75 } });
    editor.destroy();
  });
});

describe("buildPdfStructElements", () => {
  it("carries image alt text and table fragments onto their pages", () => {
    const editor = build(DOC);
    const pages = [
      flowPage([]),
      flowPage([]),
      flowPage([{ yPx: 0, block: TABLE_BLOCK }]),
      flowPage([]),
    ];
    const elements = buildPdfStructElements(editor.state.doc, viewFor(editor, pages));
    expect(elements).toEqual([
      { type: "Figure", pageIndex: 3, altText: "A dog", title: "A dog" },
      { type: "Table", pageIndex: 2 },
    ]);
    editor.destroy();
  });

  it("recurses into nested tables and skips drawings without a mapped page", () => {
    const nested: LaidOutTable = {
      ...TABLE_BLOCK,
      rows: [
        {
          heightPx: 20,
          cells: [
            {
              colspan: 1,
              rowspan: 1,
              innerWidthPx: 140,
              insets: { top: 4, right: 4, bottom: 4, left: 4 },
              stack: [{ yPx: 0, block: TABLE_BLOCK }],
            },
          ],
        },
      ],
    };
    const editor = build({
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "image",
              attrs: { src: "data:image/png;base64,iVBORw0KGgo=", width: 10, height: 10, alt: "X" },
            },
          ],
        },
      ],
    });
    const view = viewFor(editor, [flowPage([{ yPx: 0, block: nested }])]);
    view.boxOf = () => null;
    expect(buildPdfStructElements(editor.state.doc, view)).toEqual([
      { type: "Table", pageIndex: 0 },
      { type: "Table", pageIndex: 0 },
    ]);
    editor.destroy();
  });
});

describe("buildPdfStructure + pagesToPdf (exported bytes)", () => {
  it("emits /Outlines, /PageLabels, /Dests and /Alt from one derivation", async () => {
    const editor = build(DOC);
    const pages = [
      flowPage([]),
      flowPage([]),
      flowPage([{ yPx: 0, block: TABLE_BLOCK }]),
      flowPage([]),
    ];
    const structure = buildPdfStructure(viewFor(editor, pages));
    expect(structure.outline.map((item) => item.title)).toEqual(["Chapter 1", "Chapter 2"]);
    expect(structure.structElements).toHaveLength(2);

    // Page 1 carries the internal link the layout would extract from the
    // paragraph's #Target mark; the shot round-trips it to a /Dest annotation.
    const scene = (width: number, height: number): PdfPageShot["scene"] => ({
      width,
      height,
      nodes: [],
    });
    const shots: PdfPageShot[] = [0, 1, 2, 3].map((page) => ({
      width: 612,
      height: 792,
      scene: scene(612, 792),
      ...(page === 1
        ? {
            links: [
              { rect: [72, 600, 150, 615] as [number, number, number, number], url: "#Target" },
            ],
          }
        : {}),
    }));

    const blob = await pagesToPdf(shots, { tagged: true, ...structure });
    const pdf = Buffer.from(await blob.arrayBuffer()).toString("latin1");

    // Outline: the root count, per-item titles, the nested level and the
    // /XYZ destination the heading's page box resolved to.
    expect(pdf).toMatch(/\/Type \/Outlines \/First \d+ 0 R \/Last \d+ 0 R \/Count 3 >>/);
    expect(pdf).toContain("/Title (Chapter 1)");
    expect(pdf).toContain("/Title (Section 1.1)");
    expect(pdf).toContain("/Title (Chapter 2)");
    expect(pdf).toMatch(
      /\/Title \(Chapter 1\) \/Parent \d+ 0 R \/Dest \[ \d+ 0 R \/XYZ 0 754\.50 0 \][^\n]*\/Count 1/,
    );

    // Page labels: section 1 roman, section 2 restarts at 5 on page index 2.
    expect(pdf).toContain("/PageLabels << /Nums [ 0 << /S /r >> 2 << /S /D /St 5 >> ] >>");

    // Named destination + the link annotation that targets it — the internal
    // link resolves against the same name the bookmark contributed.
    expect(pdf).toMatch(/\/Dests << \/Target \[ \d+ 0 R \/XYZ 0\.00 739\.50 0 \] >>/);
    expect(pdf).toContain("/Dest (Target)");

    // Tagged structure: the figure's alt text/title and the table tag.
    expect(pdf).toContain("/Type /StructTreeRoot");
    expect(pdf).toContain("/S /Figure");
    expect(pdf).toContain("/Alt (A dog)");
    expect(pdf).toContain("/T (A dog)");
    expect(pdf).toContain("/S /Table");

    // The xref walks: pdfinfo parses the whole file.
    const tmpPdf = path.join(os.tmpdir(), `pdf-structure-${Date.now()}.pdf`);
    try {
      fs.writeFileSync(tmpPdf, Buffer.from(await blob.arrayBuffer()));
      const info = execFileSync("pdfinfo", [tmpPdf], { encoding: "utf-8" });
      expect(info).toContain("Pages:           4");
      expect(info).toMatch(/Tagged:\s+yes/);
    } finally {
      fs.rmSync(tmpPdf, { force: true });
    }
    editor.destroy();
  });

  it("emits /P (prefix) in /PageLabels when ranges carry prefixes", async () => {
    const pageLabels = buildPdfPageLabels(
      [
        { pageNumbering: { prefix: "Pref-", format: "lowerRoman" } },
        { pageNumbering: { prefix: "Sec-", format: "decimal", start: 1 } },
      ],
      [0, 1],
    );
    const scene = (width: number, height: number): PdfPageShot["scene"] => ({
      width,
      height,
      nodes: [],
    });
    const shots: PdfPageShot[] = [0, 1].map(() => ({
      width: 612,
      height: 792,
      scene: scene(612, 792),
    }));

    const blob = await pagesToPdf(shots, { pageLabels });
    const pdf = Buffer.from(await blob.arrayBuffer()).toString("latin1");

    expect(pdf).toContain(
      "/PageLabels << /Nums [ 0 << /S /r /P (Pref-) >> 1 << /S /D /P (Sec-) >> ] >>",
    );
  });
});
