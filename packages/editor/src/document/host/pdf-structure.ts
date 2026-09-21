/**
 * PDF structure & navigation derivation — the editor half of the R10-P3
 * wiring. `pagesToPdf` can always emit bookmarks, page labels, named
 * destinations and tagged figures/tables; this module turns the live
 * print-run state (the PM document + the pagination the canvas just produced)
 * into those options so File → Export as PDF carries them.
 *
 * The inputs are deliberately the same projections the export already reads —
 * never a second pagination model:
 *
 * - `boxOf` is the caret map's page-local box (EditBridge.caretBox), the
 *   position → page/top mapping the caret and overlays are placed through.
 * - `sections[].pageNumbering` + `computePageNumberOffsets` are the exact
 *   source the PAGE field paints its displayed numbers from, so the PDF's
 *   page labels and the printed footer numbers agree.
 * - `pages` is the laid flow (`host.pages()`): table struct elements attach
 *   to the page fragments the layout actually produced.
 */

import { detectHeadingLevel, type StylesOptions } from "@docen/docx";
import {
  computePageNumberOffsets,
  type FlowPage,
  type LaidOutBlock,
  type ProjectedPageNumbering,
} from "@docen/layout";
import type { Node as PMNode } from "@tiptap/pm/model";

import type {
  PdfDestination,
  PdfOutlineItem,
  PdfPageLabelRange,
  PdfStructElement,
} from "../export-pdf";

/** The print-run slice the structure pass reads — the IODomain host view. */
export interface PdfStructureView {
  /** The live document (undefined before the editor mounts). */
  doc: PMNode | undefined;
  /** The print run's sections, with their w:pgNumType and page geometry. */
  sections: readonly {
    pageNumbering?: ProjectedPageNumbering;
    flow?: { pageHeightPx: number };
  }[];
  /** Page index → section index, parallel to the print run's page list. */
  sectionOfPage: readonly number[];
  /** The print run's laid pages (table fragments and their pages). */
  pages: readonly FlowPage[];
  /** The caret map's page-local box for a document position (null unmapped). */
  boxOf(pos: number): { page: number; yPx: number; heightPx: number } | null;
}

/** Everything the exporter's structure options need, derived in one pass. */
export interface PdfStructure {
  outline: PdfOutlineItem[];
  pageLabels: PdfPageLabelRange[];
  destinations: Record<string, PdfDestination>;
  structElements: PdfStructElement[];
}

/** A document-order heading with the page destination it resolves to. */
interface HeadingEntry {
  level: number;
  title: string;
  dest: PdfOutlineItem["dest"];
}

/** One outline node under construction (children materialize on push). */
interface OutlineNode {
  title: string;
  dest: PdfOutlineItem["dest"];
  children?: OutlineNode[];
}

/** The section's page height in CSS px for a page (undefined when the page
 *  index or the section's flow is missing). */
function pageHeightPxOf(view: PdfStructureView, page: number): number | undefined {
  const section = view.sectionOfPage[page] ?? 0;
  const height = view.sections[section]?.flow?.pageHeightPx;
  return height && height > 0 ? height : undefined;
}

/** The page-local box top as a /XYZ y in PDF points from the page bottom
 *  (undefined when the page's geometry is unknown). */
function pdfTopOf(view: PdfStructureView, box: { page: number; yPx: number }): number | undefined {
  const heightPx = pageHeightPxOf(view, box.page);
  if (heightPx === undefined) return undefined;
  const top = (heightPx - box.yPx) * 0.75;
  return Math.max(0, Math.min(heightPx * 0.75, top));
}

/** The document's heading bookmarks as a nested outline — Word's "Create
 *  bookmarks using: Headings" shape: single lines of paragraph text at
 *  Heading 1–9 / outline levels, nested by level (a level jump parents to the
 *  nearest shallower heading). Headings whose page cannot be resolved (no
 *  caret-map box: a header/footer story, a not-yet-laid position) are skipped
 *  rather than pointed at a guessed page; TOC fields' cached entry paragraphs
 *  are excluded like the navigation outline does. */
export function buildPdfOutline(doc: PMNode, view: PdfStructureView): PdfOutlineItem[] {
  const styles = (doc.attrs as { styles?: StylesOptions }).styles;
  const headings: HeadingEntry[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "tocField") return false;
    if (node.type.name !== "paragraph") return true;
    const attrs = node.attrs as { heading?: string; style?: string; outlineLevel?: number };
    const level = detectHeadingLevel(
      {
        heading: attrs.heading || undefined,
        style: attrs.style || undefined,
        outlineLevel: attrs.outlineLevel,
      },
      styles,
    );
    if (level == null) return true;
    const title = node.textContent.trim();
    if (!title) return true;
    // Inside the paragraph: its start position sits on the block boundary the
    // caret map rejects (see the navigation outline's pos + 1).
    const box = view.boxOf(pos + 1);
    if (!box) return true;
    const top = pdfTopOf(view, box);
    headings.push({
      level,
      title,
      dest: { pageIndex: box.page, ...(top !== undefined ? { top } : {}) },
    });
    return true;
  });

  const roots: OutlineNode[] = [];
  const stack: { level: number; node: OutlineNode }[] = [];
  for (const heading of headings) {
    const node: OutlineNode = { title: heading.title, dest: heading.dest };
    while (stack.length > 0 && stack[stack.length - 1]!.level >= heading.level) stack.pop();
    const parent = stack[stack.length - 1];
    if (parent) {
      parent.node.children = [...(parent.node.children ?? []), node];
    } else {
      roots.push(node);
    }
    stack.push({ level: heading.level, node });
  }
  return roots;
}

/** A w:numFmt token as the closest PDF page-label style. Unsupported formats
 *  (chineseCounting, ordinal, bullets, …) degrade to decimal — PDF has no
 *  glyph-run label styles. */
function pageLabelStyle(format: string | undefined): NonNullable<PdfPageLabelRange["style"]> {
  switch (format) {
    case "lowerRoman":
      return "romanLower";
    case "upperRoman":
      return "romanUpper";
    case "lowerLetter":
      return "alphaLower";
    case "upperLetter":
      return "alphaUpper";
    case "none":
      return "none";
    default:
      return "decimal";
  }
}

/** One page-label range per section, keyed to the section's first physical
 *  page and carrying the number that page displays (w:start restarts and
 *  continuation offsets both resolve through computePageNumberOffsets — the
 *  same arithmetic the PAGE field paints). A continuous section that opens on
 *  the previous section's last page replaces that page's range, so every page
 *  index appears at most once. */
export function buildPdfPageLabels(
  sections: PdfStructureView["sections"],
  sectionOfPage: readonly number[],
): PdfPageLabelRange[] {
  const offsets = computePageNumberOffsets(sections, sectionOfPage);
  const firstPageOf = new Map<number, number>();
  for (let page = 0; page < sectionOfPage.length; page++) {
    const section = sectionOfPage[page]!;
    if (!firstPageOf.has(section)) firstPageOf.set(section, page);
  }
  const ranges = new Map<number, PdfPageLabelRange>();
  for (const [section, firstPage] of [...firstPageOf].sort((a, b) => a[1] - b[1])) {
    const startNumber = firstPage + 1 + (offsets[section] ?? 0);
    ranges.set(firstPage, {
      startPageIndex: firstPage,
      style: pageLabelStyle(sections[section]?.pageNumbering?.format),
      ...(startNumber !== 1 ? { startNumber } : {}),
    });
  }
  return [...ranges.values()];
}

/** Every w:bookmarkStart name → its page destination (first occurrence wins —
 *  the same rule PAGEREF resolves by). These are the targets internal
 *  `#anchor` links (TOC entries, cross-references, hyperlinks to bookmarks)
 *  arrive with from the layout's link annotations. */
export function buildPdfDestinations(
  doc: PMNode,
  view: PdfStructureView,
): Record<string, PdfDestination> {
  const destinations: Record<string, PdfDestination> = {};
  doc.descendants((node, pos) => {
    if (node.type.name !== "inlinePassthrough" && node.type.name !== "passthrough") return true;
    const data = node.attrs?.data;
    if (typeof data !== "string") return true;
    let start: { name?: unknown } | undefined;
    try {
      start = (JSON.parse(data) as { bookmarkStart?: { name?: unknown } }).bookmarkStart;
    } catch {
      return true;
    }
    if (!start || typeof start.name !== "string" || start.name === "") return true;
    if (destinations[start.name]) return true;
    const box = view.boxOf(pos);
    if (!box) return true;
    const top = pdfTopOf(view, box);
    destinations[start.name] = { pageIndex: box.page, ...(top !== undefined ? { y: top } : {}) };
    return true;
  });
  return destinations;
}

/** The alt text + name a drawing node carries in the editor model. The image
 *  node's `title` is the docPr description (Word's alt text) and `alt` the
 *  docPr name/title; the shape payloads carry `descr` (description) and
 *  `title` (name); a chart's payload `title` is its rendered chart title,
 *  which is not alt text. */
function altTextOf(
  kind: string,
  attrs: Record<string, unknown>,
): { altText?: string; title?: string } | null {
  const str = (value: unknown): string | undefined =>
    typeof value === "string" && value !== "" ? value : undefined;
  switch (kind) {
    case "image":
      return { altText: str(attrs.title) ?? str(attrs.alt), title: str(attrs.alt) };
    case "model3d":
    case "ink":
      return { altText: str(attrs.descr) ?? str(attrs.altText), title: str(attrs.title) };
    case "wpsShape":
    case "wpgGroup": {
      const payload = (attrs[kind] ?? {}) as Record<string, unknown>;
      return { altText: str(payload.descr) ?? str(payload.altText), title: str(payload.title) };
    }
    case "chart": {
      const payload = (attrs.chart ?? {}) as Record<string, unknown>;
      return { altText: str(payload.descr) ?? str(payload.altText) };
    }
    default:
      return null;
  }
}

/** Tagged-structure elements the document's drawings and tables contribute:
 *  every picture/shape/3D/ink/chart node as a Figure with its alt text, and
 *  every laid table fragment (nested ones included) as a Table. Page indexes
 *  come from the same caret map the outline reads; a drawing whose page is
 *  unmappable is skipped with it. */
export function buildPdfStructElements(doc: PMNode, view: PdfStructureView): PdfStructElement[] {
  const elements: PdfStructElement[] = [];
  doc.descendants((node, pos) => {
    const info = altTextOf(node.type.name, node.attrs as Record<string, unknown>);
    if (!info) return true;
    const box = view.boxOf(pos);
    if (!box) return true;
    elements.push({
      type: "Figure",
      pageIndex: box.page,
      ...(info.altText ? { altText: info.altText } : {}),
      ...(info.title ? { title: info.title } : {}),
    });
    return true;
  });

  const addTables = (block: LaidOutBlock, pageIndex: number): void => {
    if (block.kind === "table") {
      elements.push({ type: "Table", pageIndex });
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const item of cell.stack) addTables(item.block, pageIndex);
        }
      }
      return;
    }
    if (block.kind === "group") {
      for (const item of block.children) addTables(item.block, pageIndex);
    }
  };
  view.pages.forEach((page, pageIndex) => {
    for (const item of page.items) addTables(item.block, pageIndex);
  });
  return elements;
}

/** Derive every structure option the PDF writer accepts from the print run. */
export function buildPdfStructure(view: PdfStructureView): PdfStructure {
  const doc = view.doc;
  if (!doc) {
    return { outline: [], pageLabels: [], destinations: {}, structElements: [] };
  }
  return {
    outline: buildPdfOutline(doc, view),
    pageLabels: buildPdfPageLabels(view.sections, view.sectionOfPage),
    destinations: buildPdfDestinations(doc, view),
    structElements: buildPdfStructElements(doc, view),
  };
}
