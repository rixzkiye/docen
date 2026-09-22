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
 *
 * The headless server path (`renderPdf`) has no PM document: it passes the
 * same view with `doc: undefined` and every derivation reads the LAID model —
 * the projection resolves each paragraph's heading level (style cascade +
 * outline level), carries each w:bookmarkStart as paragraph metadata (name +
 * inline slot) and each picture's docPr alt text/name onto the laid atoms and
 * drawings; `locateBlocks` maps every laid block to its page and page-local
 * top, and outline nesting, bookmark destinations and tagged figures/tables
 * all resolve from there with the same first-occurrence/page rules as the PM
 * walk.
 *
 * Exclusions & Standards Decisions (R12-W2):
 * - `viewerPreferences`: OOXML/Word DOCX has no document-level source for PDF
 *   viewer window controls (HideToolbar, HideMenubar, FitWindow, CenterWindow).
 *   Hence, they are not derived from DOCX models (written exclusion).
 *   DisplayDocTitle is reserved for PDF/UA-1 accessibility conformance in Q1.
 * - Exotic `numFmt`: ISO 32000-1 §12.4.2 only defines five numbering styles
 *   (/D, /R, /r, /A, /a). Unsupported formats fall back to decimal (/D).
 */

import { detectHeadingLevel, type StylesOptions } from "@docen/docx";
import {
  computePageNumberOffsets,
  type FlowPage,
  type LaidOutBlock,
  type LaidOutParagraph,
  type LayoutDrawingMember,
  type ProjectedPageNumbering,
} from "@docen/layout";
import type { Node as PMNode } from "@tiptap/pm/model";

import type {
  PdfDestination,
  PdfOutlineItem,
  PdfPageLabelRange,
  PdfStructElement,
} from "./export-pdf";

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

  return nestOutline(headings);
}

/** Nest document-order headings into Word's "Create bookmarks using:
 *  Headings" shape: single lines of paragraph text nested by level (a level
 *  jump parents to the nearest shallower heading). Shared by the PM walk and
 *  the laid-model server path so both outlines nest identically. */
function nestOutline(headings: readonly HeadingEntry[]): PdfOutlineItem[] {
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

/** The document-order text of a projected paragraph — the title the outline
 *  paints. Synthesized paint content (a numbering bullet's glyph) has no
 *  document-model characters behind it and is skipped, matching the PM walk's
 *  textContent. */
function projectedParagraphText(block: LaidOutParagraph): string {
  let text = "";
  for (const inline of block.inline) {
    if (inline.kind !== "text") continue;
    if (inline.synthetic || inline.suppressed) continue;
    text += inline.text;
  }
  return text.trim();
}

/** Every laid block on the print run with its page-local box top: page items
 *  first, nested table/group children composed onto their container's origin
 *  (the same offsets the PDF text layer walks). The first occurrence wins. */
function locateBlocks(view: PdfStructureView): Map<LaidOutBlock, { page: number; yPx: number }> {
  const located = new Map<LaidOutBlock, { page: number; yPx: number }>();
  const locate = (block: LaidOutBlock, page: number, yPx: number): void => {
    if (!located.has(block)) located.set(block, { page, yPx });
    if (block.kind === "table") {
      let rowY = yPx;
      for (const row of block.rows) {
        for (const cell of row.cells) {
          const top = rowY + (cell.insets.top ?? 0) + (cell.contentOffsetYPx ?? 0);
          for (const item of cell.stack) locate(item.block, page, top + item.yPx);
        }
        rowY += row.heightPx;
      }
      return;
    }
    if (block.kind === "group") {
      for (const child of block.children) locate(child.block, page, yPx + child.yPx);
    }
  };
  view.pages.forEach((page, pageIndex) => {
    for (const item of page.items) locate(item.block, pageIndex, item.yPx);
  });
  return located;
}

/** The document's heading bookmarks from the LAID model — the server path's
 *  outline when the view has no PM document (renderPdf's projection-only
 *  pass). Every laid paragraph carrying a resolved heading level (the
 *  projection resolved the style cascade) becomes an entry in page/document
 *  order, nested exactly like the PM walk; a heading whose page geometry is
 *  unknown still points at its page without a /XYZ top. */
export function buildPdfOutlineFromPages(view: PdfStructureView): PdfOutlineItem[] {
  const headings: HeadingEntry[] = [];
  for (const [block, at] of locateBlocks(view)) {
    if (block.kind !== "paragraph" || block.headingLevel == null) continue;
    const title = projectedParagraphText(block);
    if (!title) continue;
    const top = pdfTopOf(view, at);
    headings.push({
      level: block.headingLevel,
      title,
      dest: { pageIndex: at.page, ...(top !== undefined ? { top } : {}) },
    });
  }
  return nestOutline(headings);
}

/**
 * Map a w:numFmt token to the closest PDF page-label style (/S).
 *
 * Per ISO 32000-1 §12.4.2 (Table 159 — Entries in a page label dictionary),
 * the PDF standard specifies only five page label numbering styles:
 *   - /D: Decimal arabic numerals (1, 2, 3...)
 *   - /R: Uppercase roman numerals (I, II, III...)
 *   - /r: Lowercase roman numerals (i, ii, iii...)
 *   - /A: Uppercase letters (A, B, C...)
 *   - /a: Lowercase letters (a, b, c...)
 *
 * Decision (R12-W2): OOXML supports numerous exotic numbering schemes (such
 * as chineseCounting, chineseLegalSimplified, ordinal, cardinalText, aiueo,
 * iroha, katakana, etc.). The PDF standard provides no glyph-run label
 * numbering generator for exotic schemas inside /PageLabels; PDF readers
 * natively support only the five ISO 32000-1 styles. Therefore, all exotic
 * and unsupported formats gracefully fall back to "decimal" (/D), or "none"
 * (unnumbered) when explicitly specified as "none".
 */
export function pageLabelStyle(
  format: string | undefined,
): NonNullable<PdfPageLabelRange["style"]> {
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
 *  index appears at most once. Emits optional prefix string (w:chapStyle /
 *  w:sep / explicit prefix) as /P in the page label dictionary. */
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
    const pageNumbering = sections[section]?.pageNumbering;
    const startNumber = firstPage + 1 + (offsets[section] ?? 0);
    ranges.set(firstPage, {
      startPageIndex: firstPage,
      style: pageLabelStyle(pageNumbering?.format),
      ...(pageNumbering?.prefix ? { prefix: pageNumbering.prefix } : {}),
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

/** The paragraph-local top of the line covering an inline slot — the line the
 *  bookmark's zero-width marker anchors to. Falls back to the paragraph top
 *  when no line owns the slot (an empty paragraph, or a marker past the
 *  content). */
function bookmarkLineY(block: LaidOutParagraph, inlineIndex: number): number {
  const lines = block.lines;
  if (lines.length === 0) return 0;
  for (const line of lines) {
    if (line.items.some((item) => item.inlineIndex === inlineIndex)) return line.yPx;
  }
  const line =
    lines.find((candidate) => inlineIndex <= candidate.endInlineIndex) ?? lines[lines.length - 1]!;
  return line.yPx;
}

/** Bookmark names → page destinations from the LAID model — the server
 *  path's named destinations. The projection carries every w:bookmarkStart as
 *  paragraph metadata (name + inline slot), the laid line covering that slot
 *  supplies the /XYZ top, and the first occurrence wins exactly like the PM
 *  walk. Names are the document's own bookmark names — never guessed from
 *  heading text. */
export function buildPdfDestinationsFromPages(
  view: PdfStructureView,
): Record<string, PdfDestination> {
  const destinations: Record<string, PdfDestination> = {};
  for (const [block, at] of locateBlocks(view)) {
    if (block.kind !== "paragraph" || !block.bookmarks) continue;
    for (const bookmark of block.bookmarks) {
      if (destinations[bookmark.name]) continue;
      const top = pdfTopOf(view, {
        page: at.page,
        yPx: at.yPx + bookmarkLineY(block, bookmark.inlineIndex),
      });
      destinations[bookmark.name] = {
        pageIndex: at.page,
        ...(top !== undefined ? { y: top } : {}),
      };
    }
  }
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

  collectTableStructElements(view, elements);
  return elements;
}

/** The alt text a drawing member carries (picture docPr name/description,
 *  3D/ink payloads) — undefined for members with no accessibility metadata. */
function memberAltTextOf(
  member: LayoutDrawingMember,
): { altText?: string; title?: string } | undefined {
  if (member.kind !== "picture" && member.kind !== "model3d" && member.kind !== "ink") {
    return undefined;
  }
  if (!member.altText && !member.title) return undefined;
  return {
    ...(member.altText ? { altText: member.altText } : {}),
    ...(member.title ? { title: member.title } : {}),
  };
}

function pushFigure(
  elements: PdfStructElement[],
  pageIndex: number,
  info: { altText?: string; title?: string } | undefined,
): void {
  if (!info?.altText && !info?.title) return;
  elements.push({
    type: "Figure",
    pageIndex,
    ...(info.altText ? { altText: info.altText } : {}),
    ...(info.title ? { title: info.title } : {}),
  });
}

/** Figures + tables the laid model contributes — the doc-less server path's
 *  tagged structure. Every inline picture atom and drawing member carrying
 *  docPr alt text/name becomes a Figure on its laid page (a drawing's own alt
 *  wins over its members', matching the editor path's one-figure-per-node
 *  walk); every laid table fragment becomes a Table. */
export function buildPdfLaidStructElements(view: PdfStructureView): PdfStructElement[] {
  const elements: PdfStructElement[] = [];
  for (const [block, at] of locateBlocks(view)) {
    if (block.kind !== "paragraph") continue;
    for (const inline of block.inline) {
      if (inline.kind !== "picture") continue;
      if (inline.altText || inline.title) {
        pushFigure(elements, at.page, { altText: inline.altText, title: inline.title });
        continue;
      }
      for (const member of inline.members ?? []) {
        pushFigure(elements, at.page, memberAltTextOf(member));
      }
    }
    for (const drawing of block.drawings ?? []) {
      if (drawing.altText || drawing.title) {
        pushFigure(elements, at.page, { altText: drawing.altText, title: drawing.title });
        continue;
      }
      for (const member of drawing.members) {
        pushFigure(elements, at.page, memberAltTextOf(member));
      }
    }
  }
  collectTableStructElements(view, elements);
  return elements;
}

function collectTableStructElements(view: PdfStructureView, elements: PdfStructElement[]): void {
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
}

/** Derive every structure option the PDF writer accepts from the print run.
 *  The editor path supplies the PM document (outline from live headings,
 *  bookmark destinations, alt-text figures); the server path passes none and
 *  derives all three from the laid model instead — the projection carries
 *  heading levels, w:bookmarkStart markers and picture docPr alt text onto
 *  the laid blocks, so the same derivation runs without a PM node. */
export function buildPdfStructure(view: PdfStructureView): PdfStructure {
  const doc = view.doc;
  return {
    outline: doc ? buildPdfOutline(doc, view) : buildPdfOutlineFromPages(view),
    pageLabels: buildPdfPageLabels(view.sections, view.sectionOfPage),
    destinations: doc ? buildPdfDestinations(doc, view) : buildPdfDestinationsFromPages(view),
    structElements: doc ? buildPdfStructElements(doc, view) : buildPdfLaidStructElements(view),
  };
}
