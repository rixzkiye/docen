import {
  createMeasurer,
  stackBlocks,
  type FontMetrics,
  type LaidOutStackItem,
  type LayoutBlock,
  type ProjectedColumns,
  type ProjectedFlowBox,
  type ProjectedLineNumbers,
  type ProjectedPageBorders,
  type ProjectedPageFurniture,
  type ProjectedPageNumbering,
  type FlowSection,
} from "@docen/layout";

export interface LaidFurnitureSlot {
  stack: readonly LaidOutStackItem[];
  heightPx: number;
}

/** One section's laid furniture slots — [default, first, even]. */
export interface LaidFurnitureSection {
  header: (LaidFurnitureSlot | undefined)[];
  footer: (LaidFurnitureSlot | undefined)[];
}

/** One section's paint inputs: the page geometry its pages paginate against
 *  and the headers/footers its pages display. */
export interface PdfStageSection {
  flow: ProjectedFlowBox;
  /** The section break type (sectPr @w:type) — drives continuous merging and
   *  even/odd blank interleaves in the flow; absent = nextPage. */
  type?: FlowSection["type"];
  /** The section's page borders (w:pgBorders), absent when none. */
  pageBorders?: ProjectedPageBorders;
  /** The section's line numbering (w:lnNumType), absent when none. */
  lineNumbers?: ProjectedLineNumbers;
  /** The section's page numbering (w:pgNumType), absent when the section
   *  continues the previous one's decimal numbers. */
  pageNumbering?: ProjectedPageNumbering;
  /** The section's columns (w:cols), absent for a single-column section. */
  columns?: ProjectedColumns;
  /** Headers/footers for this section's pages (absent = none). */
  furniture?: ProjectedPageFurniture;
  /** The slots of `furniture` laid out once (layFurnitureSections) — the
   *  page insets push the body by these heights and the painter draws these
   *  same stacks, so push-down == painted band height by construction. */
  furnitureLaid?: LaidFurnitureSection;
  blocks?: readonly LayoutBlock[];
  /** Footnote definitions (absent when document has no footnotes). */
  footnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
  /** Endnote definitions (absent when document has no endnotes). */
  endnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
}

export type CanvasStageSection = PdfStageSection;

const FURNITURE_SLOTS = [0, 1, 2] as const;

export function layFurnitureSections(
  sections: readonly PdfStageSection[],
  metrics: FontMetrics,
): (LaidFurnitureSection | undefined)[] {
  const measurer = createMeasurer(metrics);
  return sections.map((section) => {
    const f = section.furniture;
    if (!f) return undefined;
    const lay = (blocks: readonly LayoutBlock[] | undefined): LaidFurnitureSlot | undefined => {
      if (!blocks) return undefined;
      const laid = stackBlocks(blocks, section.flow.contentWidthPx, undefined, measurer);
      return { stack: laid.stack, heightPx: laid.heightPx };
    };
    return {
      header: FURNITURE_SLOTS.map((slot) => lay([f.header, f.firstHeader, f.evenHeader][slot])),
      footer: FURNITURE_SLOTS.map((slot) => lay([f.footer, f.firstFooter, f.evenFooter][slot])),
    };
  });
}

export function computePageInsets(
  flow: ProjectedFlowBox,
  furniture: ProjectedPageFurniture | undefined,
  laid: LaidFurnitureSection | undefined,
): import("@docen/layout").FlowPageInsets | undefined {
  if (!furniture) return undefined;
  const topMargin = flow.contentTopPx;
  const bottomMargin = flow.pageHeightPx - flow.contentTopPx - flow.contentHeightPx;
  const headerDistance = furniture.headerDistancePx ?? 48;
  const footerDistance = furniture.footerDistancePx ?? 48;
  const height = (kind: "header" | "footer", slot: 0 | 1 | 2): number | undefined =>
    laid?.[kind][slot]?.heightPx;
  const inset = (headerPx: number | undefined, footerPx: number | undefined) => {
    const top = Math.max(0, headerDistance + (headerPx ?? 0) - topMargin);
    const bottom = Math.max(0, footerDistance + (footerPx ?? 0) - bottomMargin);
    return top > 0 || bottom > 0
      ? { topPx: Math.round(top), bottomPx: Math.round(bottom) }
      : undefined;
  };
  const def = inset(height("header", 0), height("footer", 0));
  if (!def) return undefined;
  const out: import("@docen/layout").FlowPageInsets = { default: def };
  if (furniture.titlePage) {
    out.first =
      inset(
        height("header", 1) ?? height("header", 0),
        height("footer", 1) ?? height("footer", 0),
      ) ?? undefined;
  }
  if (furniture.evenAndOddHeaders) {
    out.even =
      inset(
        height("header", 2) ?? height("header", 0),
        height("footer", 2) ?? height("footer", 0),
      ) ?? undefined;
  }
  return out;
}
