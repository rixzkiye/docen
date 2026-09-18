import type { SectionPropertiesOptions } from "@office-open/docx";

/**
 * docen-owned page-geometry defaults.
 *
 * `@office-open/docx` ships its own `sectionMarginDefaults` (top/bottom 1440,
 * left/right 1800 twips) and `sectionPageSizeDefaults` (A4) — the MS Office
 * zh-CN "Normal" preset. Letting those fill an absent sectPr silently makes
 * every generated/opened document inherit locale-specific geometry (1.25"
 * side margins, 851/992 header/footer distances) that Word does not agree
 * with. Generation paths therefore always write our own explicit
 * sectionProperties; nothing reaches office-open without one.
 *
 * The values are Word's international A4 "Normal" preset: A4 paper, 1"
 * (1440 twip) margins on all four sides, 0.5" (720 twip) header/footer
 * distance. They are deliberately locale-neutral — no zh-CN 1.25" sides, no
 * hand-me-down engine constants.
 */
export const DOCEN_DEFAULT_PAGE_SIZE = {
  WIDTH: 11906,
  HEIGHT: 16838,
} as const;

export const DOCEN_DEFAULT_PAGE_MARGIN = {
  TOP: 1440,
  RIGHT: 1440,
  BOTTOM: 1440,
  LEFT: 1440,
  HEADER: 720,
  FOOTER: 720,
  GUTTER: 0,
} as const;

/** A fresh SectionPropertiesOptions carrying the docen defaults — every
 *  section we generate gets an explicit page size and margin, so office-open
 *  never falls back to its zh-CN `sectionMarginDefaults`. Returned as a new
 *  object each call: callers may patch/merge it without cross-document
 *  aliasing. */
export function docenDefaultSectionProperties(): SectionPropertiesOptions {
  return {
    pageSize: { width: DOCEN_DEFAULT_PAGE_SIZE.WIDTH, height: DOCEN_DEFAULT_PAGE_SIZE.HEIGHT },
    pageMargin: {
      top: DOCEN_DEFAULT_PAGE_MARGIN.TOP,
      right: DOCEN_DEFAULT_PAGE_MARGIN.RIGHT,
      bottom: DOCEN_DEFAULT_PAGE_MARGIN.BOTTOM,
      left: DOCEN_DEFAULT_PAGE_MARGIN.LEFT,
      header: DOCEN_DEFAULT_PAGE_MARGIN.HEADER,
      footer: DOCEN_DEFAULT_PAGE_MARGIN.FOOTER,
      gutter: DOCEN_DEFAULT_PAGE_MARGIN.GUTTER,
    },
  };
}
