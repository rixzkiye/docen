// Heading-level resolution — the pure, renderer-agnostic half of the
// paragraph extension: which pStyle / outline level marks a heading. Kept
// free of TipTap so the layout projection (and any Node-only consumer) can
// resolve headings without pulling the editor graph in.

import type { StylesOptions } from "@office-open/docx";

import { indexParagraphStyles } from "./style-cascade";

/** HeadingLevel style-id literals: "Heading1".."Heading9", "Title". */
export const HEADING_PARSE_MAP: Record<string, number> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
  Heading7: 7,
  Heading8: 8,
  Heading9: 9,
  Title: 1,
};

/** Heading level (1-9) from a localized style NAME: "heading 1"/"标题 1" → 1,
 *  "title" → 1. office-open's built-in names are English ("heading 1"), but
 *  zh-CN Word labels the same styles "标题 1"; both map to the same level. */
function headingLevelFromName(name: string | undefined): number | undefined {
  if (!name) return undefined;
  const m = /^heading\s+(\d)$/i.exec(name) ?? /^标题\s*(\d)$/.exec(name);
  if (m) {
    const lvl = Number(m[1]);
    if (lvl >= 1 && lvl <= 9) return lvl;
  }
  return /^title$/i.test(name) ? 1 : undefined;
}

/** Heading level (1-9) for a paragraph, or undefined when it isn't a heading.
 *  DOCX marks a heading several ways, checked in priority order:
 *  1. office-open lifts a HeadingLevel pStyle ("Heading1".."Title") into `heading`.
 *  2. An explicit `outlineLevel` (0-8 → 1-9) — Word's outline/TOC key off this
 *     even without a heading pStyle; the Heading1-9 styles carry outlineLvl 0-8.
 *  3. A pStyle that names a heading style: directly ("Heading7", which stays on
 *     `style` because office-open's HeadingLevel type caps at 6), by localized
 *     NAME ("heading 1"/"标题 1"), or via the `basedOn` chain (a custom style
 *     "MyTitle" basedOn="Heading1"). `heading` and `style` carry the same pStyle.
 *  `resolved` accepts a full office-open ParagraphOptions (parse/compile) OR a
 *  PM-node attrs subset — the editor outline walks PM nodes whose style /
 *  outlineLevel mark a heading without a lifted `heading` attr (a numeric
 *  pStyle id common in WPS / Chinese Word). Pure (no `this`): resolved + the
 *  document styles snapshot are all it reads. */
export function detectHeadingLevel(
  resolved: { heading?: string; style?: string; outlineLevel?: number },
  styles: StylesOptions | undefined,
): number | undefined {
  if (resolved.heading) {
    const lvl = HEADING_PARSE_MAP[resolved.heading];
    if (lvl) return lvl;
  }
  const outline = resolved.outlineLevel;
  if (typeof outline === "number" && outline >= 0 && outline <= 8) {
    return outline + 1;
  }
  const styleId = resolved.style;
  if (!styleId || !styles) return undefined;
  const byId = indexParagraphStyles(styles);
  const visited = new Set<string>();
  let curId: string | undefined = styleId;
  while (curId && !visited.has(curId)) {
    visited.add(curId);
    if (HEADING_PARSE_MAP[curId]) return HEADING_PARSE_MAP[curId];
    const style = byId.get(curId);
    if (!style) break;
    const lvl = headingLevelFromName(style.name);
    if (lvl) return lvl;
    curId = style.basedOn ?? undefined;
  }
  return undefined;
}
