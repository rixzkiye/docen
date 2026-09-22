import type { ParagraphOptions, StylesOptions } from "@office-open/docx";
import { Node as TiptapNode } from "@tiptap/core";
import type { JSONContent } from "@tiptap/core";

import { indexParagraphStyles } from "../style-cascade";
import { HTML_ORDERED_TEMP } from "./list-numbering";
import { docxParagraphAttrs, SECTION_ATTR_KEYS } from "./utils";

/**
 * Paragraph extension with nested office-open attrs — the single textblock node.
 *
 * Attrs mirror ParagraphPropertiesOptionsBase verbatim (alignment/indent/
 * spacing/border/shading/frame as nested objects + every scalar OOXML property
 * + heading/style/bullet/numbering/thematicBreak): a heading IS a paragraph in
 * OOXML, so its HeadingLevel pStyle rides on the `heading` attr instead of a
 * separate node, and DOCX round-trip is near-identity — renderDocx/parseDocx
 * pass attrs through. Consumers derive the display level via detectHeadingLevel.
 */

/** HeadingLevel literals: "Heading1".."Heading9", "Title". */
export const HEADING_COMPILE_MAP: Record<number, string> = {
  1: "Heading1",
  2: "Heading2",
  3: "Heading3",
  4: "Heading4",
  5: "Heading5",
  6: "Heading6",
  7: "Heading7",
  8: "Heading8",
  9: "Heading9",
};

// Heading-level resolution lives in the TipTap-free `heading-level` module
// (the layout projection resolves headings too); re-exported here so the
// extension's existing importers keep their entry point.
export { detectHeadingLevel, HEADING_PARSE_MAP } from "../heading-level";

/** Every name a paragraph style is known by, lowercased: the style id itself
 *  plus each resolved name up the `basedOn` chain. A TOC `\t` switch lists
 *  style NAMES ("My Heading,2") while paragraph attrs carry the style id, so
 *  matching has to try both. Empty when `styles` carries no definitions. */
export function paragraphStyleNames(
  styles: StylesOptions | undefined,
  styleId: string | undefined,
): string[] {
  if (!styleId) return [];
  const names = new Set<string>([styleId.toLowerCase()]);
  if (!styles) return [...names];
  const byId = indexParagraphStyles(styles);
  const visited = new Set<string>();
  let curId: string | undefined = styleId;
  while (curId && !visited.has(curId)) {
    visited.add(curId);
    const style = byId.get(curId);
    if (!style) break;
    if (typeof style.name === "string" && style.name) names.add(style.name.toLowerCase());
    curId = style.basedOn ?? undefined;
  }
  return [...names];
}

// ── DOCX serialization (near-identity: attrs mirror ParagraphPropertiesOptionsBase) ──

/** office-open writes `w:hSpace`/`w:vSpace` from `frame.space.{horizontal,
 *  vertical}`. A drop cap's editable `distance`/`vDistance` and any legacy
 *  `hSpace`/`vSpace` keys on the frame attr are folded into that shape so the
 *  twip distances survive the writer. */
function frameForDocx(attrs: Record<string, unknown>): Record<string, unknown> | undefined {
  let frame: Record<string, unknown> | undefined;
  if (attrs.frame && typeof attrs.frame === "object") {
    frame = { ...(attrs.frame as Record<string, unknown>) };
    const legacyH = frame.hSpace;
    const legacyV = frame.vSpace;
    const space =
      frame.space && typeof frame.space === "object"
        ? { ...(frame.space as Record<string, unknown>) }
        : undefined;
    delete frame.hSpace;
    delete frame.vSpace;
    if (legacyH !== undefined || legacyV !== undefined || space) {
      frame.space = {
        horizontal: space?.horizontal ?? legacyH ?? 0,
        vertical: space?.vertical ?? legacyV ?? 0,
      };
    }
  }
  const dc = attrs.dropCap;
  if (dc && typeof dc === "object") {
    const d = dc as { val?: string; lines?: number; distance?: number; vDistance?: number };
    if (d.val && d.val !== "none") {
      return {
        ...frame,
        dropCap: d.val,
        lines: d.lines ?? 3,
        space: { horizontal: d.distance ?? 0, vertical: d.vDistance ?? 0 },
      };
    }
  }
  return frame;
}

export function renderDocx(node: JSONContent): Record<string, unknown> {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>;
  const opts: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (value === null || value === undefined) continue;
    if (SECTION_ATTR_KEYS.has(key)) continue;
    // Runtime-only attrs the TableOfContents extension injects on each heading
    // (id / data-toc-id) are regenerated on every load — never persist them.
    if (key === "id" || key === "data-toc-id") continue;
    // docen-only round-trip data with no OOXML paragraph counterpart (the
    // markdown code-fence info string) — keeps the JSON lossless but must not
    // reach ParagraphOptions.
    if (key === "codeLanguage" || key === "dropCap" || key === "frame") continue;
    opts[key] = value;
  }
  const frame = frameForDocx(attrs);
  if (frame) opts.frame = frame;
  return opts;
}

/**
 * Structural/semantic keys handled elsewhere (run/text children — `run` is
 * intentionally NOT skipped: ParagraphOptions.run is the paragraph's default
 * run properties, kept as an attr for lossless round-trip, e.g. header/footer
 * paragraphs whose styling lives there).
 */
const SKIP_KEYS = new Set(["children", "text"]);

const asNumber = (v: unknown): number | undefined => {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v))) return Number(v);
  return undefined;
};

export function parseDocx(opts: ParagraphOptions | string): Record<string, unknown> {
  const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;
  const attrs: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(resolved)) {
    if (SKIP_KEYS.has(key)) continue;
    attrs[key] = value ?? null;
  }
  const frame =
    (resolved.frame as Record<string, unknown> | undefined) ??
    ((resolved as Record<string, unknown>).framePr as Record<string, unknown> | undefined);
  if (frame && frame.dropCap && !attrs.dropCap) {
    const space =
      frame.space && typeof frame.space === "object"
        ? (frame.space as Record<string, unknown>)
        : undefined;
    attrs.dropCap = {
      val: frame.dropCap,
      lines: asNumber(frame.lines) ?? 3,
      distance: asNumber(space?.horizontal) ?? asNumber(frame.hSpace) ?? 0,
      vDistance: asNumber(space?.vertical) ?? asNumber(frame.vSpace) ?? 0,
    };
  }
  return attrs;
}

// ── Extension ──

export const Paragraph = TiptapNode.create({
  name: "paragraph",
  group: "block",
  content: "inline*",
  // A heading is a paragraph in OOXML (a <w:p> with pStyle="Heading1"), so the
  // paragraph node carries the full office-open mirror — heading/style/bullet/
  // numbering/thematicBreak included. See utils.
  addAttributes() {
    return docxParagraphAttrs();
  },

  // HTML input (clipboard paste): h1-h6 parse back with a lifted HeadingLevel
  // `heading` attr; a <h6 data-heading-level="N"> proxy restores levels 7-9.
  // The proxy rule runs before the native h6 rule so a plain <h6> maps to
  // level 6.
  // A <li> maps to a flat list paragraph: nesting depth = ancestor ul/ol
  // count, and the NEAREST list element decides bullet vs ordered (an <ol>
  // wrapping a nested <ul> marks that sublist's items as bullets). Ordered
  // items carry the html-ordered placeholder reference — parseHTML
  // (converters/html.ts) rewrites each consecutive run to a fresh
  // docen-ordered-* reference so independent lists number separately.
  // A <hr> maps to a thematic-break paragraph (OOXML has no HR element).
  parseHTML() {
    return [
      {
        tag: "hr",
        getAttrs: () => ({ thematicBreak: true }),
      },
      // A code block is a "Code"-styled paragraph (no dedicated OOXML
      // element). The inner <code> keeps the Code character mark;
      // preserveWhitespace keeps code newlines (the DOM parser collapses
      // them to spaces by default).
      { tag: "pre", attrs: { style: "Code" }, preserveWhitespace: "full" },
      {
        tag: "li",
        getAttrs: (el: HTMLElement) => {
          // Pasted li elements are flattened by parseHTMLBody (paste.ts):
          // level + kind ride on attributes.
          const flatLevel = el.getAttribute("data-docen-level");
          if (flatLevel != null && flatLevel !== "") {
            const level = Math.max(0, Number(flatLevel) || 0);
            return el.getAttribute("data-docen-kind") === "ol"
              ? { numbering: { reference: HTML_ORDERED_TEMP, level } }
              : { bullet: { level } };
          }
          let depth = 0;
          let nearest: string | null = null;
          for (let p = el.parentElement; p; p = p.parentElement) {
            const tag = p.tagName.toUpperCase();
            if (tag === "UL" || tag === "OL") {
              depth++;
              nearest ??= tag;
            }
          }
          // The nearest list is depth 1 → nesting level 0.
          const level = Math.max(0, depth - 1);
          return nearest === "OL"
            ? { numbering: { reference: HTML_ORDERED_TEMP, level } }
            : { bullet: { level } };
        },
      },
      {
        tag: "h6[data-heading-level]",
        getAttrs: (el: HTMLElement) => {
          const level = Number((el as HTMLElement).getAttribute("data-heading-level"));
          return {
            heading:
              HEADING_COMPILE_MAP[Number.isInteger(level) && level >= 1 && level <= 9 ? level : 6],
          };
        },
      },
      ...[1, 2, 3, 4, 5, 6].map((level) => ({
        tag: `h${level}`,
        attrs: { heading: HEADING_COMPILE_MAP[level] },
      })),
      { tag: "p" },
    ];
  },

  renderDocx,
  parseDocx,
});
