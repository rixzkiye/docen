import type {
  ParagraphChild,
  ParagraphOptions,
  SectionChild,
  ShapeOptions,
} from "@office-open/docx";

import { cleanAttrs } from "../converters/styles";
import { Node } from "../core";
import type { JSONContent } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";

/**
 * wpsShape — inline node carrying a standalone DOCX text-box shape
 * (wp:anchor > wps:wsp > wps:txbx; NOT inside a wpg group). The shape geometry
 * + styling (transformation/floating/fill/outline/bodyProperties) ride on
 * attrs.wpsShape; the editable text body is PM content (block+), one paragraph
 * per office-open ParagraphOptions. Unlike a group's interior wps children
 * (laid out in the group's coordinate space), this shape floats on its own
 * anchor. The engine node is UI-free; rendering is owned by the editor layer.
 */

/** The standalone text-box shape ParagraphChild branch. */
type WpsBranch = Extract<ParagraphChild, { wpsShape: ShapeOptions }>;

/** Block tag keys — a ShapeTextBoxChild that is not a paragraph carries one of
 *  these (a nested block branch); a ParagraphOptions carries none. */
const BLOCK_TAGS = [
  "paragraph",
  "table",
  "toc",
  "textbox",
  "sdt",
  "altChunk",
  "customXml",
  "bookmarkStart",
  "bookmarkEnd",
  "rawXml",
] as const;

const isBlockBranch = (child: object): child is SectionChild =>
  BLOCK_TAGS.some((tag) => tag in child);

const attrWpsShape = () => ({
  default: null,
  rendered: false,
  parseHTML: (element: HTMLElement) => {
    const raw = element.getAttribute("data-wps-shape");
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  },
});

/**
 * Fold the editor's flat shape `name` into the OOXML slot
 * (ShapeCoreOptions.nonVisualProperties.name → wps:cNvSpPr/@name). The editor
 * model carries the name at the top level (the watermark gallery stamps
 * `WordPictureWatermark` there and `isWatermarkNode` reads it back); the
 * engine only writes it from nonVisualProperties. Compile therefore folds on
 * the way out and {@link resolveWpsShape} lifts it back on the way in, so the
 * name survives the round-trip instead of landing in no part at all.
 */
export function foldWpsShapeName(geometry: Record<string, unknown>): Record<string, unknown> {
  const { name, ...rest } = geometry;
  if (typeof name !== "string" || name === "") return geometry;
  const nvp = rest.nonVisualProperties;
  return {
    ...rest,
    nonVisualProperties: {
      ...(typeof nvp === "object" && nvp !== null ? (nvp as Record<string, unknown>) : {}),
      name,
    },
  };
}

/** ParagraphChild `{ wpsShape: {...} }` → wpsShape node. Mirrors the old
 *  DocxManager wpsShape branch: the shape's text body (children) becomes PM
 *  content (one node per paragraph); geometry/styling ride on attrs.wpsShape.
 *  Each paragraph's defRPr (para.run) is merged into its runs then dropped —
 *  it is the box's default run-properties, not the ¶-mark rPr (see inline
 *  note). Shared with group-members: a wpg group's wps child resolves through
 *  the same path (its child-space box becomes the transformation). */
export function resolveWpsShape(ws: WpsBranch["wpsShape"], ctx: ResolveContext): JSONContent {
  const content: JSONContent[] = [];
  if (ws?.children) {
    for (const child of ws.children) {
      if (typeof child !== "object" || child === null) {
        const node = ctx.resolveParagraph(child);
        if (node) content.push(node);
        continue;
      }
      // A text-box body is paragraphs; a stray nested block branch resolves
      // through the block stream instead of being forced through the
      // paragraph path.
      if (isBlockBranch(child)) {
        const node = ctx.resolveBlock(child);
        if (node) content.push(node);
        continue;
      }
      const para: ParagraphOptions = child;
      // DrawingML defRPr (para.run) is the default run-properties for the box's
      // runs, NOT the OOXML ¶-mark rPr. Merge it into each run, then drop it from the
      // paragraph (run: undefined): paragraph.ts renders attrs.run.size as
      // ¶-mark line-height, which would override the box's grid line-height —
      // but defRPr is a run default, not a ¶ mark. Round-trip safe — runs carry
      // the full rPr, so compile emits per-run rPr and Word renders identically.
      const defRPr = para.run ?? {};
      const children = Array.isArray(para.children)
        ? para.children.map((c) =>
            typeof c !== "object" || c === null ? { ...defRPr, text: c } : { ...defRPr, ...c },
          )
        : undefined;
      const node = ctx.resolveParagraph({
        ...para,
        run: undefined,
        ...(children ? { children } : {}),
      });
      if (node) content.push(node);
    }
  }
  if (content.length === 0) content.push({ type: "paragraph" });
  const { children: _omit, ...rawGeometry } = ws ?? {};
  // Lift the shape name from its OOXML slot (wps:cNvSpPr/@name is parsed onto
  // nonVisualProperties.name) to the editor-facing top level, where the shape
  // model and the watermark detection read it. Compile folds it back
  // (foldWpsShapeName), so the name survives the round-trip.
  const nvp = rawGeometry.nonVisualProperties;
  let geometry: Record<string, unknown> = rawGeometry as Record<string, unknown>;
  if (typeof nvp === "object" && nvp !== null) {
    const { name, ...restNvp } = nvp as Record<string, unknown>;
    if (typeof name === "string" && name !== "") {
      geometry = {
        ...rawGeometry,
        name,
        nonVisualProperties: Object.keys(restNvp).length > 0 ? restNvp : undefined,
      };
    }
  }
  const node: JSONContent = { type: "wpsShape", content };
  const cleanGeometry = cleanAttrs(geometry);
  if (Object.keys(cleanGeometry).length > 0) node.attrs = { wpsShape: cleanGeometry };
  return node;
}

// DOCX standalone text-box shape → office-open ParagraphChild `{ wpsShape }`.
export const parseDocxInline: ParseInlineRule<WpsBranch> = {
  match: (child): child is WpsBranch => "wpsShape" in child,
  convert: (child, ctx) => resolveWpsShape(child.wpsShape, ctx),
};

export const WpsShape = Node.create({
  name: "wpsShape",
  group: "inline",
  inline: true,
  // Editable text body (was atom). content:"block+" holds the textbox's
  // paragraph(s); isolating stops Backspace at the start from merging the first
  // paragraph back into the anchor paragraph; defining keeps the node when the
  // body is fully selected+replaced.
  content: "block+",
  isolating: true,
  defining: true,

  addAttributes() {
    return {
      wpsShape: attrWpsShape(),
    };
  },

  parseHTML() {
    return [
      {
        tag: "div[data-wps-shape]",
        // The editable body lives in the inner <div> (the contentDOM); parse it
        // from there (querySelector "div" = the first child div) instead of the
        // outer positioning wrapper, so paragraphs resolve as content rather
        // than getting hoisted out of the inline node.
        contentElement: "div",
      },
    ];
  },

  parseDocxInline,
});
