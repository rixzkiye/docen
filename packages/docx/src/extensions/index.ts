export * from "./extensions";
// HTML paste input: parsed-HTML body → Tiptap JSON via the schema's rules.
export { parseHTMLBody } from "./paste";
// Word's highlighter palette (ST_HighlightColor tokens + their RGB) — the
// editor's highlight color picker renders from it (off-palette colors are
// illegal in w:highlight).
export { HIGHLIGHT_PALETTE_RGB, HighlightColor } from "@office-open/docx";
// Section geometry helpers, shared with the editor's page geometry
// (packages/docx layout/project.ts consumes resolvePageSize internally).
// DOCEN_CLIP_MIME + selectionSlicePayload: the docen-lossless clipboard lane
// (a PM slice JSON payload that survives copy/cut → paste with all marks).
export {
  resolvePageSize,
  resolveFontName,
  DOCEN_CLIP_MIME,
  selectionSlicePayload,
  SECTION_ATTR_KEYS,
} from "./utils";
export { renderDocx as renderParagraphDocx, parseDocx as parseParagraphDocx } from "./paragraph";
export * from "./vml-promotion";
export * from "./drawing-3d-ink";
