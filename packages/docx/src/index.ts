/**
 * @docen/docx — DOCX editor and converter powered by @office-open/docx.
 *
 * @module
 */

// Core: @tiptap/core re-exports + DOCX extension registry
export { docxExtensions, type JSONContent, type AnyExtension } from "./core";

// Re-export the OOXML section-properties type (page size/margin/orientation +
// document grid) so the editor layer can type section geometry without a direct
// @office-open/docx dependency.
export type { SectionPropertiesOptions } from "@office-open/docx";
// Glossary / building-blocks model (word/glossary/document.xml): the parsed
// part shape plus the gallery/type/behavior token enums. office-open parses
// and stringifies the part natively, so real Word Quick Parts round-trip
// through documentExtras.glossary — same no-direct-dependency rationale.
export {
  DocPartBehavior,
  DocPartGallery,
  DocPartType,
  type DocPartOptions,
  type DocPartSectionOptions,
  type GlossaryDocumentOptions,
} from "@office-open/docx";
// Re-export the chart payload type: the chart node's attrs.chart IS a
// ChartOptions verbatim (extensions/chart.ts), and the editor's command layer
// edits it as one — same rationale as SectionPropertiesOptions above. The
// @office-open/docx barrel doesn't re-export the core chart value-domain
// names, so they come straight from @office-open/core (already a dependency).
import type { ChartType as OfficeChartType } from "@office-open/core";
export type ChartType = OfficeChartType | "combo";
export type { ChartOptions } from "@office-open/docx";
export type { ChartSeriesData, LegendPosition } from "@office-open/core";
// docen-owned section-geometry defaults — generation stamps these into every
// sectPr (never office-open's zh-CN `sectionMarginDefaults`), and editor-side
// geometry fallbacks — content-width for image capping, page measurement —
// reuse the SAME constants instead of hardcoding divergent ones.
export {
  DOCEN_DEFAULT_PAGE_SIZE,
  DOCEN_DEFAULT_PAGE_MARGIN,
  docenDefaultSectionProperties,
} from "./converters/section-defaults";
// Re-export the engine's length conversion (mm → twips) so the editor layer can
// build OOXML page geometry from mm presets without a direct @office-open/core
// dependency. (@office-open/docx does not re-export this from core.) Sourced from
// the `util` subpath so bundlers can tree-shake the rest of core.
export { convertMillimetersToTwip } from "@office-open/core/util";

// Editor factory
export { createDocxEditor, type DocxEditorOptions } from "./editor";

// Extensions
export * from "./extensions";

// Converters: DOCX pipeline (DOCX binary ↔ Tiptap JSON)
export {
  parseDOCX,
  parseDOCXSync,
  generateDOCX,
  generateDOCXSync,
  generateDOCXStream,
  resolveDocument,
  compileDocument,
  normalizeDocument,
  DocxManager,
  formatMarkNames,
  runPropsFromMarks,
  runPropsToMarks,
  createCompileCache,
  type CompileCache,
  type DocxGenerateOptions,
  type DocxVariant,
  type RunPropMark,
} from "./converters/docx";
export { decodePassthroughData, encodePassthroughData } from "./extensions/passthrough";
// Generated-field cache pipeline (SEQ/REF/page/TOC caches at generation time)
// plus the pure instruction/format helpers the editor's update commands share.
export {
  fillGeneratedFields,
  parseFieldInstruction,
  fieldRef,
  parseCustomStyles,
  styleLevelsOf,
  seqLabelOfData,
  formatSeqNumber,
  seqChapterLevel,
  SEQ_NUMBER_FORMATS,
  CAPTION_SEPARATOR_CHARS,
  type ParsedFieldInstruction,
  type FieldRef,
  type FieldCacheOptions,
} from "./converters/field-eval";
export { DOCX_EPOCH } from "./converters/determinism";
export {
  ENCRYPTED_DOCUMENT_CODE,
  EncryptedDocumentError,
  isEncryptedContainerBytes,
} from "./converters/encrypted";
export {
  UNSUPPORTED_CONTENT_KINDS,
  detectUnsupportedContent,
  type UnsupportedContentItem,
  type UnsupportedContentKind,
  type UnsupportedContentReport,
} from "./converters/unsupported";
export {
  type TocSwitches,
  tokenizeTocInstruction,
  parseTocSwitches,
  generateTocInstruction,
} from "./converters/toc-switches";

// Converters: DOCX template patching (placeholder replacement via office-open patchDocument)
export { patchDOCX, type DocxPatchOptions, type DocxPatchContent } from "./converters/patch";

// Converters: Document prepare pipeline (pre-process before compile)
export {
  prepareDocument,
  prepareImages,
  prepareImageSizes,
  fetchImageHandler,
  DEFAULT_IMAGE_MAX_BYTES,
  DEFAULT_IMAGE_MAX_REDIRECTS,
  DEFAULT_IMAGE_TIMEOUT_MS,
  type PrepareImagesPolicy,
  type FetchImageOptions,
  type PrepareStep,
  type ImageFetchHandler,
} from "./converters/prepare";

// Converters: Markdown pipeline (Markdown string ↔ Tiptap JSON)
export { parseMarkdown, generateMarkdown } from "./converters/markdown";

// Converters: RTF pipeline (RTF string ↔ Tiptap JSON)
export { parseRTF, generateRTF } from "./converters/rtf";

// Converters: HTML pipeline (HTML string ↔ Tiptap JSON)
export { parseHTML, generateHTML, type HtmlGenerateOptions } from "./converters/html";

// Converters: Plain Text pipeline (Plain text ↔ Tiptap JSON)
export { parsePlainText, generatePlainText } from "./converters/text";

// Converters: ODT pipeline (ODT zip package from Tiptap JSON)
export { generateODT } from "./converters/odt";

// Style-facing editor helpers (Quick Styles gallery entries, caret run props)
export {
  quickStyles,
  effectiveRunProps,
  type QuickStyleEntry,
  type StylesOptions,
} from "./converters/styles";

// Style-inheritance primitives (styles.xml model → resolved cascade), shared
// by the layout projection, the CSS route, and the editor's paginator
// (measure.ts) so all of them resolve the SAME effective properties: a
// paragraph whose direct attrs are empty still inherits its style's (and its
// basedOn chain's) spacing/indent/run.
export {
  indexParagraphStyles,
  defaultParagraphStyleId,
  mergeStyleChain,
  deepMergeInto,
  mergeTableStyleProps,
  resolveTableLook,
  resolveTableStyle,
  resolveTableCellStyle,
  activeConditionalTypes,
  CONDITIONAL_FORMAT_PRIORITY,
  type StyleEntry,
  type ResolvedTableLook,
  type ResolvedTableStyle,
  type TableCellPosition,
  type EffectiveTableCellStyle,
} from "./style-cascade";

// Numbering (list) level index — the same reference → levels table the layout
// projection resolves markers against, for editor-side list-number probes
// (the cross-reference commit's `\n` paragraph numbers).
export { indexNumberings, type NumberingLevel } from "./layout/project/numbering";

// Preset geometry evaluator (ECMA-376 prstGeom → SVG path data), shared by the
// layout projection (non-box shapes become path members) and the editor's
// shapes gallery (live SVG previews) so both render from one evaluator. The
// evaluator itself is format-neutral and lives in @docen/core.
export { presetShapePaths, type PresetShapeOutline } from "@docen/core/geometry";

// Font embedding (ECMA-376 Part 4 §14.2.14 obfuscation & fontTable)
export * from "./font-embedding";

// Types
export type * from "./types";
