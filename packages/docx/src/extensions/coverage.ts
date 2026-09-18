import type { ParagraphChild, RunOptions, SectionChild } from "@office-open/docx";

/**
 * Concept-coverage disposition tables — the explicit registry of how EVERY
 * office-open document-model branch is handled by the Tiptap layer.
 *
 * Covered direction (resolve, office-open → Tiptap): every SectionChild /
 * ParagraphChild branch is claimed by an editable node, carried verbatim
 * through a passthrough atom, or dropped for a logged reason. Every node now
 * maps back on compile (paragraph attrs mirror ParagraphPropertiesOptions
 * verbatim), so there is no compile-loss table.
 *
 * The satisfies guards make the registry complete by construction: a new
 * office-open union branch widens the tag type, a missing entry fails the
 * build; an entry naming no real branch also fails. This file is consumed by
 * coverage.spec.ts, which drives round-trip fixtures per tag — the registry
 * claims, the spec proves.
 */

/** Extract the tag key of every single-key branch member of a union. MUST be
 *  distributive (naked `U extends object`): a bare `keyof UnionType` would be
 *  the INTERSECTION of member keys — never, for tag unions with no shared key —
 *  and Record<never, …> would satisfy any table (the guard would silently pass).
 *  Branches carrying optional companion keys (e.g. {footnoteReference,
 *  properties?}) yield those too — filtered by the caller's NonTag set. */
type TagOf<U> = U extends object ? { [K in keyof U]-?: K }[keyof U] : never;

/** Keys that appear on union branches but are NOT tags: the companion optional
 *  keys plus everything the untagged RunOptions fallback member contributes.
 *  keyof RunOptions stays a live reference, so a new RunOptions field never
 *  false-flags as a missing tag. Exceptions — office-open dual-models these as
 *  BOTH a run property and a real ParagraphChild branch; the branch wins as a
 *  tag: math (w:rPr flag vs {math: MathInput[]}), footnoteReference and
 *  endnoteReference (run-level w:*Reference vs the reference branch). */
type NonTagKeys = Exclude<
  "properties" | keyof RunOptions,
  "math" | "footnoteReference" | "endnoteReference"
>;

type ParagraphChildTag = Exclude<TagOf<ParagraphChild>, NonTagKeys>;
type SectionChildTag = TagOf<SectionChild>;

/** How a branch is handled: claimed by an editable extension (value names the
 *  claiming route), carried verbatim by a passthrough atom, or dropped (with
 *  the reason — a drop must always be a decision). */
export type Disposition =
  | { editable: string }
  | { passthrough: string }
  | { dropped: { reason: string } };

/** SectionChild branches (block level) — resolve direction. */
export const SECTION_CHILD_DISPOSITIONS = {
  paragraph: { editable: "Paragraph/Heading nodes" },
  table: { editable: "Table.parseDocxBlock" },
  toc: { editable: "tocField node (entries editable, field switches opaque)" },
  textbox: { editable: "textbox node (content editable, VML style verbatim)" },
  sdt: { editable: "sdtBlock node (content editable, control settings verbatim)" },
  altChunk: { passthrough: "block Passthrough atom" },
  customXml: { passthrough: "block Passthrough atom" },
  bookmarkStart: { passthrough: "block Passthrough atom" },
  bookmarkEnd: { passthrough: "block Passthrough atom" },
  rawXml: { passthrough: "block Passthrough atom" },
} satisfies Record<SectionChildTag, Disposition>;

/** ParagraphChild branches (inline level) — resolve direction. Unclaimed
 *  shapes all land in the inlinePassthrough atom (verbatim, byte-faithful).
 *  The untagged RunOptions fallback member is editable by construction: it
 *  resolves to text nodes + marks (see resolveRun in converters/docx.ts) —
 *  its keys are all NonTag, so it has no row here. */
export const PARAGRAPH_CHILD_DISPOSITIONS = {
  picture: { editable: "image node" },
  hyperlink: { editable: "Link mark (container)" },
  insertion: { editable: "Insertion mark (container)" },
  deletion: { editable: "Deletion mark (container)" },
  pageBreak: { editable: "pageBreak node" },
  columnBreak: { editable: "columnBreak node" },
  wpsShape: { editable: "wpsShape node (text body editable, geometry opaque)" },
  wpgGroup: { editable: "wpgGroup node (members editable, chart/contentPart members opaque)" },
  bookmarkStart: { passthrough: "inlinePassthrough atom (name exposed for TOC anchors)" },
  bookmarkEnd: { passthrough: "inlinePassthrough atom" },
  bookmark: { passthrough: "inlinePassthrough atom" },
  chart: { editable: "chart node (ChartSpaceOptions verbatim; data/type adjustments ride attrs)" },
  smartArt: { passthrough: "inlinePassthrough atom" },
  math: { editable: "MathInline node (MathInput verbatim; linear label derived)" },
  symbolRun: { passthrough: "inlinePassthrough atom" },
  footnoteReference: { passthrough: "inlinePassthrough atom" },
  endnoteReference: { passthrough: "inlinePassthrough atom" },
  commentRangeStart: { passthrough: "inlinePassthrough atom" },
  commentRangeEnd: { passthrough: "inlinePassthrough atom" },
  commentReference: { passthrough: "inlinePassthrough atom" },
  comment: { passthrough: "inlinePassthrough atom" },
  object: { passthrough: "inlinePassthrough atom" },
  pict: { passthrough: "inlinePassthrough atom" },
  contentPart: { passthrough: "inlinePassthrough atom" },
  proofErr: { passthrough: "inlinePassthrough atom" },
  positionalTab: { passthrough: "inlinePassthrough atom" },
  permStart: { editable: "permStart atom (permission range start)" },
  permEnd: { editable: "permEnd atom (permission range end)" },
  moveFromRangeStart: { editable: "moveFromRangeStart atom (move tracking source range start)" },
  moveFromRangeEnd: { editable: "moveFromRangeEnd atom (move tracking source range end)" },
  moveToRangeStart: { editable: "moveToRangeStart atom (move tracking target range start)" },
  moveToRangeEnd: { editable: "moveToRangeEnd atom (move tracking target range end)" },
  movedFrom: { editable: "moveFrom mark (container)" },
  movedTo: { editable: "moveTo mark (container)" },
  moveFrom: { editable: "moveFrom mark (container)" },
  moveTo: { editable: "moveTo mark (container)" },
  customXmlInsRangeStart: { passthrough: "inlinePassthrough atom" },
  customXmlInsRangeEnd: { passthrough: "inlinePassthrough atom" },
  customXmlDelRangeStart: { passthrough: "inlinePassthrough atom" },
  customXmlDelRangeEnd: { passthrough: "inlinePassthrough atom" },
  customXmlMoveFromRangeStart: { passthrough: "inlinePassthrough atom" },
  customXmlMoveFromRangeEnd: { passthrough: "inlinePassthrough atom" },
  customXmlMoveToRangeStart: { passthrough: "inlinePassthrough atom" },
  customXmlMoveToRangeEnd: { passthrough: "inlinePassthrough atom" },
  simpleField: { passthrough: "inlinePassthrough atom" },
  formField: { editable: "formField node (interactive legacy form field)" },
  complexField: { passthrough: "inlinePassthrough atom" },
  seqIdentifier: { passthrough: "inlinePassthrough atom" },
  pageReference: { passthrough: "inlinePassthrough atom" },
  dir: { editable: "dir mark (native bidi direction)" },
  bdo: { editable: "bdo mark (native bidi override)" },
  smartTag: { passthrough: "inlinePassthrough atom" },
  customXml: { passthrough: "inlinePassthrough atom" },
  sdt: { editable: "sdtInline node (content editable, control settings verbatim)" },
  subDoc: { passthrough: "inlinePassthrough atom" },
  rawXml: { passthrough: "inlinePassthrough atom" },
} satisfies Record<ParagraphChildTag, Disposition>;

/** RunOptions children shapes the resolve side PRESERVES as the `runMarker`
 *  atom (extensions/run-marker.ts). These are the empty `EG_RunInnerContent`
 *  elements that carry document semantics but no text — field-result
 *  placeholders, note/comment auto-marks, note separators, and Word's
 *  pagination hint. office-open emits them as `{ tag: true }` run children and
 *  its writer re-emits every one (EMPTY_RUN_ELEMENTS), so the atom round-trips
 *  byte-faithfully; before the table they were dropped at resolve and the
 *  OOXML element was lost on re-export. The `semantics` string is the tested
 *  disposition — coverage.spec drives one fixture per tag through
 *  resolve → compile and a real-XML generate → parse cycle. */
export const PRESERVED_RUN_ELEMENTS = {
  pgNum: {
    element: "w:pgNum",
    semantics: "field result placeholder for the current page number (PAGE field)",
  },
  dayShort: { element: "w:dayShort", semantics: "date field result: short day name" },
  dayLong: { element: "w:dayLong", semantics: "date field result: long day name" },
  monthShort: { element: "w:monthShort", semantics: "date field result: short month name" },
  monthLong: { element: "w:monthLong", semantics: "date field result: long month name" },
  yearShort: { element: "w:yearShort", semantics: "date field result: short year" },
  yearLong: { element: "w:yearLong", semantics: "date field result: long year" },
  annotationRef: { element: "w:annotationRef", semantics: "comment annotation auto-mark" },
  footnoteRef: { element: "w:footnoteRef", semantics: "footnote auto-number mark" },
  endnoteRef: { element: "w:endnoteRef", semantics: "endnote auto-number mark" },
  separator: { element: "w:separator", semantics: "footnote/endnote separator rule" },
  continuationSeparator: {
    element: "w:continuationSeparator",
    semantics: "footnote/endnote continuation separator rule",
  },
  lastRenderedPageBreak: {
    element: "w:lastRenderedPageBreak",
    semantics: "renderer pagination hint (Word recomputes, kept for byte fidelity)",
  },
} satisfies Record<string, { element: string; semantics: string }>;

/** RunOptions children shapes still dropped: unknown object entries outside
 *  {@link PRESERVED_RUN_ELEMENTS}. Empty by design — every known
 *  EG_RunInnerContent empty element now has a preserve route, and non-empty
 *  members are claimed by their own rules. Kept as the spec's negative probe
 *  anchor (an unregistered tag must still drop — no silent passthrough into a
 *  writer that cannot emit it). */
export const RUN_CHILDREN_DROPPED: readonly { tag: string; reason: string }[] = [];
