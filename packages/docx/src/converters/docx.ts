import { PART_REGISTRIES, type ContentTypeOverride } from "@office-open/core";
import {
  generateDocument,
  generateDocumentStream,
  generateDocumentSync,
  parseDocument,
  parseDocumentSync,
} from "@office-open/docx";
import type {
  DocumentOptions,
  SectionChild,
  SectionPropertiesOptions,
  ParagraphOptions,
  ParagraphChild,
  RunOptions,
  TableOptions,
  TableCellOptions,
  BorderOptions,
  LevelsOptions,
  OutputByType,
  OutputType,
  PackerOptions,
  StylesOptions,
  TableOfContentsOptions,
  RubyPropertiesOptions,
  GroupChildMediaData,
  SdtBlockOptions,
  SdtRunOptions,
} from "@office-open/docx";
import { flattenExtensions, getExtensionField, getSchema } from "@tiptap/core";

import type { Extensions, JSONContent } from "../core";
import { docxExtensions } from "../core";
import {
  type DrawingShapeLayout,
  stringifyVmlShapeLayout,
} from "../extensions/drawing-shape-layout";
import { memberNodeToGroupChild } from "../extensions/group-members";
import { buildListLevels, isGeneratedListReference } from "../extensions/list-numbering";
import { decodePassthroughData, encodePassthroughData } from "../extensions/passthrough";
import type { ParseBlockRule, ParseInlineRule, ResolveContext } from "../extensions/types";
import {
  isPromotableVmlPict,
  promoteVmlPictToWpsShape,
  revertWpsShapeToVmlPict,
} from "../extensions/vml-promotion";
import { foldWpsShapeName } from "../extensions/wps-shape";
import {
  assertArchiveWithinLimits,
  normalizeArchiveInput,
  normalizeArchiveInputSync,
} from "./archive-guard";
import { DOCX_EPOCH, toIsoDate, withGenerationScope } from "./determinism";
import { EncryptedDocumentError, assertNotEncryptedContainer } from "./encrypted";
import { prepareDocument, type PrepareStep } from "./prepare";
import { docenDefaultSectionProperties } from "./section-defaults";
import { buildTextBlock } from "./styles";

export type { DocumentOptions };

// ── Helpers ──

/** True when two cell-margin sets (w:tcMar / w:tblCellMar) match on every side's
 *  size — used to detect a cell that merely echoes the table's default so its
 *  redundant tcMar can be dropped for a near-identity round-trip. */
function sameCellMargins(
  a: NonNullable<TableCellOptions["margins"]>,
  b: NonNullable<TableCellOptions["margins"]>,
): boolean {
  const sz = (
    m: NonNullable<TableCellOptions["margins"]>,
    k: "top" | "right" | "bottom" | "left",
  ) => m[k]?.size ?? null;
  return (["top", "right", "bottom", "left"] as const).every((k) => sz(a, k) === sz(b, k));
}

/** True when two single-side borders match on style/size/color — used to detect
 *  a cell side that merely echoes the table's insideHorizontal/insideVertical
 *  so it can be dropped for a near-identity round-trip (resolveTable pushes
 *  those table-level grid lines onto cell sides lacking their own tcBorder). */
function sameBorder(a: BorderOptions | undefined, b: BorderOptions | undefined): boolean {
  if (!a || !b) return false;
  return a.style === b.style && a.size === b.size && a.color === b.color;
}

/** Push one item or a spread of items onto target — block/inline parse results
 *  and section-child compiles legitimately return either shape. */
function pushAll<T>(target: T[], item: T | T[]): void {
  if (Array.isArray(item)) target.push(...item);
  else target.push(item);
}

/**
 * Core document properties (docProps/core.xml) carried on `doc.attrs.core` for
 * lossless round-trip. Mirrors @office-open/core's CorePropertiesOptions, which
 * @office-open/docx's DocumentOptions extends. Inlined here (not imported from
 * @office-open/core) to keep @docen/docx's dependency surface on @office-open/docx.
 */
interface DocxCoreProperties {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  lastPrinted?: string;
  /** W3CDTF; null = the source core.xml carried none (office-open semantics). */
  created?: string | null;
  modified?: string | null;
  revision?: number;
  category?: string;
  contentStatus?: string;
  contentType?: string;
  identifier?: string;
  language?: string;
  version?: string;
}

/** Keys round-tripped between DocumentOptions core properties and
 *  `doc.attrs.core` — every data field of @office-open/core's
 *  CorePropertiesOptions except the defaultNamespace emit flag. */
const CORE_PROPERTY_KEYS: readonly (keyof DocxCoreProperties)[] = [
  "title",
  "subject",
  "creator",
  "keywords",
  "description",
  "lastModifiedBy",
  "lastPrinted",
  "created",
  "modified",
  "revision",
  "category",
  "contentStatus",
  "contentType",
  "identifier",
  "language",
  "version",
];

/**
 * DocumentOptions keys that DocxManager reconstructs (sections/numbering) or
 * carries in dedicated attrs (styles/background/core). Excluded from the
 * documentExtras pass-through so they aren't duplicated. The satisfies guard
 * pins every listed key to a real DocumentOptions field — renaming or
 * inventing one fails the build (documentExtras stays an open set: its value
 * is auto-compat with new top-level office-open keys).
 */
const COMPILE_OWNED_KEYS = new Set<string>([
  "sections",
  "numbering",
  "styles",
  "background",
  "bibliography",
  ...CORE_PROPERTY_KEYS,
] satisfies (keyof DocumentOptions)[]);

/** Collect core properties present on DocumentOptions into a plain object. */
function extractCoreProperties(docOpts: DocumentOptions): DocxCoreProperties | null {
  // Double cast: DocumentOptions (no index signature) is not comparable
  // to Record<string, unknown> in a single assertion.
  const source = docOpts as unknown as Record<string, unknown>;
  const core: Record<string, unknown> = {};
  for (const key of CORE_PROPERTY_KEYS) {
    const value = source[key];
    if (value !== undefined && value !== null) core[key] = value;
  }
  return Object.keys(core).length > 0 ? (core as DocxCoreProperties) : null;
}

/**
 * Whether a rawPart's data still has a legal DataType shape. Persistence that
 * serializes the JSON (autosave, v-model) destroys non-JSON values — a
 * Uint8Array comes back as a `{"0": …}` plain object — and such corrupted
 * bytes must not reach office-open's media reader on the next save. Binaries
 * held in memory (Uint8Array/ArrayBuffer/DataView, plus Blob/ReadableStream)
 * and JSON-native strings/number arrays pass.
 */
function isLegalDataType(value: unknown): boolean {
  if (typeof value === "string") return true;
  if (Array.isArray(value)) return value.every((n) => typeof n === "number");
  return (
    value instanceof Uint8Array ||
    value instanceof ArrayBuffer ||
    ArrayBuffer.isView(value) ||
    value instanceof Blob ||
    value instanceof ReadableStream
  );
}

/**
 * Header/footer slot shapes for the two sides of the round-trip:
 * - {@link SectionHeaderFooterGroup} — persistence side (SectionChild[] per
 *   slot); matches SectionOptions.headers/footers.
 * - {@link HeaderFooterSlots} — runtime side (JSONContent[] per slot), produced
 *   by resolveSectionChildren and consumed by compileSectionChild.
 */
type SectionHeaderFooterGroup = {
  default?: SectionChild[];
  first?: SectionChild[];
  even?: SectionChild[];
};

interface HeaderFooterSlots {
  default?: JSONContent[];
  first?: JSONContent[];
  even?: JSONContent[];
}

// ── DocxManager ──

/** One rPr mark in JSONContent's shape (name + optional attrs) — the value
 *  {@link DocxManager.runPropsToMarks} emits and `runPropsFromMarks` consumes. */
export interface RunPropMark {
  type: string;
  attrs?: Record<string, unknown>;
}

/** The format-change revision carrier mark (w:rPrChange) — not an rPr mark
 *  itself, so the format restore set excludes it. */
const FORMAT_CHANGE_MARK = "formatChange";

/** Identity memo for {@link compileDocument} — a re-compile against a JSON
 *  tree whose unchanged subtrees are referentially stable (see the editor's
 *  `pmNodeToJSON`) skips the whole subtree walk. Two entry kinds: per
 *  top-level body child (JSON object identity → compiled output + the list
 *  references it consumed) and per assembled section (input identities →
 *  the exact section object). Anything the cache cannot prove unchanged
 *  re-compiles — correct over fast. */
export interface CompileCache {
  children: WeakMap<
    object,
    { child: SectionChild | SectionChild[] | null; refs: readonly string[] }
  >;
  sections: {
    children: readonly SectionChild[];
    inputs: readonly unknown[];
    out: DocumentOptions["sections"][number];
    /** List references consumed while compiling this section's headers and
     *  footers — replayed into `usedListReferences` when the section reuses. */
    refs?: readonly string[];
  }[];
}

export function createCompileCache(): CompileCache {
  return { children: new WeakMap(), sections: [] };
}

function splitTextHyphens(
  text: string,
): Array<string | { noBreakHyphen: true } | { softHyphen: true }> {
  const pieces: Array<string | { noBreakHyphen: true } | { softHyphen: true }> = [];
  let current = "";
  for (const char of text) {
    if (char === "\u2011") {
      if (current) {
        pieces.push(current);
        current = "";
      }
      pieces.push({ noBreakHyphen: true });
    } else if (char === "\u00AD") {
      if (current) {
        pieces.push(current);
        current = "";
      }
      pieces.push({ softHyphen: true });
    } else {
      current += char;
    }
  }
  if (current) pieces.push(current);
  return pieces;
}

/**
 * Manages DOCX serialization (Tiptap JSON ↔ DocumentOptions).
 *
 * Each extension provides renderDocx/parseDocx for its own attrs mapping.
 * DocxManager handles tree walking, child assembly, and dispatching.
 */
export class DocxManager {
  // Numbering definitions registered during compile for editor-created lists
  // (a `docen-bullet`/`docen-ordered-*` reference no source definition covers).
  // Round-tripped references (`list_<numId>`) carry their definitions in the
  // source numbering and never land here.
  private numberingConfigs: { reference: string; levels: LevelsOptions[] }[] = [];
  // Generated list references seen on list paragraphs during this compile —
  // each gets a definition in numberingConfigs unless the source numbering
  // already defines it.
  private usedListReferences = new Set<string>();
  /** The open compile scope (see {@link CompileCache.children}): list refs
   *  consumed by the subtree currently compiling land here. */
  private refScope: string[] | undefined;
  // Styles table (styles.xml) carried through resolve() so consumers can
  // resolve a NUMERIC pStyle whose NAME is a heading (style "2" → name
  // "heading 1") — office-open lifts only pStyle literals that ARE
  // HeadingLevels ("Heading1".."Title"); real DOCX files often use numeric
  // ids. Set per resolve(); compile never reads it.
  private resolveStyles: StylesOptions | undefined;
  // DOCX conversion hooks collected via reflection from the extension list.
  // Each mark/node extension declares renderDocx/parseDocx as config fields;
  // the constructor reads them with getExtensionField so user-supplied
  // extensions plug in without a fork. Container marks (link/insertion/
  // deletion) wrap runs and are handled directly in compile/resolve, not
  // through these maps. Node hooks cover attrs↔opts only; block-level children
  // assembly (table rows/cells, TOC entries) is owned by each
  // block extension's parseDocxBlock rule (blockRules below); inline content
  // and the compile-side tree walk stay in DocxManager.
  private markRender = new Map<
    string,
    (attrs: Record<string, unknown>) => Record<string, unknown>
  >();
  private markParse: Array<{
    name: string;
    parse: (opts: RunOptions) => Record<string, unknown> | null;
  }> = [];
  private nodeRender = new Map<string, (node: JSONContent) => Record<string, unknown> | null>();
  private nodeParse = new Map<string, (opts: object) => Record<string, unknown>>();
  // Declarative block parse rules collected via reflection (mirrors markParse).
  // Each block node extension declares parseDocxBlock; resolveSectionChild walks
  // them in docxExtensions order before the paragraph/passthrough fallbacks.
  private blockRules: Array<{ name: string; rule: ParseBlockRule }> = [];
  // Declarative inline parse rules collected via reflection (mirrors blockRules).
  // Each inline node/mark extension declares parseDocxInline; resolveParagraphChild
  // walks them in docxExtensions order before the run/sdt/passthrough fallbacks.
  private inlineRules: Array<{ name: string; rule: ParseInlineRule }> = [];
  // Per-resolve façade over the recursive resolve entry points + read-only
  // styles, handed to every block/inline rule. Built at the start of resolve().
  private resolveCtx: ResolveContext | undefined;

  constructor(extensions: Extensions = docxExtensions) {
    const schema = getSchema(extensions);
    for (const ext of flattenExtensions(extensions)) {
      const name = ext.name;
      if (!name) continue;
      if (schema.marks[name]) {
        const render = getExtensionField(ext, "renderDocx") as
          | ((attrs: Record<string, unknown>) => Record<string, unknown>)
          | undefined;
        const parse = getExtensionField(ext, "parseDocx") as
          | ((opts: RunOptions) => Record<string, unknown> | null)
          | undefined;
        if (render) this.markRender.set(name, render);
        if (parse) this.markParse.push({ name, parse });
      } else if (schema.nodes[name]) {
        const render = getExtensionField(ext, "renderDocx") as
          | ((node: JSONContent) => Record<string, unknown> | null)
          | undefined;
        const parse = getExtensionField(ext, "parseDocx") as
          | ((opts: object) => Record<string, unknown>)
          | undefined;
        if (render) this.nodeRender.set(name, render);
        if (parse) this.nodeParse.set(name, parse);
        const blockRule = getExtensionField(ext, "parseDocxBlock") as ParseBlockRule | undefined;
        if (blockRule) this.blockRules.push({ name, rule: blockRule });
      }
      // parseDocxInline is collected for both nodes and marks — inline shapes
      // include mark containers (hyperlink/insertion/deletion) that yield text[].
      const inlineRule = getExtensionField(ext, "parseDocxInline") as ParseInlineRule | undefined;
      if (inlineRule) this.inlineRules.push({ name, rule: inlineRule });
    }
  }

  /** Reflective node renderDocx lookup: the node's DOCX opts, or {} when the
   *  node type has no renderDocx hook (degrades to a plain paragraph). node.type
   *  is optional on JSONContent — an absent type simply misses the map. */
  private renderNodeOpts(node: JSONContent): Record<string, unknown> {
    return this.nodeRender.get(node.type ?? "")?.(node) ?? {};
  }

  compile(json: JSONContent, cache?: CompileCache): DocumentOptions {
    this.numberingConfigs = [];
    this.usedListReferences = new Set();

    // Split doc content into sections. A non-final section's sectPr attaches to
    // its LAST paragraph's pPr (OOXML) — that paragraph carries sectionProperties/
    // sectionHeaders/sectionFooters attrs and closes the section here. The
    // trailing blocks form the final section, whose sectPr rides on
    // doc.attrs.sectionProperties/sectionHeaders/sectionFooters (body-level).
    // No section-carrying paragraph → single section (backward compatible).
    const sections: DocumentOptions["sections"] = [];
    let currentChildren: SectionChild[] = [];
    // Section assembly memo: the previous compile's entries stay readable
    // while this compile builds its replacements (the same array shape).
    const sectionCache = cache?.sections;
    const nextSections: typeof sectionCache = [];
    // Refs consumed by a section's headers/footers (compiled fresh only when
    // the section itself is — a reused section replays them).
    const addSection = (
      children: SectionChild[],
      props: unknown,
      headers: unknown,
      footers: unknown,
    ): void => {
      const sIdx = sections.length;
      const prev = sectionCache?.[sIdx];
      // Reuse when every input is referentially identical to the previous
      // compile's — the compiled children ride their own per-child cache, so
      // an unchanged section rebuilds no descendants at all.
      if (
        prev &&
        prev.children.length === children.length &&
        prev.children.every((c, i) => c === children[i]) &&
        prev.inputs[0] === props &&
        prev.inputs[1] === headers &&
        prev.inputs[2] === footers
      ) {
        for (const ref of prev.refs ?? []) this.usedListReferences.add(ref);
        sections.push(prev.out);
        nextSections![sIdx] = prev;
        return;
      }
      const refs: string[] = [];
      this.refScope = refs;
      const headerGroup =
        headers === null ? undefined : this.compileHeaderFooter(headers as HeaderFooterSlots);
      const footerGroup =
        footers === null ? undefined : this.compileHeaderFooter(footers as HeaderFooterSlots);
      this.refScope = undefined;
      const out = this.buildSection(
        children,
        props as SectionPropertiesOptions | null,
        headerGroup,
        footerGroup,
      );
      if (nextSections) {
        nextSections[sIdx] = { children, inputs: [props, headers, footers], out, refs };
      }
      sections.push(out);
    };
    // The child walk's scope: every list reference consumed by the compiled
    // subtree lands here (and in the node's cache entry) so a cache hit can
    // replay them into `usedListReferences`.
    if (json.content) {
      for (const node of json.content) {
        const cached = cache?.children.get(node);
        let child: SectionChild | SectionChild[] | null;
        if (cached) {
          child = cached.child;
          for (const ref of cached.refs) this.usedListReferences.add(ref);
        } else {
          const refs: string[] = [];
          this.refScope = refs;
          child = this.compileSectionChild(node);
          this.refScope = undefined;
          cache?.children.set(node, { child, refs });
        }
        if (child) pushAll(currentChildren, child);
        if (node.type === "paragraph") {
          const na = (node.attrs ?? {}) as Record<string, unknown>;
          if (na.sectionProperties != null) {
            addSection(
              currentChildren,
              na.sectionProperties,
              na.sectionHeaders ?? null,
              na.sectionFooters ?? null,
            );
            currentChildren = [];
          }
        }
      }
    }
    const docAttrs = json.attrs ?? {};
    addSection(
      currentChildren,
      docAttrs.sectionProperties ?? null,
      docAttrs.sectionHeaders ?? null,
      docAttrs.sectionFooters ?? null,
    );

    const styles = (docAttrs.styles ?? undefined) as DocumentOptions["styles"] | undefined;
    const core = (docAttrs.core ?? undefined) as DocxCoreProperties | undefined;
    const background = (docAttrs.background ?? undefined) as
      | DocumentOptions["background"]
      | undefined;
    // Install this compile's section memo — see addSection.
    if (cache) cache.sections = nextSections!;
    const documentExtras = (docAttrs.documentExtras ?? undefined) as
      | Partial<DocumentOptions>
      | undefined;
    const bibliography = (docAttrs.bibliography ?? undefined) as
      | DocumentOptions["bibliography"]
      | undefined;
    // Source numbering carries verbatim from resolve: the pic-bullet
    // definitions, the cleanup id, and the abstract definitions themselves
    // (instances, overrides, aliases). Only abstractNumberings is rebuilt
    // below — the merge of source definitions with regenerated ordered-list
    // definitions; drop originals shadowed by a regenerated reference to
    // avoid duplicates.
    const sourceNumbering = (docAttrs.numbering ?? undefined) as DocumentOptions["numbering"];
    const origNumberingConfig = sourceNumbering?.abstractNumberings ?? [];
    // Register a generated reference's definition unless the source numbering
    // already carries one (an editor-created list re-toggling a round-tripped
    // reference must not shadow the original marker definition).
    for (const reference of this.usedListReferences) {
      if (!isGeneratedListReference(reference)) continue;
      if (origNumberingConfig.some((c) => c.reference === reference)) continue;
      const levels = buildListLevels(reference);
      if (levels) this.numberingConfigs.push({ reference, levels });
    }
    // Direct concatenation: the loop above skips every reference the source
    // numbering already carries, so the two lists never intersect.
    const numberingConfig = [...origNumberingConfig, ...this.numberingConfigs];
    // rawParts entries whose bytes a JSON round-trip already corrupted (see
    // isLegalDataType) drop on a copy — the theme part regenerates from
    // scratch without its rawPart — and the input attrs stay untouched
    // (compile never mutates its input).
    let extras = documentExtras;
    const sourceParts = extras?.rawParts;
    if (
      extras &&
      Array.isArray(sourceParts) &&
      sourceParts.some((part) => !isLegalDataType(part?.data))
    ) {
      const legal = sourceParts.filter((part) => isLegalDataType(part?.data));
      const { rawParts: _dropped, ...rest } = extras;
      extras = legal.length > 0 ? { ...rest, rawParts: legal } : rest;
    }
    let numbering: DocumentOptions["numbering"];
    if (numberingConfig.length > 0) {
      numbering = { ...sourceNumbering, abstractNumberings: numberingConfig };
    } else if (sourceNumbering) {
      numbering = sourceNumbering;
    }
    return {
      sections,
      ...(styles ? { styles } : {}),
      ...core,
      ...(background ? { background } : {}),
      ...(bibliography ? { bibliography } : {}),
      ...extras,
      ...(numbering ? { numbering } : {}),
    };
  }

  /** Assemble a SectionOptions from compiled children + optional layout/headers/footers.
   *  A section the model gives no properties is stamped with the docen
   *  defaults — every generated document carries its own explicit page
   *  geometry instead of inheriting office-open's zh-CN
   *  `sectionMarginDefaults` at stringify time. */
  private buildSection(
    children: SectionChild[],
    properties: SectionPropertiesOptions | null,
    headers: SectionHeaderFooterGroup | undefined,
    footers: SectionHeaderFooterGroup | undefined,
  ): DocumentOptions["sections"][number] {
    return {
      children,
      properties: properties ?? docenDefaultSectionProperties(),
      ...(headers ? { headers } : {}),
      ...(footers ? { footers } : {}),
    };
  }

  /**
   * Compile resolved header/footer slots (JSONContent[] per slot) back into
   * SectionChild[] per slot. Returns undefined when no slot has content.
   */
  private compileHeaderFooter(
    slots: HeaderFooterSlots | null,
  ): SectionHeaderFooterGroup | undefined {
    if (!slots) return undefined;
    const group: SectionHeaderFooterGroup = {};
    for (const slot of ["default", "first", "even"] as const) {
      const json = slots[slot];
      if (!json?.length) continue;
      const children: SectionChild[] = [];
      for (const node of json) {
        const child = this.compileSectionChild(node);
        if (!child) continue;
        pushAll(children, child);
      }
      if (children.length > 0) group[slot] = children;
    }
    return Object.keys(group).length > 0 ? group : undefined;
  }

  /**
   * Resolve a section's header/footer group (SectionChild[] per slot) into
   * Tiptap JSON slots. Returns null when no slot has content.
   */
  private resolveHeaderFooter(
    group: SectionHeaderFooterGroup | undefined,
  ): HeaderFooterSlots | null {
    if (!group) return null;
    const slots: HeaderFooterSlots = {};
    for (const slot of ["default", "first", "even"] as const) {
      const children = group[slot];
      if (children?.length) {
        slots[slot] = this.resolveSectionChildren(children);
      }
    }
    return Object.keys(slots).length > 0 ? slots : null;
  }

  resolve(docOpts: DocumentOptions): JSONContent {
    // The manager is a long-lived singleton: the per-document façade fields
    // must not outlive this resolve, or the last document's styles snapshot
    // and resolve closures stay pinned until the next one replaces them.
    try {
      return this.resolveInner(docOpts);
    } finally {
      this.resolveCtx = undefined;
      this.resolveStyles = undefined;
    }
  }

  private resolveInner(docOpts: DocumentOptions): JSONContent {
    this.resolveStyles = docOpts.styles ?? undefined;
    const sections = docOpts.sections ?? [];
    if (sections.length === 0) {
      return { type: "doc", content: [{ type: "paragraph" }] };
    }
    // Per-resolve façade handed to block parse rules. Arrow closures capture
    // `this`; `styles` is a read-only snapshot of the instance field set above
    // (stable for this resolve's lifetime).
    const styles = this.resolveStyles;
    this.resolveCtx = {
      resolveBlockStream: (children) => this.resolveSectionChildren(children),
      resolveBlock: (child) => this.resolveSectionChild(child),
      resolveInlineContent: (para) => this.resolveInlineContent(para),
      resolveInlineChildren: (children) => this.resolveParagraphChildren(children),
      resolveParagraph: (para) => this.resolveParagraph(para),
      parseNodeAttrs: (type, opts) => this.nodeParse.get(type)?.(opts) ?? {},
      resolveMarks: (opts) => this.resolveMarks(opts),
      styles,
    };

    // Resolve every section's children into blocks. A non-final section's
    // sectPr attaches to that section's LAST paragraph's pPr (OOXML) — stamp it
    // on that paragraph's attrs, not a standalone sectionBreak node. The final
    // section's sectPr rides on doc.attrs (it lives at <w:body>'s end).
    const content: JSONContent[] = [];
    const lastIndex = sections.length - 1;
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i];
      const sectionContent = this.resolveSectionChildren(section.children ?? []);
      if (i < lastIndex) {
        const sectAttrs: Record<string, unknown> = {
          sectionProperties: section.properties ?? null,
          sectionHeaders: this.resolveHeaderFooter(section.headers),
          sectionFooters: this.resolveHeaderFooter(section.footers),
        };
        const last = sectionContent[sectionContent.length - 1];
        // Stamp sectPr on the section's last paragraph directly (a heading is
        // a paragraph too — same node) instead of appending a stray empty one.
        if (last?.type === "paragraph") {
          last.attrs = { ...last.attrs, ...sectAttrs };
        } else {
          // Section ends on a non-textblock (table/etc.) — Word still needs the
          // sectPr on a paragraph, so append an empty one to carry it.
          sectionContent.push({ type: "paragraph", attrs: sectAttrs });
        }
      }
      content.push(...sectionContent);
    }

    const doc: JSONContent = {
      type: "doc",
      content: content.length > 0 ? content : [{ type: "paragraph" }],
    };
    // Carry the styles library (styles.xml) and core properties (docProps/
    // core.xml) through the JSON for lossless round-trip. office-open regenerates
    // importedStyles/docDefaultsXml/latentStylesXml from `styles`, and writes
    // `core` back to docProps/core.xml.
    const attrs: Record<string, unknown> = {};
    if (docOpts.styles) attrs.styles = docOpts.styles;
    if (docOpts.background) attrs.background = docOpts.background;
    // Source numbering definitions (abstractNum) carried verbatim so list
    // markers round-trip; compile merges these with regenerated ordered defs.
    if (docOpts.numbering) attrs.numbering = docOpts.numbering;
    // Bibliography sources (word/bibliography.xml) — the citation dialog's
    // master list and the bibliography block's data.
    if (docOpts.bibliography) attrs.bibliography = docOpts.bibliography;
    const core = extractCoreProperties(docOpts);
    if (core) attrs.core = core;
    const lastSection = sections[lastIndex];
    if (lastSection.properties) attrs.sectionProperties = lastSection.properties;
    const lastHeaders = this.resolveHeaderFooter(lastSection.headers);
    if (lastHeaders) attrs.sectionHeaders = lastHeaders;
    const lastFooters = this.resolveHeaderFooter(lastSection.footers);
    if (lastFooters) attrs.sectionFooters = lastFooters;
    // Pass through document-level fields DocxManager doesn't reconstruct
    // (settings.xml flags like displayBackgroundShape, zoom, fonts, footnotes,
    // customProperties, …). Word needs displayBackgroundShape to render the
    // <w:background> element, so losing it makes the page background invisible.
    const documentExtras: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(docOpts as unknown as Record<string, unknown>)) {
      if (COMPILE_OWNED_KEYS.has(k)) continue;
      if (v === undefined || v === null) continue;
      documentExtras[k] = v;
    }
    if (Object.keys(documentExtras).length > 0) attrs.documentExtras = documentExtras;
    if (Object.keys(attrs).length > 0) doc.attrs = attrs;
    return doc;
  }

  // ── Compile: Tiptap JSON → DocumentOptions ──

  private compileSectionChild(node: JSONContent): SectionChild | SectionChild[] | null {
    switch (node.type) {
      case "paragraph":
        return { paragraph: this.compileParagraphNode(node) };
      case "table":
        return { table: this.compileTableNode(node) };
      case "image": {
        const imageRun = this.nodeRender.get(node.type)?.(node) ?? null;
        if (!imageRun) return null;
        return { paragraph: { children: [imageRun] } };
      }
      case "tocField": {
        const options = (node.attrs?.options as TableOfContentsOptions | undefined) ?? {};
        const entries: SectionChild[] = [];
        for (const child of node.content ?? []) {
          const compiled = this.compileSectionChild(child);
          if (!compiled) continue;
          pushAll(entries, compiled);
        }
        return { toc: { ...options, entries } };
      }
      case "sdtBlock": {
        // Content-control container (reverse of the sdt block rule): children
        // walk the shared SectionChild dispatch; the control settings ride
        // attrs verbatim.
        const sdtChildren: SectionChild[] = [];
        for (const child of node.content ?? []) {
          const compiled = this.compileSectionChild(child);
          if (!compiled) continue;
          pushAll(sdtChildren, compiled);
        }
        const attrs = (node.attrs ?? {}) as {
          properties?: SdtBlockOptions["properties"];
          endProperties?: SdtBlockOptions["endProperties"];
        };
        return {
          sdt: {
            properties: attrs.properties ?? ({} as SdtBlockOptions["properties"]),
            // SdtBlockOptions.children is typed narrower (BlockContentChild[])
            // than what CT_SdtContent actually allows; the stringify runtime
            // dispatches the full SectionChild set (ctx.stringifyChild).
            children: sdtChildren as NonNullable<SdtBlockOptions["children"]>,
            ...(attrs.endProperties ? { endProperties: attrs.endProperties } : {}),
          },
        };
      }
      case "textbox": {
        // VML text box (reverse of the textbox block rule): the box data rides
        // attrs verbatim; children walk the shared SectionChild dispatch. The
        // branch type is an inline intersection in SectionChild — extract it.
        type TextboxBranch = Extract<SectionChild, { textbox: unknown }>;
        const box = {
          ...((node.attrs?.textbox ?? {}) as Omit<TextboxBranch["textbox"], "children">),
        };
        const layout = node.attrs?.layout as DrawingShapeLayout | undefined;
        if (layout && !box.style) {
          (box as any).style = stringifyVmlShapeLayout(layout);
        }
        const boxChildren: SectionChild[] = [];
        for (const child of node.content ?? []) {
          const compiled = this.compileSectionChild(child);
          if (!compiled) continue;
          pushAll(boxChildren, compiled);
        }
        return { textbox: { ...box, children: boxChildren } };
      }
      case "passthrough": {
        // Opaque SectionChild (rawXml/bookmark/toc/altChunk/…) round-tripped
        // verbatim; binary leaves (OLE embed bytes, …) are revived from the
        // codec's base64 marker.
        const data = (node.attrs?.data as string) ?? "{}";
        try {
          return decodePassthroughData<SectionChild>(data);
        } catch {
          return null;
        }
      }
      default:
        return null;
    }
  }

  private compileParagraphNode(node: JSONContent): ParagraphOptions | string {
    // A list paragraph references its numbering definition by attr — collect
    // the reference so compile registers a generated definition for it.
    const numRef = (node.attrs?.numbering as { reference?: string } | null | undefined)?.reference;
    if (typeof numRef === "string" && numRef) {
      this.usedListReferences.add(numRef);
      this.refScope?.push(numRef);
    }
    const opts = this.renderNodeOpts(node);
    const childList = this.compileInlineContent(node.content);
    if (childList.length > 0) opts.children = childList;
    return this.simplifyParagraph(opts);
  }

  /** Simple text optimization: merge plain runs into the text field (a
   *  children-less paragraph with no other properties collapses to a plain
   *  string — SectionChild.paragraph accepts both). */
  private simplifyParagraph(opts: Record<string, unknown>): ParagraphOptions | string {
    const children = opts.children as ParagraphChild[] | undefined;
    if (!children || children.length === 0) return opts as ParagraphOptions;

    const allSimpleText = children.every(
      (c) =>
        typeof c === "object" && c !== null && "text" in c && Object.keys(c as object).length === 1,
    );
    if (allSimpleText) {
      const combined = children.map((c) => (c as { text: string }).text).join("");
      delete opts.children;
      if (combined && Object.keys(opts).length === 0) return combined;
      if (combined) opts.text = combined;
    }

    return opts as ParagraphOptions;
  }

  private compileTableNode(node: JSONContent): TableOptions {
    const opts = this.renderNodeOpts(node);
    const colCount = this.getTableColumnCount(node);
    // Table-level tblCellMar default — passed to compileTableCellNode so it can
    // drop a cell tcMar that merely echoes it (see resolveTable's push-down).
    const tableCellMargins = (opts.cellMargin ?? opts.margins ?? null) as NonNullable<
      TableCellOptions["margins"]
    > | null;
    // Table-level insideH/V — passed to compileTableCellNode so it can drop a
    // cell side that merely echoes them (resolveTable pushed them onto cells).
    const tableBorders = (opts.borders ?? null) as {
      insideHorizontal?: BorderOptions;
      insideVertical?: BorderOptions;
    } | null;
    const insideH = tableBorders?.insideHorizontal ?? null;
    const insideV = tableBorders?.insideVertical ?? null;
    // Compute tblGrid column widths up front so each cell's tcW can be derived
    // from the grid columns it spans. A cell without its own width attr would
    // otherwise collapse to the library default and autofit tables with short
    // content render as a sliver.
    const { columnWidths, tableWidth } =
      colCount > 0
        ? this.computeColumnWidths(node, colCount)
        : { columnWidths: [] as number[], tableWidth: { size: 0, type: "twips" } };
    const rows: Record<string, unknown>[] = [];

    // Cell attrs already mirror TableCellOptions (columnSpan/verticalMerge/
    // width included), so rows/cells pass straight through — no vMerge
    // interleave or rowspan rebuild. A cell's tcW is still derived from the
    // tblGrid columns it spans when it carries no width of its own.
    for (const rowNode of node.content ?? []) {
      if (rowNode.type !== "tableRow") continue;

      const rowOpts = this.nodeRender.get(rowNode.type)?.(rowNode) ?? {};
      const compiledCells: Record<string, unknown>[] = [];
      let colIdx = 0;
      for (const cellNode of rowNode.content ?? []) {
        if (cellNode.type !== "tableCell") continue;
        compiledCells.push(
          this.compileTableCellNode(
            cellNode,
            tableCellMargins,
            insideH,
            insideV,
            columnWidths,
            colIdx,
          ),
        );
        colIdx += (cellNode.attrs?.columnSpan as number) ?? 1;
      }
      rowOpts.cells = compiledCells;
      rows.push(rowOpts);
    }

    opts.rows = rows;

    if (colCount > 0) {
      opts.columnWidths = columnWidths;
      if (!opts.width) opts.width = tableWidth;
      if (!opts.layout) opts.layout = "autofit";
    }

    // Double cast: TableOptions requires `rows`, which opts gains only at
    // runtime from the loop above.
    return opts as unknown as TableOptions;
  }

  /** Count grid columns from the first table row (summing columnSpan). */
  private getTableColumnCount(tableNode: JSONContent): number {
    for (const rowNode of tableNode.content ?? []) {
      if (rowNode.type !== "tableRow") continue;
      let count = 0;
      for (const cell of rowNode.content ?? []) {
        if (cell.type === "tableCell") {
          count += (cell.attrs?.columnSpan as number) ?? 1;
        }
      }
      return count;
    }
    return 0;
  }

  /** Sum the tblGrid widths covered by a cell starting at `start` and spanning
   *  `span` grid columns — the OOXML tcW for that cell. */
  private sumGridSpan(columnWidths: number[], start: number, span: number): number {
    let sum = 0;
    for (let i = start; i < start + span && i < columnWidths.length; i++) {
      sum += columnWidths[i] ?? 0;
    }
    return sum;
  }

  /**
   * Compute tblGrid column widths (twips). The table's columnWidths attr
   * (tblGrid from a DOCX round-trip or insert-table) carries exact twips;
   * without it columns split the A4 content width evenly and the table
   * takes 100% width.
   */
  private computeColumnWidths(
    tableNode: JSONContent,
    colCount: number,
  ): { columnWidths: number[]; tableWidth: { size: number; type: string } } {
    const DEFAULT_CONTENT_TWIPS = 9026; // A4 with 1-inch margins (~15.9cm)

    const tableColWidths = tableNode.attrs?.columnWidths as number[] | null | undefined;
    if (tableColWidths && tableColWidths.length > 0) {
      const filled = tableColWidths.slice(0, colCount);
      while (filled.length < colCount) {
        filled.push(filled[filled.length - 1] ?? Math.floor(DEFAULT_CONTENT_TWIPS / colCount));
      }
      const total = filled.reduce((a, b) => a + b, 0);
      return { columnWidths: filled, tableWidth: { size: total, type: "twips" } };
    }

    // No explicit widths → equal distribution, percentage table width
    const equal = Math.floor(DEFAULT_CONTENT_TWIPS / colCount);
    return {
      columnWidths: Array(colCount).fill(equal),
      tableWidth: { size: 5000, type: "pct" }, // 100%
    };
  }

  private compileTableCellNode(
    cellNode: JSONContent,
    tableMargins?: NonNullable<TableCellOptions["margins"]> | null,
    insideH?: BorderOptions | null,
    insideV?: BorderOptions | null,
    columnWidths?: number[] | null,
    colIdx?: number,
  ): Record<string, unknown> {
    const cellOpts = this.renderNodeOpts(cellNode);

    // Restore the table-level form: resolveTable pushed the table's tblCellMar
    // default onto cells without their own tcMar. A cell whose margins equal
    // that default is the inherited one — drop its cell-level tcMar so the
    // regenerated docx keeps the compact table-level tblCellMar (near-identity
    // round-trip) instead of duplicating tcMar on every cell.
    if (
      tableMargins &&
      cellOpts.margins &&
      sameCellMargins(cellOpts.margins as NonNullable<TableCellOptions["margins"]>, tableMargins)
    ) {
      delete cellOpts.margins;
    }

    // Restore the table-level form: resolveTable pushed the table's insideH/V
    // onto cell sides lacking their own tcBorder. A side equal to the table's
    // insideH/V is the inherited one — drop it so the regenerated docx keeps
    // tblBorders.insideH/V instead of duplicating as tcBorders on every cell.
    // The side drops land on a COPY: the opts' borders object is shared with
    // the live editor attrs, and deleting its sides in place would strip the
    // document model on every render (an explicit stroke equal to the table
    // grid would vanish from the cells that carry it).
    if ((insideH || insideV) && cellOpts.borders) {
      const b = { ...(cellOpts.borders as Record<string, BorderOptions | undefined>) };
      if (insideH && sameBorder(b.top, insideH)) delete b.top;
      if (insideH && sameBorder(b.bottom, insideH)) delete b.bottom;
      if (insideV && sameBorder(b.left, insideV)) delete b.left;
      if (insideV && sameBorder(b.right, insideV)) delete b.right;
      if (Object.keys(b).length === 0) delete cellOpts.borders;
      else cellOpts.borders = b;
    }

    // A cell may contain ANY block (nested table/list/…), not just
    // paragraphs. Route each child through the shared block compiler so a
    // nested list or table survives the round-trip.
    const cellChildren: SectionChild[] = [];
    for (const childNode of cellNode.content ?? []) {
      const compiled = this.compileSectionChild(childNode);
      if (compiled == null) continue;
      pushAll(cellChildren, compiled);
    }
    if (cellChildren.length > 0) cellOpts.children = cellChildren;

    // tcW = sum of the tblGrid columns the cell spans. tblGrid carries the
    // authoritative column widths — recomputing from it keeps autofit tables
    // with short content from collapsing to a sliver when the parsed tcW
    // disagrees.
    if (columnWidths && columnWidths.length > 0 && colIdx != null) {
      const span = (cellNode.attrs?.columnSpan as number) ?? 1;
      const tw = this.sumGridSpan(columnWidths, colIdx, span);
      if (tw > 0) cellOpts.width = { size: tw, type: "twips" };
    }

    return cellOpts;
  }

  // ── Inline content ──

  private compileInlineContent(content?: JSONContent[]): ParagraphChild[] {
    if (!content) return [];
    const children: ParagraphChild[] = [];

    for (const node of content) {
      switch (node.type) {
        case "text":
          this.compileTextNode(node, children);
          break;
        case "hardBreak":
          if (node.attrs?.variant === "carriageReturn") {
            children.push({ children: [{ carriageReturn: true }] } as unknown as ParagraphChild);
          } else {
            children.push({ break: 1 });
          }
          break;
        case "permStart": {
          const ps = (node.attrs ?? {}) as Record<string, unknown>;
          children.push({
            permStart: {
              id: ps.id ?? 0,
              ...(ps.editGroup ? { editGroup: ps.editGroup as any } : {}),
              ...(ps.editor ? { editor: ps.editor as string } : {}),
              ...(ps.colFirst != null ? { colFirst: Number(ps.colFirst) } : {}),
              ...(ps.colLast != null ? { colLast: Number(ps.colLast) } : {}),
            },
          } as unknown as ParagraphChild);
          break;
        }
        case "permEnd": {
          const pe = (node.attrs ?? {}) as Record<string, unknown>;
          children.push({
            permEnd: pe.id ?? 0,
          } as unknown as ParagraphChild);
          break;
        }
        case "moveFromRangeStart": {
          const m = (node.attrs ?? {}) as {
            id?: number | string;
            name?: string;
            author?: string;
            date?: string;
            displacedByCustomXml?: unknown;
            colFirst?: number | string;
            colLast?: number | string;
          };
          children.push({
            moveFromRangeStart: {
              id: m.id != null ? Number(m.id) : 0,
              ...(m.name != null ? { name: String(m.name) } : {}),
              ...(m.author != null ? { author: String(m.author) } : {}),
              ...(m.date != null ? { date: String(m.date) } : {}),
              ...(m.displacedByCustomXml != null
                ? { displacedByCustomXml: m.displacedByCustomXml as any }
                : {}),
              ...(m.colFirst != null ? { colFirst: Number(m.colFirst) } : {}),
              ...(m.colLast != null ? { colLast: Number(m.colLast) } : {}),
            },
          } as unknown as ParagraphChild);
          break;
        }
        case "moveFromRangeEnd": {
          const m = (node.attrs ?? {}) as { id?: number | string; displacedByCustomXml?: unknown };
          children.push({
            moveFromRangeEnd: {
              id: m.id != null ? Number(m.id) : 0,
              ...(m.displacedByCustomXml != null
                ? { displacedByCustomXml: m.displacedByCustomXml as any }
                : {}),
            },
          } as unknown as ParagraphChild);
          break;
        }
        case "moveToRangeStart": {
          const m = (node.attrs ?? {}) as {
            id?: number | string;
            name?: string;
            author?: string;
            date?: string;
            displacedByCustomXml?: unknown;
            colFirst?: number | string;
            colLast?: number | string;
          };
          children.push({
            moveToRangeStart: {
              id: m.id != null ? Number(m.id) : 0,
              ...(m.name != null ? { name: String(m.name) } : {}),
              ...(m.author != null ? { author: String(m.author) } : {}),
              ...(m.date != null ? { date: String(m.date) } : {}),
              ...(m.displacedByCustomXml != null
                ? { displacedByCustomXml: m.displacedByCustomXml as any }
                : {}),
              ...(m.colFirst != null ? { colFirst: Number(m.colFirst) } : {}),
              ...(m.colLast != null ? { colLast: Number(m.colLast) } : {}),
            },
          } as unknown as ParagraphChild);
          break;
        }
        case "moveToRangeEnd": {
          const m = (node.attrs ?? {}) as Record<string, unknown>;
          children.push({
            moveToRangeEnd: {
              id: m.id != null ? Number(m.id) : 0,
              ...(m.displacedByCustomXml != null
                ? { displacedByCustomXml: m.displacedByCustomXml as any }
                : {}),
            },
          } as unknown as ParagraphChild);
          break;
        }
        case "formField": {
          const ff = (node.attrs?.formField ?? {}) as Record<string, unknown>;
          let text = "";
          if (node.content) {
            for (const c of node.content) {
              if (c.type === "text" && c.text) text += c.text;
            }
          }
          const updatedFf: Record<string, unknown> = { ...ff };
          if (updatedFf.checkBox && typeof updatedFf.checkBox === "object") {
            const cb = updatedFf.checkBox as Record<string, unknown>;
            updatedFf.checkBox = {
              ...cb,
              checked: text.includes("☒") || (text !== "☐" && Boolean(cb.checked)),
            };
          } else if (updatedFf.dropDownList && typeof updatedFf.dropDownList === "object") {
            const ddl = updatedFf.dropDownList as {
              entries?: string[];
              result?: number;
              default?: number;
            };
            const entries = Array.isArray(ddl.entries) ? ddl.entries : [];
            const idx = entries.indexOf(text);
            updatedFf.dropDownList = {
              ...ddl,
              result: idx >= 0 ? idx : (ddl.result ?? 0),
            };
          } else if (updatedFf.textInput && typeof updatedFf.textInput === "object") {
            const ti = updatedFf.textInput as Record<string, unknown>;
            updatedFf.textInput = {
              ...ti,
              value: text,
            };
          }
          children.push({ formField: updatedFf } as unknown as ParagraphChild);
          break;
        }
        case "pageBreak":
          children.push({ pageBreak: true });
          break;
        case "columnBreak":
          children.push({ columnBreak: true });
          break;
        case "tab":
          // office-open emits <w:tab/> from a top-level tab:true run, but its
          // ParagraphChild union doesn't list the top-level shape — double
          // assertion bridges the .d.ts gap.
          children.push({ tab: true } as Record<string, unknown> as ParagraphChild);
          break;
        case "inlinePassthrough": {
          // Opaque inline ParagraphChild (bookmark/range markers, …) carried
          // verbatim — reverse of resolveParagraphChild's fallback. Binary
          // leaves (an OLE object's embed bytes, …) revive from the codec.
          const data = (node.attrs?.data as string) ?? "{}";
          try {
            const parsed = decodePassthroughData<ParagraphChild>(data);
            if (parsed) children.push(parsed);
          } catch {
            /* malformed JSON — drop */
          }
          break;
        }
        case "mathInline": {
          // Reverse of MathInline's parseDocxInline rule: the MathInput rides
          // the `math` attr verbatim.
          const math = node.attrs?.math;
          if (math && typeof math === "object") {
            children.push({ math } as Record<string, unknown> as ParagraphChild);
          }
          break;
        }
        case "sdtInline": {
          // Content-control container (reverse of the sdt inline rule); the
          // control settings ride attrs verbatim.
          const attrs = (node.attrs ?? {}) as {
            properties?: SdtRunOptions["properties"];
            endProperties?: SdtRunOptions["endProperties"];
          };
          const inline = this.compileInlineContent(node.content);
          children.push({
            sdt: {
              properties: attrs.properties ?? ({} as SdtRunOptions["properties"]),
              ...(inline.length > 0 ? { children: inline } : {}),
              ...(attrs.endProperties ? { endProperties: attrs.endProperties } : {}),
            },
          });
          break;
        }
        case "image": {
          const imageRun = this.nodeRender.get(node.type)?.(node) ?? null;
          if (imageRun) children.push(imageRun);
          break;
        }
        case "chart": {
          // Verbatim ChartOptions (ChartSpaceOptions + anchor fields) — reverse
          // of the chart inline rule. Same reflection as image.
          const chartRun = this.nodeRender.get(node.type)?.(node) ?? null;
          if (chartRun) children.push(chartRun);
          break;
        }
        case "model3d":
        case "ink": {
          const drawingRun = this.nodeRender.get(node.type)?.(node) ?? null;
          if (drawingRun) children.push(drawingRun as ParagraphChild);
          break;
        }
        case "wpgGroup": {
          // attrs.wpgGroup is GroupOptions minus children; the member sequence
          // compiles back through group-members (reverse of resolveGroupOptions).
          // Same .d.ts gap as tab:true above — office-open consumes a wpgGroup
          // run but lists the branch only in TrackChangeChild, not here.
          const wpgGroup = node.attrs?.wpgGroup;
          if (wpgGroup) {
            const members: GroupChildMediaData[] = [];
            for (const member of node.content ?? []) {
              const child = memberNodeToGroupChild(member, (n) => this.compileShapeBody(n));
              if (child) members.push(child);
            }
            children.push({
              wpgGroup: { ...wpgGroup, children: members },
            } as unknown as ParagraphChild);
          }
          break;
        }
        case "wpsShape": {
          // Editable text body: compileShapeBody (shared with the wpg group
          // member compile) rebuilds wpsShape.children. The editor-facing
          // `name` folds into the OOXML slot (wps:cNvSpPr/@name).
          // Same .d.ts gap as tab:true above for the ParagraphChild union.
          const geometry = (node.attrs?.wpsShape ?? {}) as Record<string, unknown>;
          if (geometry.vmlOrigin) {
            children.push(revertWpsShapeToVmlPict(geometry));
          } else {
            children.push({
              wpsShape: { ...foldWpsShapeName(geometry), children: this.compileShapeBody(node) },
            } as unknown as ParagraphChild);
          }
          break;
        }
      }
    }

    return children;
  }

  /** Compile a wpsShape node's editable body (PM content) back to the
   *  ParagraphOptions list ShapeCoreOptions.children carries — each content
   *  paragraph through compileSectionChild, unwrapping .paragraph. Shared by
   *  the standalone wpsShape case and the wpg group member compile. */
  private compileShapeBody(node: JSONContent): (ParagraphOptions | string)[] {
    const body: (ParagraphOptions | string)[] = [];
    for (const child of node.content ?? []) {
      const compiled = this.compileSectionChild(child);
      if (!compiled) continue;
      const items = Array.isArray(compiled) ? compiled : [compiled];
      for (const it of items) {
        if (it && typeof it === "object" && "paragraph" in (it as object)) {
          body.push((it as { paragraph: ParagraphOptions | string }).paragraph);
        }
      }
    }
    return body;
  }

  private compileTextNode(node: JSONContent, children: ParagraphChild[]): void {
    const text = node.text ?? "";
    if (!text) return;

    // Split on "\n": OOXML ignores a literal "\n" inside <w:t>, so each newline
    // becomes a {break:1} run. Shared by paragraphs and code blocks — codeBlock
    // needs no special-case newline handling.
    const segments = text.split("\n");
    for (let i = 0; i < segments.length; i++) {
      if (i > 0) children.push({ break: 1 });
      if (segments[i]) this.compileTextRun(segments[i], node.marks, children);
    }
  }

  /** Emit a single run for `text` with all inline marks applied. Container
   *  marks (hyperlink/insertion/deletion) wrap the run AFTER the rPr overlay
   *  marks are collected — a link whose run carries a character style (e.g.
   *  the stamped "Hyperlink") must keep it inside the container, so the
   *  container legs can no longer short-circuit the overlay pass. */
  private compileTextRun(
    text: string,
    marks: JSONContent["marks"],
    children: ParagraphChild[],
  ): void {
    const hasHyphenBreak = text.includes("\u2011") || text.includes("\u00AD");
    const runOpts: Record<string, unknown> = hasHyphenBreak
      ? { children: splitTextHyphens(text) }
      : { text };
    let linkMark: NonNullable<JSONContent["marks"]>[number] | undefined;
    let trackMark: NonNullable<JSONContent["marks"]>[number] | undefined;
    let rubyMark: NonNullable<JSONContent["marks"]>[number] | undefined;
    let dirMark: NonNullable<JSONContent["marks"]>[number] | undefined;
    let bdoMark: NonNullable<JSONContent["marks"]>[number] | undefined;

    for (const mark of marks ?? []) {
      if (mark.type === "link") {
        if (mark.attrs?.href) linkMark = mark;
        continue;
      }
      if (
        mark.type === "insertion" ||
        mark.type === "deletion" ||
        mark.type === "moveFrom" ||
        mark.type === "moveTo"
      ) {
        trackMark = mark;
        continue;
      }
      if (mark.type === "ruby") {
        rubyMark = mark;
        continue;
      }
      if (mark.type === "dir") {
        dirMark = mark;
        continue;
      }
      if (mark.type === "bdo") {
        bdoMark = mark;
        continue;
      }
      if (mark.type === "softHyphen") {
        continue;
      }
      // rPr overlay marks — each extension's renderDocx contributes run props.
      const render = this.markRender.get(mark.type);
      if (render) Object.assign(runOpts, render((mark.attrs ?? {}) as Record<string, unknown>));
    }

    if (rubyMark) {
      children.push(this.compileRubyRun(rubyMark.attrs, text, runOpts));
      return;
    }

    if (linkMark) {
      const href = linkMark.attrs?.href as string;
      const { text: _, children: __, ...runWithoutText } = runOpts;
      const linkChildren: (RunOptions | string)[] = [];
      if (hasHyphenBreak) {
        linkChildren.push({ ...runWithoutText, children: splitTextHyphens(text) } as RunOptions);
      } else if (text) {
        linkChildren.push({ ...runWithoutText, text } as RunOptions);
      }
      let finalLink: ParagraphChild = {
        hyperlink: {
          url: href.startsWith("#") ? undefined : href,
          anchor: href.startsWith("#") ? href.slice(1) : undefined,
          // resolveHyperlink parks w:tooltip on the mark's title attr.
          tooltip: (linkMark.attrs?.title as string | null | undefined) ?? undefined,
          children: linkChildren,
        },
      };
      if (dirMark) {
        const val = (dirMark.attrs?.val === "rtl" ? "rtl" : "ltr") as "ltr" | "rtl";
        finalLink = { dir: { val, children: [finalLink] } } as ParagraphChild;
      }
      if (bdoMark) {
        const val = (bdoMark.attrs?.val === "rtl" ? "rtl" : "ltr") as "ltr" | "rtl";
        finalLink = { bdo: { val, children: [finalLink] } } as ParagraphChild;
      }
      children.push(finalLink);
      return;
    }
    if (trackMark) {
      // Wrap the run back into a w:ins/w:del/w:moveFrom/w:moveTo container — the reverse of
      // resolveTrackedChange. compileTrackedChangeRun returns the typed
      // ParagraphChild branch, so no cast is needed here. office-open's
      // stringifyDeletedRun emits <w:delText> for deletion children.
      const kind = trackMark.type as "insertion" | "deletion" | "moveFrom" | "moveTo";
      if (kind === "insertion" || kind === "deletion" || kind === "moveFrom" || kind === "moveTo") {
        let finalTrack = this.compileTrackedChangeRun(kind, trackMark.attrs, text, runOpts);
        if (dirMark) {
          const val = (dirMark.attrs?.val === "rtl" ? "rtl" : "ltr") as "ltr" | "rtl";
          finalTrack = { dir: { val, children: [finalTrack] } } as ParagraphChild;
        }
        if (bdoMark) {
          const val = (bdoMark.attrs?.val === "rtl" ? "rtl" : "ltr") as "ltr" | "rtl";
          finalTrack = { bdo: { val, children: [finalTrack] } } as ParagraphChild;
        }
        children.push(finalTrack);
        return;
      }
    }

    let finalRun: ParagraphChild = runOpts as RunOptions;
    if (dirMark) {
      const val = (dirMark.attrs?.val === "rtl" ? "rtl" : "ltr") as "ltr" | "rtl";
      finalRun = { dir: { val, children: [finalRun] } } as ParagraphChild;
    }
    if (bdoMark) {
      const val = (bdoMark.attrs?.val === "rtl" ? "rtl" : "ltr") as "ltr" | "rtl";
      finalRun = { bdo: { val, children: [finalRun] } } as ParagraphChild;
    }
    children.push(finalRun);
  }

  /**
   * Wrap a run back into a w:ins/w:del/w:moveFrom/w:moveTo container — the reverse of
   * resolveTrackedChange. A literal-key ternary (`{insertion: body}` /
   * `{deletion: body}`) lets TS narrow to the ParagraphChild branch without a
   * cast, and `typeof` guards read `attrs` type-safely (no `as number/string`).
   * stringifyDeletedRun emits `<w:delText>` automatically for deletion children.
   */
  private compileTrackedChangeRun(
    type: "insertion" | "deletion" | "moveFrom" | "moveTo",
    attrs: Record<string, unknown> | undefined,
    text: string,
    runOpts: Record<string, unknown>,
  ): ParagraphChild {
    const { text: _, children: __, ...runWithoutText } = runOpts;
    const trackChildren: (RunOptions | string)[] = [];
    if (text.includes("\u2011") || text.includes("\u00AD")) {
      trackChildren.push({ ...runWithoutText, children: splitTextHyphens(text) } as RunOptions);
    } else if (text) {
      trackChildren.push({ ...runWithoutText, text } as RunOptions);
    }
    const id = typeof attrs?.id === "number" ? attrs.id : 0;
    const author = typeof attrs?.author === "string" ? attrs.author : "";
    const date = typeof attrs?.date === "string" ? attrs.date : "";
    const body = { id, author, date, children: trackChildren };
    if (type === "insertion") return { insertion: body };
    if (type === "deletion") return { deletion: body };
    if (type === "moveFrom") return { movedFrom: body } as unknown as ParagraphChild;
    return { movedTo: body } as unknown as ParagraphChild;
  }

  /** Wrap a run back into a w:ruby container — the reverse of resolveRuby.
   *  w:ruby lives INSIDE its w:r (a run-children member, like w:tab), so the
   *  emitted shape is a text-less run `{ children: [{ ruby }] }`. The
   *  annotation text is flat on the mark; the base run keeps its other rPr
   *  props. Null attrs fall back to Word's conventions (half-size ruby,
   *  raised by its own size, centered). */
  private compileRubyRun(
    attrs: Record<string, unknown> | undefined,
    text: string,
    runOpts: Record<string, unknown>,
  ): ParagraphChild {
    const { text: _, ...runWithoutText } = runOpts;
    const baseChildren: (RunOptions | string)[] = [];
    if (text) baseChildren.push({ ...runWithoutText, text } as RunOptions);
    const fontSize = typeof attrs?.fontSize === "number" ? attrs.fontSize : 5;
    const baseFontSize = typeof attrs?.baseFontSize === "number" ? attrs.baseFontSize : 10;
    const annotation = typeof attrs?.text === "string" ? attrs.text : "";
    return {
      children: [
        {
          ruby: {
            properties: {
              alignment: (attrs?.alignment as RubyPropertiesOptions["alignment"]) ?? "center",
              fontSize,
              raise: typeof attrs?.raise === "number" ? attrs.raise : fontSize,
              baseFontSize,
              languageId: (attrs?.languageId as string) ?? "zh-CN",
              ...(attrs?.dirty ? { dirty: true } : {}),
            },
            text: { children: [{ text: annotation }] },
            base: { children: baseChildren },
          },
        },
      ],
    } as RunOptions;
  }

  // ── Resolve: DocumentOptions → Tiptap JSON ──

  private resolveSectionChild(child: SectionChild): JSONContent | null {
    // Declarative block dispatch: each block extension's parseDocxBlock rule
    // (collected in docxExtensions order) gets a chance to recognize the shape.
    // table/toc own their shapes here; a non-matching or null-converting
    // rule falls through. The shapes are mutually exclusive (different
    // SectionChild keys), so order among them is irrelevant in practice.
    const ctx = this.resolveCtx!;
    for (const { rule } of this.blockRules) {
      if (rule.match(child, ctx)) {
        const node = rule.convert(child, ctx);
        if (node) return node;
      }
    }
    // paragraph is not a block rule — the paragraph subtype resolver and the
    // plain-paragraph fallback own it.
    if ("paragraph" in child) {
      return this.resolveParagraph(child.paragraph);
    }
    // rawXml (incl. aggregated TOC field), generic SDT, bookmarkStart/End,
    // altChunk, subDoc, customXml — no native Tiptap node. Carry the
    // SectionChild verbatim so the round-trip is byte-faithful.
    return this.resolvePassthrough(child);
  }

  /** Wrap an opaque SectionChild in a passthrough atom (attrs.data = the
   *  branch's JSON-safe encoding — see {@link encodePassthroughData}). */
  private resolvePassthrough(child: SectionChild): JSONContent {
    return { type: "passthrough", attrs: { data: encodePassthroughData(child) } };
  }

  private resolveParagraph(opts: string | ParagraphOptions): JSONContent {
    const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;

    // Plain paragraph: reflective attrs parse + inline content. A paragraph
    // subtype (heading/code) is just attrs on this node.
    return buildTextBlock("paragraph", resolved, this.resolveCtx!);
  }

  /**
   * Walk a SectionChild[] block stream — a section's body, a header/footer
   * slot, or a table cell's children (a cell is just another block stream).
   */
  private resolveSectionChildren(children: SectionChild[]): JSONContent[] {
    const content: JSONContent[] = [];
    for (const child of children) {
      const node = this.resolveSectionChild(child);
      if (!node) continue;
      pushAll(content, node);
    }
    return content;
  }

  // ── Inline content resolution ──

  /**
   * Resolve a paragraph's inline content. @office-open collapses a plain-text
   * paragraph (a single run with no properties) to a bare string or a `{ text }`
   * object with no `children` — recover that text here so it round-trips back to
   * a text node instead of being dropped.
   */
  private resolveInlineContent(opts: ParagraphOptions): JSONContent[] {
    const content = this.resolveParagraphChildren(opts.children);
    if (content.length === 0 && opts.text) {
      // The collapsed plain-text fallback is a bare run: paragraph-level
      // options (pStyle, alignment, and especially the w:pPrChange `revision`)
      // are NOT run properties. Dropping `revision` here keeps a paragraph
      // format change paragraph-only — otherwise resolveMarks would stamp a
      // phantom run-level formatChange (and compile an empty w:rPrChange).
      const { revision: _, ...runOpts } = opts;
      const marks = this.resolveMarks(runOpts as unknown as RunOptions);
      const node: JSONContent = { type: "text", text: opts.text };
      if (marks) node.marks = marks;
      return [node];
    }
    return content;
  }

  private resolveParagraphChildren(children?: (ParagraphChild | string)[]): JSONContent[] {
    if (!children || children.length === 0) return [];

    const nodes: JSONContent[] = [];
    for (const child of children) {
      if (typeof child === "string") {
        if (child) nodes.push({ type: "text", text: child });
        continue;
      }
      if (typeof child === "object" && child !== null) {
        const resolved = this.resolveParagraphChild(child);
        if (resolved) pushAll(nodes, resolved);
      }
    }
    return nodes;
  }

  private resolveParagraphChild(child: ParagraphChild): JSONContent | JSONContent[] | null {
    // Declarative inline dispatch: each inline node/mark extension's
    // parseDocxInline rule (collected in docxExtensions order) gets a chance to
    // recognize the shape. tab/image/wpgGroup/wpsShape/hyperlink/
    // insertion/deletion/pageBreak/columnBreak own their shapes here; a
    // non-matching or null-converting rule falls through. The shapes are mutually
    // exclusive (different ParagraphChild keys), so order among them is
    // irrelevant in practice.
    const ctx = this.resolveCtx!;
    for (const { rule } of this.inlineRules) {
      if (rule.match(child, ctx)) {
        const node = rule.convert(child, ctx);
        if (node) return node;
      }
    }
    if ("dir" in child && (child as Record<string, unknown>).dir) {
      const d = (child as Record<string, unknown>).dir as {
        val?: string;
        children?: (ParagraphChild | string)[];
      };
      const inner = this.resolveParagraphChildren(d.children);
      for (const node of inner) {
        if (!node.marks) node.marks = [];
        node.marks.push({ type: "dir", attrs: { val: d.val ?? "ltr" } });
      }
      return inner;
    }
    if ("bdo" in child && (child as Record<string, unknown>).bdo) {
      const b = (child as Record<string, unknown>).bdo as {
        val?: string;
        children?: (ParagraphChild | string)[];
      };
      const inner = this.resolveParagraphChildren(b.children);
      for (const node of inner) {
        if (!node.marks) node.marks = [];
        node.marks.push({ type: "bdo", attrs: { val: b.val ?? "ltr" } });
      }
      return inner;
    }
    // run catch-all: a plain run (text/children/break). Left in the manager — it
    // is the fallback every non-owned shape reaches, not an owned shape itself.
    if ("text" in child || "children" in child || "break" in child) {
      return this.resolveRun(child as RunOptions);
    }
    // Non-textbox VML shapes in <w:pict> are promoted to structured wpsShape
    if (
      "pict" in child &&
      child.pict &&
      isPromotableVmlPict(child.pict as Record<string, unknown>)
    ) {
      const promoted = promoteVmlPictToWpsShape(child.pict as Record<string, unknown>);
      if (promoted) return promoted;
    }
    // Any remaining inline shape (an inline SDT, bookmark/range
    // markers, proofErr, …) carries verbatim via inlinePassthrough so the
    // round-trip stays byte-faithful — mirrors block resolvePassthrough, which
    // keeps every unrecognized SectionChild instead of dropping it. An
    // inline SDT used to be dropped here; carrying it restores the symmetry.
    return { type: "inlinePassthrough", attrs: { data: encodePassthroughData(child) } };
  }

  private resolveRun(opts: RunOptions): JSONContent | JSONContent[] | null {
    // Pure break (no text/children) → hardBreak node
    if (opts.break && opts.text === undefined && !opts.children) {
      return { type: "hardBreak" };
    }
    if (
      (opts as Record<string, unknown>).carriageReturn &&
      opts.text === undefined &&
      !opts.children
    ) {
      return { type: "hardBreak", attrs: { variant: "carriageReturn" } };
    }
    const text = opts.text;
    if (text === undefined && !opts.children) return null;

    // office-open 0.10.11+ may nest a run's inline elements under run.children
    // (empty run elements like tab/date/lastRenderedPageBreak, or mixed breaks).
    // Walk the children so pageBreak/columnBreak/hardBreak atoms are still
    // yielded; text fragments join into one text node carrying the run's marks.
    if (opts.children) {
      const marks = this.resolveMarks(opts);
      const nodes: JSONContent[] = [];
      let parts: string[] = [];
      const flushText = () => {
        if (parts.length > 0) {
          const node: JSONContent = { type: "text", text: parts.join("") };
          if (marks) node.marks = marks;
          nodes.push(node);
          parts = [];
        }
      };
      for (const c of opts.children) {
        if (typeof c === "string") {
          parts.push(c);
        } else if (c && typeof c === "object") {
          if ("noBreakHyphen" in c && (c as Record<string, unknown>).noBreakHyphen) {
            parts.push("\u2011");
            continue;
          }
          if ("softHyphen" in c && (c as Record<string, unknown>).softHyphen) {
            parts.push("\u00AD");
            continue;
          }
          if ("carriageReturn" in c && (c as Record<string, unknown>).carriageReturn) {
            flushText();
            nodes.push({ type: "hardBreak", attrs: { variant: "carriageReturn" } });
            continue;
          }
          if ("dir" in c && (c as Record<string, unknown>).dir) {
            flushText();
            const d = (c as Record<string, unknown>).dir as {
              val?: string;
              children?: (ParagraphChild | string)[];
            };
            const inner = this.resolveParagraphChildren(d.children);
            for (const node of inner) {
              if (!node.marks) node.marks = [];
              node.marks.push({ type: "dir", attrs: { val: d.val ?? "ltr" } });
            }
            pushAll(nodes, inner);
            continue;
          }
          if ("bdo" in c && (c as Record<string, unknown>).bdo) {
            flushText();
            const b = (c as Record<string, unknown>).bdo as {
              val?: string;
              children?: (ParagraphChild | string)[];
            };
            const inner = this.resolveParagraphChildren(b.children);
            for (const node of inner) {
              if (!node.marks) node.marks = [];
              node.marks.push({ type: "bdo", attrs: { val: b.val ?? "ltr" } });
            }
            pushAll(nodes, inner);
            continue;
          }
          // Reflective: reuse the inlineRules dispatch (same as top-level
          // ParagraphChild) so tab/pageBreak/columnBreak — and any custom
          // inline atom — are recognized here too. This replaces a parallel
          // if-in chain that duplicated those rules and dropped shapes that
          // lacked a hardcoded branch (e.g. a custom inline node nested in a
          // run was silently lost).
          const ctx = this.resolveCtx!;
          let handled = false;
          for (const { rule } of this.inlineRules) {
            if (rule.match(c as ParagraphChild, ctx)) {
              const node = rule.convert(c as ParagraphChild, ctx);
              if (node) {
                flushText();
                pushAll(nodes, node);
                handled = true;
                break;
              }
            }
          }
          if (handled) continue;
          // hardBreak: a bare `<w:br/>` inside run.children — no owning inline
          // rule (HardBreak declares no parseDocxInline), stays a manager case.
          if ("break" in c) {
            flushText();
            nodes.push({ type: "hardBreak" });
          }
          // {lastRenderedPageBreak} is a Word render hint — drop (office-open
          // does not emit it on output). date fields/separator/pgNum
          // are unsupported inline elements, dropped for now.
        }
      }
      flushText();
      if (text !== undefined) {
        // A run may carry opts.text alongside children (rare); fold it in.
        const node: JSONContent = { type: "text", text, marks };
        nodes.push(node);
      }
      if (nodes.length === 0) return null;
      return nodes.length === 1 ? nodes[0] : nodes;
    }

    return { type: "text", text: text ?? "", marks: this.resolveMarks(opts) };
  }

  private resolveMarks(opts: RunOptions): JSONContent["marks"] {
    const marks: NonNullable<JSONContent["marks"]> = [];

    // Each mark extension's parseDocx returns its attrs, or null when the run
    // does not carry the mark. The code/textStyle coupling (rStyle "CodeChar"
    // belongs to `code`; Consolas font and the styleId are skipped inside
    // textStyle.parseDocx) is handled within those extensions.
    for (const { name, parse } of this.markParse) {
      const attrs = parse(opts);
      if (attrs === null) continue;
      marks.push(Object.keys(attrs).length ? { type: name, attrs } : { type: name });
    }

    return marks.length > 0 ? marks : undefined;
  }

  /**
   * The marks a run's rPr options imply (name + attrs) — the reverse of
   * {@link runPropsFromMarks}, reusing the same parse hooks `resolveMarks`
   * walks. The editor's format-change reject uses it to restore the old rPr
   * snapshot stored in a w:rPrChange.
   */
  runPropsToMarks(props: Record<string, unknown>): RunPropMark[] {
    const marks: RunPropMark[] = [];
    for (const { name, parse } of this.markParse) {
      // The revision carrier is not an rPr mark — restore reset it separately.
      if (name === FORMAT_CHANGE_MARK) continue;
      const attrs = parse(props as RunOptions);
      if (attrs === null) continue;
      marks.push(Object.keys(attrs).length ? { type: name, attrs } : { type: name });
    }
    return marks;
  }

  /** Merge marks' renderDocx outputs into one rPr options object — the same
   *  overlay compileTextRun performs for a run's marks. */
  runPropsFromMarks(marks: readonly RunPropMark[]): Record<string, unknown> {
    const props: Record<string, unknown> = {};
    for (const mark of marks) {
      const render = this.markRender.get(mark.type);
      if (render) Object.assign(props, render(mark.attrs ?? {}));
    }
    return props;
  }

  /** Names of the marks carrying w:rPr run properties — the set a
   *  format-change reject replaces wholesale (Word: the old rPr wins as a
   *  whole, so a prop absent from it disappears). The formatChange revision
   *  carrier is excluded — review commands remove it themselves. */
  formatMarkNames(): string[] {
    return this.markParse.filter((m) => m.name !== FORMAT_CHANGE_MARK).map((m) => m.name);
  }
}

// ── Standalone functions (backward compat) ──

const defaultManager = new DocxManager(docxExtensions);

/** Resolve a DocxManager for a conversion call: the shared default singleton
 *  when no extensions are given, or a fresh instance bound to custom extensions
 *  so user-supplied marks/nodes plug into compile/resolve without a fork. */
function getDocxManager(extensions?: Extensions): DocxManager {
  return extensions ? new DocxManager(extensions) : defaultManager;
}

/** The marks a run's rPr options imply (name + attrs). See
 *  {@link DocxManager.runPropsToMarks} — the format-change restore helper. */
export function runPropsToMarks(
  props: Record<string, unknown>,
  extensions?: Extensions,
): RunPropMark[] {
  return getDocxManager(extensions).runPropsToMarks(props);
}

/** Merge run-prop marks' renderDocx outputs into one rPr options object (the
 *  inverse of {@link runPropsToMarks}). */
export function runPropsFromMarks(
  marks: readonly RunPropMark[],
  extensions?: Extensions,
): Record<string, unknown> {
  return getDocxManager(extensions).runPropsFromMarks(marks);
}

/** Names of the marks carrying w:rPr run properties — the set a format-change
 *  reject replaces. See {@link DocxManager.formatMarkNames}. */
export function formatMarkNames(extensions?: Extensions): string[] {
  return getDocxManager(extensions).formatMarkNames();
}

/**
 * Parse a DOCX file into Tiptap JSON (runtime model).
 *
 * Combines @office-open/docx's `parseDocument` (DOCX binary → DocumentOptions)
 * with `DocxManager.resolve` (DocumentOptions → Tiptap JSON). Async since
 * office-open 0.14, so `Blob` (including `File`) and `ReadableStream` inputs
 * are accepted alongside raw bytes; `parseDOCXSync` covers synchronous bytes.
 *
 * The input is untrusted: before parsing, the ZIP package is validated against
 * {@link ARCHIVE_LIMITS} (entry/size/ratio/media caps, bounded actual
 * inflation and XML budgets — see converters/archive-guard.ts) and throws
 * {@link ArchiveRejection} on violation.
 */
export async function parseDOCX(
  data: Parameters<typeof parseDocument>[0],
  extensions?: Extensions,
): Promise<JSONContent> {
  const bytes = await normalizeArchiveInput(data);
  assertNotEncryptedContainer(bytes);
  assertArchiveWithinLimits(bytes);
  const parsed = await parseDocument(bytes);
  // Belt-and-braces: a container office-open recognized as encrypted without
  // the CFB signature (or a future input path we do not read bytes for).
  if (parsed.encrypted) throw new EncryptedDocumentError();
  return getDocxManager(extensions).resolve(parsed);
}

/**
 * Synchronous counterpart of {@link parseDOCX} for already-normalized bytes —
 * `Blob` and `ReadableStream` inputs throw (use the async entry). The same
 * untrusted-input admission guard applies.
 */
export function parseDOCXSync(
  data: Parameters<typeof parseDocumentSync>[0],
  extensions?: Extensions,
): JSONContent {
  const bytes = normalizeArchiveInputSync(data);
  assertNotEncryptedContainer(bytes);
  assertArchiveWithinLimits(bytes);
  const parsed = parseDocumentSync(bytes);
  if (parsed.encrypted) throw new EncryptedDocumentError();
  return getDocxManager(extensions).resolve(parsed);
}

/**
 * DOCX package variants — the macro-enabled and template siblings that share
 * the OOXML part layout but declare a different main document part content
 * type (ECMA-376 Part 1 §11.3.10 + the MS-OFFMACRO main types). Word keeps the
 * document part at `word/document.xml` in every variant; only the
 * `[Content_Types].xml` Override changes.
 */
export type DocxVariant = "docx" | "docm" | "dotx" | "dotm";

/** `[Content_Types].xml` Override for `/word/document.xml`, per variant. */
const MAIN_DOCUMENT_CONTENT_TYPES: Record<DocxVariant, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  docm: "application/vnd.ms-word.document.macroEnabled.main+xml",
  dotx: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
  dotm: "application/vnd.ms-word.template.macroEnabled.main+xml",
};

/**
 * Content-type overrides for the parts office-open's writer treats as present
 * on a fresh compile (registry `presence.flag === "freshCompile"`).
 *
 * The writer reads `DocxWriteContext.hasNumbering` / `hasFootnotes` /
 * `hasEndnotes` off the injected `[Content_Types].xml` table whenever one is
 * present (`!options.contentTypes || sourceOverrides.some(...)`), so a variant
 * stamp built from scratch must declare those parts — otherwise the model's
 * `word/numbering.xml` / `word/footnotes.xml` / `word/endnotes.xml` are never
 * emitted while `word/document.xml` keeps referencing them (dangling `w:numId`
 * and note references). `/word/document.xml` is excluded: the variant's own
 * main override is appended last.
 */
const FRESH_COMPILE_CONTENT_TYPES: readonly ContentTypeOverride[] =
  PART_REGISTRIES.docx.parts.flatMap((part) =>
    part.contentType !== undefined &&
    part.path !== "word/document.xml" &&
    part.presence.kind === "conditional" &&
    part.presence.flag === "freshCompile"
      ? [{ partName: `/${part.path}`, contentType: part.contentType }]
      : [],
  );

/**
 * Stamp a variant's main-part content type onto compiled options.
 *
 * office-open's content-type merge keeps a surviving source Override for a
 * part it rebuilds, so a save that requests a variant must REPLACE the
 * source's `/word/document.xml` override rather than add a second one — while
 * merging (not replacing) the rest keeps every other source declaration
 * (macro/ole parts, theme, footnotes) intact. An explicit `docx` request
 * therefore flips a macro-enabled/template source back to the standard
 * document main type (bytes must agree with a .docx name/MIME); only an
 * absent `variant` leaves the source main type untouched (source-faithful
 * round-trip — see the "no variant requested" case).
 *
 * A source-less compile (`compiled.contentTypes === undefined`) has no table
 * to merge, and injecting a partial one disables office-open's fresh-compile
 * defaults: its writer gates `word/numbering.xml` / `word/footnotes.xml` /
 * `word/endnotes.xml` on the table declaring them, so the table is seeded with
 * the registry's fresh-compile declarations first — the variant output then
 * carries the same parts a plain (no-variant) fresh compile would.
 */
function applyVariant(
  compiled: DocumentOptions,
  variant: DocxVariant | undefined,
): DocumentOptions {
  if (!variant) return compiled;
  const source = compiled.contentTypes;
  return {
    ...compiled,
    contentTypes: {
      defaults: source?.defaults ?? [],
      overrides: [
        ...(source ? [] : FRESH_COMPILE_CONTENT_TYPES),
        ...(source?.overrides ?? []).filter(
          (o) => o.partName.toLowerCase() !== "/word/document.xml",
        ),
        { partName: "/word/document.xml", contentType: MAIN_DOCUMENT_CONTENT_TYPES[variant] },
      ],
    },
  };
}

/**
 * Options for {@link generateDOCX} / {@link generateDOCXStream}.
 */
export interface DocxGenerateOptions<T extends OutputType = "nodebuffer"> {
  /**
   * Package variant whose main document part content type is stamped on the
   * output — `docx`, `docm`, `dotx`, or `dotm`. Requesting `docx` flips a
   * macro-enabled/template source package back to the standard document main
   * type. Omit it to keep the source's own main type (source-faithful save).
   * Macro parts carried in `documentExtras.rawParts` stay in every variant.
   */
  variant?: DocxVariant;
  /**
   * Pre-compilation steps run on a copy of the JSON (default:
   * `[prepareImageSizes()]` — local-only, never touches the network).
   * - `true` / `undefined`: default local preparation (no network)
   * - `false`: skip preparation
   * - `PrepareStep[]`: custom steps (e.g.
   *   `[prepareImages({ allow: ["cdn.example.com"] }), prepareImageSizes()]`)
   *
   * External `http(s)` images require an explicit {@link prepareImages}
   * `allow` list — the default drops them at render time (see
   * extensions/image.ts). The input JSON is never mutated.
   */
  prepare?: boolean | PrepareStep[];
  /** Packer options; `type` controls the output format (default `"nodebuffer"` → Buffer). */
  packer?: PackerOptions<T>;
  /**
   * Document-level options injected into the compiled `DocumentOptions` — core
   * properties (`title`/`creator`/`description`/…), `styles`, `background`,
   * `features`, `fonts`, etc. Excludes `sections` (always compiled from the
   * JSON) and `numbering` (collected from ordered-list nodes).
   *
   * `styles` here takes precedence over any `styles` carried on
   * `json.attrs.styles` (e.g. from a prior `parseDOCX`).
   */
  document?: Omit<Partial<DocumentOptions>, "sections" | "numbering">;
  /**
   * Extension list used to build the conversion registry (default:
   * `docxExtensions`). Pass `[...docxExtensions, MyMark]` to plug a custom
   * mark/node into compile/resolve via its renderDocx/parseDocx hooks.
   */
  extensions?: Extensions;
  /**
   * Fixed clock for this generation — the reproducibility switch.
   *
   * Every date the document does not already carry (core properties
   * `created`/`modified`, new comment dates) resolves to this value, and every
   * generated id sequence (drawing/shape/SmartArt ids, font keys, altChunk
   * part names) restarts deterministically, so two generations of the same
   * input are byte-identical — same process or a fresh one.
   *
   * - `string | Date`: use this timestamp.
   * - `undefined` (default): {@link DOCX_EPOCH} (`1980-01-01T00:00:00.000Z`).
   * - `null`: omit generated dates entirely (source-carried dates still win).
   *
   * Dates/ids the source document already carries are preserved — this only
   * replaces values the generator would otherwise take from the wall clock,
   * `crypto.randomUUID()` or process-global counters.
   */
  date?: string | Date | null;
}

/**
 * Merge {@link DocxGenerateOptions.document} into the `DocumentOptions`
 * compiled from Tiptap JSON.
 *
 * - `sections`/`numbering`: compile-owned (excluded from `document`, never
 *   overridden).
 * - `styles`: option wins over `json.attrs.styles`.
 * - Everything else (core properties, background, features, …): injected.
 */
function applyDocumentOptions(
  base: DocumentOptions,
  document?: Omit<Partial<DocumentOptions>, "sections" | "numbering">,
): DocumentOptions {
  if (!document) return base;
  return { ...base, ...document };
}

/**
 * Fill core-properties dates the compiled options do not carry with the
 * generation clock. `undefined` means "the source carried no value" and gets
 * the fixed date (or `null` to omit); an explicit source value — including
 * `null`, office-open's "source had no date" — is preserved.
 */
function applyGenerationDate(
  compiled: DocumentOptions,
  date: string | Date | null | undefined,
): DocumentOptions {
  if (compiled.created !== undefined && compiled.modified !== undefined) return compiled;
  const fallback = date === null ? null : toIsoDate(date ?? DOCX_EPOCH);
  return {
    ...compiled,
    created: compiled.created !== undefined ? compiled.created : fallback,
    modified: compiled.modified !== undefined ? compiled.modified : fallback,
  };
}

/** The scope clock for a `date` option — `null` opens the scope date-less. */
function scopeDate(date: string | Date | null | undefined): string | undefined {
  return date === null ? undefined : toIsoDate(date ?? DOCX_EPOCH);
}

/**
 * Generate a DOCX file from Tiptap JSON (runtime model), asynchronously.
 *
 * Pipeline: `prepareDocument` (default: local preparation on a copy; no
 * network) → `DocxManager.compile` → @office-open/docx's `generateDocument`.
 * `packer.type` controls the output format (default: `"nodebuffer"` → Buffer).
 * Non-blocking (fflate Web Workers). The input `json` is never mutated; see
 * `options.prepare` for fetching external images and `options.date` for the
 * fixed reproducibility clock (default {@link DOCX_EPOCH}).
 */
export async function generateDOCX<T extends OutputType = "nodebuffer">(
  json: JSONContent,
  options?: DocxGenerateOptions<T>,
): Promise<OutputByType[T]> {
  const { prepare = true, packer, document, extensions, variant, date } = options ?? {};
  const prepared =
    prepare === false ? json : await prepareDocument(json, prepare === true ? undefined : prepare);
  const compiled = applyGenerationDate(
    applyVariant(applyDocumentOptions(compileDocument(prepared, extensions), document), variant),
    date,
  );
  return withGenerationScope(scopeDate(date), () => generateDocument(compiled, packer));
}

/**
 * Generate a DOCX file synchronously — fastest throughput, blocks the event loop.
 *
 * Pipeline: `DocxManager.compile` → `generateDocumentSync`. Does **not** run
 * `prepareDocument` (it is async); call `await prepareDocument(json)` first
 * when images need embedding. `options.document` and `options.date` are
 * applied; the input `json` is never mutated.
 */
export function generateDOCXSync<T extends OutputType = "nodebuffer">(
  json: JSONContent,
  options?: DocxGenerateOptions<T>,
): OutputByType[T] {
  const { packer, document, extensions, variant, date } = options ?? {};
  const compiled = applyGenerationDate(
    applyVariant(applyDocumentOptions(compileDocument(json, extensions), document), variant),
    date,
  );
  return withGenerationScope(scopeDate(date), () => generateDocumentSync(compiled, packer));
}

/**
 * Generate a DOCX file as a `ReadableStream<Uint8Array>` — for large documents
 * or streaming HTTP responses.
 *
 * Pipeline: `prepareDocument` (default: local preparation on a copy; no
 * network) → `DocxManager.compile` → `generateDocumentStream`. Async due to
 * preparation; `options.date` controls the fixed reproducibility clock.
 */
export async function generateDOCXStream(
  json: JSONContent,
  options?: DocxGenerateOptions,
): Promise<ReadableStream<Uint8Array>> {
  const { prepare = true, packer, document, extensions, variant, date } = options ?? {};
  const prepared =
    prepare === false ? json : await prepareDocument(json, prepare === true ? undefined : prepare);
  const compiled = applyGenerationDate(
    applyVariant(applyDocumentOptions(compileDocument(prepared, extensions), document), variant),
    date,
  );
  return withGenerationScope(scopeDate(date), () => generateDocumentStream(compiled, packer));
}

/**
 * Convert DocumentOptions (persistence model) to Tiptap JSON (runtime model).
 */
export function resolveDocument(docOpts: DocumentOptions, extensions?: Extensions): JSONContent {
  return getDocxManager(extensions).resolve(docOpts);
}

/**
 * Convert Tiptap JSON (runtime model) to DocumentOptions (persistence model).
 */
export function compileDocument(
  json: JSONContent,
  extensions?: Extensions,
  cache?: CompileCache,
): DocumentOptions {
  return getDocxManager(extensions).compile(json, cache);
}

/**
 * Fill in the document-level defaults that a hand-built JSON lacks.
 *
 * A document constructed by hand (not via {@link parseDOCX}) carries no
 * `doc.attrs.styles` (docDefaults: body font/size/spacing + the built-in style
 * table) and no `doc.attrs.sectionProperties` (page size, margins, docGrid
 * linePitch). Those are injected only by office-open's `parseDocument`, which a
 * hand-built doc bypasses — so without them the editor has no body font, no
 * page geometry, and no document grid for snapToGrid to pitch against, and
 * rendering/pagination drift.
 *
 * Harvests the style table by round-tripping an EMPTY document through
 * office-open (`generateDOCXSync` → `parseDOCX`) and takes that attr plus
 * docen's own {@link docenDefaultSectionProperties} — the empty doc's
 * remaining attrs (documentExtras with passthrough binaries, settings,
 * contentTypes) are round-trip artifacts a hand-built doc must not inherit:
 * rawParts carries Uint8Array bytes that break JSON serialization of the
 * attrs (a host embedding `JSON.stringify(normalizeDocument(...))` then
 * crashes the next save-as in office-open's media reader). Content nodes
 * (paragraphs/runs/marks) pass through verbatim, avoiding the mark pollution a
 * full-content round-trip would cause (a paragraph's default run props leak
 * onto its text as a textStyle mark). Keys already set on `json.attrs` win, so
 * a doc that already carries its own styles/section properties (e.g. from
 * `parseDOCX` or a prior `getJSON`) is left unchanged.
 */
export function normalizeDocument(json: JSONContent, extensions?: Extensions): JSONContent {
  // parseDOCXSync, not parseDOCX: this is a synchronous public API (demo and
  // editor call it inline), and the harvest input is always normalized bytes.
  const defaults = getDocxManager(extensions).resolve(
    parseDocumentSync(generateDOCXSync({ type: "doc", content: [] }, { extensions })),
  );
  const baseAttrs = (defaults.attrs ?? {}) as Record<string, unknown>;
  // A hand-built doc (e.g. parseHTML output) carries attrs keys with null
  // values (schema defaults) — those are "lacking", not overrides, so they
  // must not shadow the harvested defaults.
  const userAttrs = Object.fromEntries(
    Object.entries((json.attrs ?? {}) as Record<string, unknown>).filter(([, v]) => v != null),
  );
  const harvested = {
    styles: baseAttrs.styles,
    // Our own explicit geometry — never the empty round-trip's (absent)
    // sectionProperties, which made a hand-built doc depend on office-open's
    // implicit zh-CN `sectionMarginDefaults` for its page box.
    sectionProperties: docenDefaultSectionProperties(),
  };
  return { ...json, attrs: { ...harvested, ...userAttrs } };
}
