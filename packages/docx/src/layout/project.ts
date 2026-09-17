// The docx adapter — projects office-open's DocumentOptions into
// @docen/layout's LayoutDoc. The PERSISTENCE model is the projection source
// (not the editor's Tiptap JSON subset): every body shape office-open can
// round-trip reaches the layout engine, and shapes the adapter cannot lay out
// yet (toc, sdt, textbox, altChunk, customXml, rawXml) become placeholder
// boxes instead of silently vanishing. Callers chain
// Tiptap JSON --compileDocument--> DocumentOptions --this--> LayoutDoc.
//
// Zero-DOM discipline is inherited from @docen/layout — this module is
// Node-safe (headless export) by construction.

import type {
  LayoutBlock,
  ProjectedColumns,
  ProjectedFlowBox,
  ProjectedLineNumbers,
  ProjectedPageBackground,
  ProjectedPageBorders,
  ProjectedPageFurniture,
  ProjectedPageNumbering,
} from "@docen/layout";
import { twipToPx } from "@docen/layout";
import type { DocumentOptions, SectionPropertiesOptions } from "@office-open/docx";

import { indexCharacterStyles } from "../style-cascade";
import type { MarkupDisplay, ProjectContext } from "./project/context";
import { isRecord, type BodyParagraph } from "./project/guards";
import { indexNumberings } from "./project/numbering";
import {
  inheritFurnitureSlots,
  projectChild,
  projectColumns,
  projectFlowBox,
  projectLineNumbers,
  projectPageBackground,
  projectPageBorders,
  projectPageFurniture,
  projectPageNumbering,
} from "./project/page";
import { projectParagraph } from "./project/paragraph";
import { projectTable } from "./project/table";

export { projectFlowBox } from "./project/page";

export interface ProjectedSection {
  blocks: LayoutBlock[];
  flow: ProjectedFlowBox;
  furniture: ProjectedPageFurniture;
  /** The section's page borders (w:pgBorders), absent when none. */
  pageBorders?: ProjectedPageBorders;
  /** The section's line numbering (w:lnNumType), absent when none. */
  lineNumbers?: ProjectedLineNumbers;
  /** The section's page numbering (w:pgNumType), absent when the section
   *  continues the previous one's decimal numbers. */
  pageNumbering?: ProjectedPageNumbering;
  /** The section's columns (w:cols), absent for a single-column section. */
  columns?: ProjectedColumns;
  /** The section break type (sectPr @w:type): a "continuous" section merges
   *  onto the previous section's flow instead of opening a fresh page. */
  type?: SectionPropertiesOptions["type"];
  /** Footnote id → definition blocks (absent when document has no footnotes). */
  footnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
  /** Endnote id → definition blocks (absent when document has no endnotes). */
  endnoteDefinitions?: Map<number, readonly LayoutBlock[]>;
  /** Endnote placement (w:pos): 'sectEnd' (at section end) or 'docEnd' (at document end). */
  endnotePlacement?: "sectEnd" | "docEnd";
}

/** One cached top-level child projection. `stateful` marks a projection that
 *  read or wrote order-dependent context state (list counters, note
 *  ordinals, comment ranges, revision author palette) — such an entry is
 *  always re-projected so its state mutations replay; `stateAfter` is the
 *  order-state snapshot leaving it, which proves whether everything after it
 *  still sees the cached run's state. A non-stateful entry depends only on
 *  its own child identity, so it is always reusable. */
interface ChildProjection {
  blocks: LayoutBlock[];
  stateful: boolean;
  stateAfter?: string;
}

/** The per-document projection record. Recreated whenever a projection input
 *  changes identity or value — the entries only describe the exact inputs
 *  they were built under. */
interface ProjectionRecord {
  /** Scalar inputs, serialized (markup flags, hyphenation, theme fonts, …). */
  key: string;
  styles: unknown;
  numbering: unknown;
  settings: unknown;
  commentsRef: unknown;
  characterStyles: ReturnType<typeof indexCharacterStyles>;
  numberings: ReturnType<typeof indexNumberings>;
  /** Top-level SectionChild identity → its projected blocks. */
  children: WeakMap<object, ChildProjection>;
  /** Section identity → the last section-level projection. */
  sections: WeakMap<
    object,
    {
      blocks: readonly LayoutBlock[];
      fields: Omit<
        ProjectedSection,
        "blocks" | "footnoteDefinitions" | "endnoteDefinitions" | "endnotePlacement"
      >;
      fieldsInputs: readonly unknown[];
      fnDefs?: Map<number, readonly LayoutBlock[]>;
      enDefs?: Map<number, readonly LayoutBlock[]>;
      placement?: ProjectedSection["endnotePlacement"];
      out: ProjectedSection;
    }
  >;
  /** Doc-level note definitions, reused while their sources and the ordinal
   *  state stay untouched. */
  defs?: {
    sources: readonly unknown[];
    fnDefs?: Map<number, readonly LayoutBlock[]>;
    enDefs?: Map<number, readonly LayoutBlock[]>;
  };
}

/** Identity memo for {@link projectDocumentOptions} — pass the same object
 *  across renders. Invalid inputs simply re-project; the cache never makes
 *  an unsafe reuse (see {@link ChildProjection}). */
export interface ProjectionCache {
  record?: ProjectionRecord;
  /** Diagnostic tally — the incremental-projection audit trail: `reuses` /
   *  `reruns` count per-child decisions, `fallbacks` counts document walks
   *  whose order-dependent state could not be proven unchanged (note
   *  definitions and section reuse then recompute from scratch). Never read
   *  by the projection itself. */
  stats: { reuses: number; reruns: number; fallbacks: number };
}

export function createProjectionCache(): ProjectionCache {
  return { stats: { reuses: 0, reruns: 0, fallbacks: 0 } };
}

function sameBlocks(a: readonly LayoutBlock[], b: readonly LayoutBlock[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Serialize the order-dependent projection state — the proof that everything
 *  after a re-projected stateful child still sees the cached run's state.
 *  Small maps (one entry per numbering reference / note / comment range /
 *  revision author), so the walk pays this only at stateful children. */
function stateSnapshot(ctx: ProjectContext): string {
  const parts: string[] = [];
  for (const [reference, counters] of ctx.listCounters) {
    parts.push(`L${reference}:${counters.join(".")}`);
  }
  for (const [id, ordinal] of ctx.footnoteOrdinals) parts.push(`F${id}=${ordinal}`);
  for (const [id, ordinal] of ctx.endnoteOrdinals) parts.push(`E${id}=${ordinal}`);
  if (ctx.openComments.size > 0) {
    parts.push(`C${[...ctx.openComments].sort((a, b) => a - b).join(".")}`);
  }
  for (const [author, slot] of ctx.revisionAuthorColors) parts.push(`A${author}=${slot}`);
  return parts.join("|");
}

function projectNoteBlocks(
  children: readonly unknown[],
  ctx: ProjectContext,
  defaultStyle: string,
): LayoutBlock[] {
  const blocks: LayoutBlock[] = [];
  for (const child of children) {
    if (typeof child === "string") {
      blocks.push(projectParagraph({ text: child, style: defaultStyle }, ctx));
    } else if (isRecord(child)) {
      if ("paragraph" in child) {
        const p = child.paragraph;
        const pObj =
          typeof p === "string"
            ? { text: p, style: defaultStyle }
            : isRecord(p) && !p.style
              ? { ...p, style: defaultStyle }
              : p;
        blocks.push(projectParagraph(pObj as BodyParagraph, ctx));
      } else if ("table" in child) {
        const t = projectTable(child.table as never, ctx);
        if (t) blocks.push(t);
      } else {
        const pObj = !child.style ? { ...child, style: defaultStyle } : child;
        blocks.push(projectParagraph(pObj as BodyParagraph, ctx));
      }
    }
  }
  return blocks;
}

/** Flatten a comment's children (office-open's CommentOptions) to plain text
 *  for the balloon body; paragraphs separate with a newline. */
function commentTextOf(children: readonly unknown[]): string {
  const parts: string[] = [];
  for (const child of children) {
    if (typeof child === "string") {
      parts.push(child);
      continue;
    }
    if (!isRecord(child)) continue;
    if (typeof child.text === "string") {
      parts.push(child.text);
      continue;
    }
    if (Array.isArray(child.children)) {
      const nested = commentTextOf(child.children);
      if (nested) parts.push(parts.length > 0 ? `\n${nested}` : nested);
    }
  }
  return parts.join("");
}

/** Project a full DocumentOptions into the engine's input: one
 *  {@link ProjectedSection} per document section plus the page background
 *  (document-wide). Sections paginate in order — see
 *  `layoutFlowSections` in @docen/layout. `markup` applies Word's Display for
 *  Review state to the tracked-changes projection (omitted = all marks show);
 *  `showFieldCodes` projects every field as its instruction text (Alt+F9);
 *  `showHiddenText` displays hidden runs (w:vanish) instead of suppressing
 *  them (Word's Options → Display toggle). */
export function projectDocumentOptions(
  doc: DocumentOptions,
  markup?: MarkupDisplay,
  showFieldCodes?: boolean,
  showHiddenText?: boolean,
  hyphenation?: { auto?: boolean; doNotHyphenateCaps?: boolean; zoneTw?: number; limit?: number },
  themeFonts?: { majorFont?: string; minorFont?: string },
  cache?: ProjectionCache,
): {
  sections: ProjectedSection[];
  background?: ProjectedPageBackground;
} {
  // The cache record is valid only under the exact inputs it was built with:
  // scalar flags by value, the style/numbering/comment models by identity.
  // Anything else starts a fresh record (the old entries die with it).
  const recordKey = [
    markup ? `${markup.view}\u0001${markup.colors ?? ""}\u0001${markup.balloons ?? ""}` : "",
    (markup?.authors ?? []).join("\u0001"),
    showFieldCodes ? "f" : "",
    showHiddenText ? "h" : "",
    hyphenation
      ? `${hyphenation.auto ?? ""}\u0001${hyphenation.doNotHyphenateCaps ?? ""}\u0001${
          hyphenation.zoneTw ?? ""
        }\u0001${hyphenation.limit ?? ""}`
      : "",
    themeFonts ? `${themeFonts.majorFont ?? ""}\u0001${themeFonts.minorFont ?? ""}` : "",
  ].join("\u0002");
  let record = cache?.record;
  if (
    record === undefined ||
    record.key !== recordKey ||
    record.styles !== doc.styles ||
    record.numbering !== doc.numbering ||
    record.settings !== doc.settings ||
    record.commentsRef !== doc.comments
  ) {
    record = {
      key: recordKey,
      styles: doc.styles,
      numbering: doc.numbering,
      settings: doc.settings,
      commentsRef: doc.comments,
      characterStyles: indexCharacterStyles(doc.styles),
      numberings: indexNumberings(doc.numbering),
      children: new WeakMap(),
      sections: new WeakMap(),
    };
    if (cache) cache.record = record;
  }

  // Comment balloon data: the compile pass spreads documentExtras.comments
  // into DocumentOptions.comments (w:comment entries with author/initials and
  // the thread body).
  const commentMeta = new Map<number, { author: string; initials: string; text: string }>();
  for (const comment of doc.comments ?? []) {
    if (typeof comment.id !== "number") continue;
    commentMeta.set(comment.id, {
      author: comment.author ?? "",
      initials: comment.initials ?? "",
      text: commentTextOf(comment.children ?? []),
    });
  }
  const fnProps = (doc.settings?.footnoteProperties ?? {}) as Record<string, unknown>;
  const enProps = (doc.settings?.endnoteProperties ?? {}) as Record<string, unknown>;
  const fnStart =
    typeof fnProps.numStart === "number" && fnProps.numStart >= 1 ? fnProps.numStart : undefined;
  const enStart =
    typeof enProps.numStart === "number" && enProps.numStart >= 1 ? enProps.numStart : undefined;

  // The order-dependent state a reused child projection could miss lives in
  // the per-call maps below; a projection that reads or advances them marks
  // itself stateful (see ChildProjection) and is then never reused blind.
  const stateful = { hit: false };
  const ctx: ProjectContext = {
    styles: doc.styles,
    ...(themeFonts ? { themeFonts } : {}),
    characterStyles: record.characterStyles,
    numberings: record.numberings,
    listCounters: new Map(),
    openComments: new Set(),
    footnoteOrdinals: new Map(),
    footnoteNumFmt: typeof fnProps.numFmt === "string" ? fnProps.numFmt : undefined,
    footnoteNumStart: fnStart,
    footnoteNumRestart: (fnProps.numRestart === "eachSect" || fnProps.numRestart === "eachPage"
      ? fnProps.numRestart
      : "continuous") as "continuous" | "eachSect" | "eachPage",
    endnoteOrdinals: new Map(),
    endnoteNumFmt: typeof enProps.numFmt === "string" ? enProps.numFmt : undefined,
    endnoteNumStart: enStart,
    endnoteNumRestart: (enProps.numRestart === "eachSect" || enProps.numRestart === "eachPage"
      ? enProps.numRestart
      : "continuous") as "continuous" | "eachSect" | "eachPage",
    revisionAuthorColors: new Map(),
    stateful,
    ...(commentMeta ? { commentMeta } : {}),
    ...(markup ? { markup } : {}),
    ...(showFieldCodes ? { showFieldCodes: true } : {}),
    ...(showHiddenText ? { showHiddenText: true } : {}),
    // The document-wide tab grid (w:defaultTabStop, twips); Word's 720 default
    // applies when settings omit it (the engine carries that fallback).
    defaultTabStopPx:
      doc.settings?.defaultTabStop != null && doc.settings.defaultTabStop > 0
        ? twipToPx(doc.settings.defaultTabStop)
        : undefined,
    autoHyphenation:
      hyphenation?.auto != null ? hyphenation.auto : doc.settings?.autoHyphenation === true,
    doNotHyphenateCaps:
      hyphenation?.doNotHyphenateCaps != null
        ? hyphenation.doNotHyphenateCaps
        : doc.settings?.doNotHyphenateCaps === true,
    hyphenationZoneTw:
      hyphenation?.zoneTw ??
      (typeof doc.settings?.hyphenationZone === "number"
        ? doc.settings.hyphenationZone
        : undefined),
    consecutiveHyphenLimit:
      hyphenation?.limit ??
      (typeof doc.settings?.consecutiveHyphenLimit === "number"
        ? doc.settings.consecutiveHyphenLimit
        : undefined),
  };

  const docSections = doc.sections ?? [];
  const sectionBlocks: LayoutBlock[][] = [];
  const sectionHit: boolean[] = [];
  // Whether the order-dependent state provably matches the cached run's
  // (proved by the last stateful child's post-state snapshot). Note
  // definitions — the only consumer that reads final ordinal state — reuse
  // only under a proof.
  let stateProvable = true;
  for (const [sIdx, section] of docSections.entries()) {
    if (sIdx > 0) {
      if (ctx.footnoteNumRestart === "eachSect") ctx.footnoteOrdinals.clear();
      if (ctx.endnoteNumRestart === "eachSect") ctx.endnoteOrdinals.clear();
    }
    const children = section.children ?? [];
    const blocks: LayoutBlock[] = [];
    let allHit = true;
    let lastChild: object | undefined;
    for (const child of children) {
      const key = child as object;
      const cached = record.children.get(key);
      // A non-stateful projection depends only on its own child identity —
      // reusable whatever the order state did.
      if (cached && !cached.stateful) {
        blocks.push(...cached.blocks);
        lastChild = key;
        if (cache) cache.stats.reuses++;
        continue;
      }
      allHit = false;
      if (cache) cache.stats.reruns++;
      stateful.hit = false;
      const block = projectChild(child, ctx);
      const out: LayoutBlock[] = [];
      if (Array.isArray(block)) out.push(...block);
      else if (block) out.push(block);
      const entry: ChildProjection = { blocks: out, stateful: stateful.hit };
      if (stateful.hit) {
        // The child re-projected (its mutations replayed). Proof of state
        // continuity is the post-state snapshot against the cached run's at
        // the same child; an unknown (never-seen) child cannot prove it.
        entry.stateAfter = stateSnapshot(ctx);
        stateProvable = cached?.stateAfter === entry.stateAfter;
      }
      record.children.set(key, entry);
      blocks.push(...out);
      lastChild = key;
    }
    // A non-final section's last paragraph carries the sectPr — Word paints
    // its mark row as "─────分节符(下一页)─────". The final section's sectPr
    // rides the body's end (no paragraph holds it) and shows no mark. The
    // marked copy replaces the block inside the child's cached array, so an
    // unchanged section keeps its exact block objects across renders.
    const last = blocks[blocks.length - 1];
    if (sIdx < docSections.length - 1 && last?.kind === "paragraph") {
      // The mark row names the break type (Word: "分节符(连续)") — nextPage
      // collapses to true, the painter's default label.
      const type = section.properties?.type;
      const marked: LayoutBlock = {
        ...last,
        sectionEnd:
          type === "continuous" || type === "evenPage" || type === "oddPage" ? type : true,
      };
      blocks[blocks.length - 1] = marked;
      const entry = lastChild != null ? record.children.get(lastChild) : undefined;
      if (entry && entry.blocks[entry.blocks.length - 1] === last) {
        entry.blocks[entry.blocks.length - 1] = marked;
      }
    }
    sectionBlocks.push(blocks);
    sectionHit.push(allHit);
  }

  // Note definitions ride the body walk's ordinal state; reuse them only when
  // no stateful projection re-ran and the source models are untouched.
  const defSources: readonly unknown[] = [doc.footnotes, doc.settings];
  const defsReusable =
    record.defs != null &&
    stateProvable &&
    record.defs.sources.every((source, i) => source === defSources[i]);
  if (cache && !stateProvable) cache.stats.fallbacks++;
  let fnDefs: Map<number, readonly LayoutBlock[]> | undefined;
  let enDefs: Map<number, readonly LayoutBlock[]> | undefined;
  if (defsReusable) {
    fnDefs = record.defs!.fnDefs;
    enDefs = record.defs!.enDefs;
  } else {
    const footnoteDefinitions = new Map<number, readonly LayoutBlock[]>();
    for (const note of doc.footnotes ?? []) {
      if (note.id == null) continue;
      const ordinal = ctx.footnoteOrdinals.get(note.id) ?? note.id;
      const noteCtx: ProjectContext = { ...ctx, currentNoteOrdinal: ordinal };
      const noteBlocks = projectNoteBlocks(note.children ?? [], noteCtx, "FootnoteText");
      footnoteDefinitions.set(note.id, noteBlocks);
    }
    const endnoteDefinitions = new Map<number, readonly LayoutBlock[]>();
    const docEndnotes = (
      doc as unknown as { endnotes?: Array<{ id?: number; children?: unknown[] }> }
    ).endnotes;
    for (const note of docEndnotes ?? []) {
      if (note.id == null) continue;
      const ordinal = ctx.endnoteOrdinals.get(note.id) ?? note.id;
      const noteCtx: ProjectContext = { ...ctx, currentNoteOrdinal: ordinal };
      const noteBlocks = projectNoteBlocks(note.children ?? [], noteCtx, "EndnoteText");
      endnoteDefinitions.set(note.id, noteBlocks);
    }
    fnDefs = footnoteDefinitions.size > 0 ? footnoteDefinitions : undefined;
    enDefs = endnoteDefinitions.size > 0 ? endnoteDefinitions : undefined;
    record.defs = { sources: defSources, fnDefs, enDefs };
  }
  const endnotePlacement: ProjectedSection["endnotePlacement"] =
    enDefs || enProps.pos
      ? ((enProps.pos === "sectEnd" ? "sectEnd" : "docEnd") as "sectEnd" | "docEnd")
      : undefined;

  // Word: a section without a header/footer reference shows the previous
  // section's — carry the effective slots forward so projection, page insets
  // and story bands all see the linked content.
  let prevHeaders: ReturnType<typeof inheritFurnitureSlots>;
  let prevFooters: ReturnType<typeof inheritFurnitureSlots>;
  const sections: ProjectedSection[] = [];
  for (const [i, section] of docSections.entries()) {
    const headers = inheritFurnitureSlots(section.headers, prevHeaders);
    const footers = inheritFurnitureSlots(section.footers, prevFooters);
    prevHeaders = headers;
    prevFooters = footers;
    const blocks = sectionBlocks[i] ?? [];
    const prevSec = record.sections.get(section);
    // Whole-section reuse: every child came from the cache, the assembled
    // list is element-identical to the previous output's, the doc-level note
    // definitions are the same objects, and no revision author color could
    // have shifted the palette. The section object itself survives.
    if (
      prevSec &&
      sectionHit[i] &&
      sameBlocks(prevSec.blocks, blocks) &&
      prevSec.fnDefs === fnDefs &&
      prevSec.enDefs === enDefs &&
      prevSec.placement === endnotePlacement
    ) {
      sections.push(prevSec.out);
      continue;
    }
    const fieldsInputs: readonly unknown[] = [
      section.properties,
      section.headers,
      section.footers,
      doc.settings,
    ];
    const fieldsReusable =
      prevSec != null &&
      ctx.revisionAuthorColors.size === 0 &&
      prevSec.fieldsInputs.every((input, k) => input === fieldsInputs[k]);
    const fields = fieldsReusable
      ? prevSec!.fields
      : {
          flow: {
            ...projectFlowBox(section.properties),
            // settings.xml compat: cell lines join the section's grid only when
            // the document declares w:adjustLineHeightInTable.
            adjustLinesInTable:
              typeof doc.settings?.compatibility === "object" &&
              doc.settings.compatibility.adjustLineHeightInTable === true,
          },
          furniture: projectPageFurniture(
            { ...section, headers, footers },
            doc,
            ctx.revisionAuthorColors,
          ),
          pageBorders: projectPageBorders(section.properties),
          lineNumbers: projectLineNumbers(section.properties),
          pageNumbering: projectPageNumbering(section.properties),
          columns: projectColumns(section.properties),
          type: section.properties?.type,
        };
    const out: ProjectedSection = {
      ...fields,
      blocks,
      footnoteDefinitions: fnDefs,
      endnoteDefinitions: enDefs,
      endnotePlacement,
    };
    record.sections.set(section, {
      blocks,
      fields,
      fieldsInputs,
      fnDefs,
      enDefs,
      placement: endnotePlacement,
      out,
    });
    sections.push(out);
  }

  return {
    sections:
      sections.length > 0
        ? sections
        : [
            {
              blocks: [],
              flow: projectFlowBox(undefined),
              furniture: projectPageFurniture(undefined, doc, ctx.revisionAuthorColors),
              pageBorders: undefined,
              lineNumbers: undefined,
              columns: undefined,
            },
          ],
    background: projectPageBackground(doc),
  };
}
