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
): {
  sections: ProjectedSection[];
  background?: ProjectedPageBackground;
} {
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

  const ctx: ProjectContext = {
    styles: doc.styles,
    ...(themeFonts ? { themeFonts } : {}),
    characterStyles: indexCharacterStyles(doc.styles),
    numberings: indexNumberings(doc.numbering),
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
  const sectionBlocks = (doc.sections ?? []).map((section, sIdx) => {
    if (sIdx > 0) {
      if (ctx.footnoteNumRestart === "eachSect") ctx.footnoteOrdinals.clear();
      if (ctx.endnoteNumRestart === "eachSect") ctx.endnoteOrdinals.clear();
    }
    const blocks: LayoutBlock[] = [];
    for (const child of section.children ?? []) {
      const block = projectChild(child, ctx);
      if (Array.isArray(block)) blocks.push(...block);
      else if (block) blocks.push(block);
    }
    return blocks;
  });

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

  const fnDefs = footnoteDefinitions.size > 0 ? footnoteDefinitions : undefined;
  const enDefs = endnoteDefinitions.size > 0 ? endnoteDefinitions : undefined;

  // Word: a section without a header/footer reference shows the previous
  // section's — carry the effective slots forward so projection, page insets
  // and story bands all see the linked content.
  let prevHeaders: ReturnType<typeof inheritFurnitureSlots>;
  let prevFooters: ReturnType<typeof inheritFurnitureSlots>;
  const sections: ProjectedSection[] = (doc.sections ?? []).map((section, i) => {
    const headers = inheritFurnitureSlots(section.headers, prevHeaders);
    const footers = inheritFurnitureSlots(section.footers, prevFooters);
    prevHeaders = headers;
    prevFooters = footers;
    const blocks = sectionBlocks[i] ?? [];
    // A non-final section's last paragraph carries the sectPr — Word paints
    // its mark row as "─────分节符(下一页)─────". The final section's sectPr
    // rides the body's end (no paragraph holds it) and shows no mark.
    const last = blocks[blocks.length - 1];
    if (i < (doc.sections?.length ?? 0) - 1 && last?.kind === "paragraph") {
      // The mark row names the break type (Word: "分节符(连续)") — nextPage
      // collapses to true, the painter's default label.
      const type = section.properties?.type;
      last.sectionEnd =
        type === "continuous" || type === "evenPage" || type === "oddPage" ? type : true;
    }
    return {
      blocks,
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
      footnoteDefinitions: fnDefs,
      endnoteDefinitions: enDefs,
      endnotePlacement:
        enDefs || enProps.pos
          ? ((enProps.pos === "sectEnd" ? "sectEnd" : "docEnd") as "sectEnd" | "docEnd")
          : undefined,
    };
  });
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
