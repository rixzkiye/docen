import type { StylesOptions } from "@office-open/docx";

import type { StyleEntry } from "../../style-cascade";
import type { NumberingIndex } from "./numbering";

/** Word's "Display for Review" state — how tracked changes project.
 *  "simple"/"none" show the accepted result (marks hidden), "all" shows every
 *  mark, "original" shows the pre-revision text. `authors` limits mark
 *  display to the listed reviewers — other authors' revisions render as
 *  accepted regardless of the view (undefined/empty = every author).
 *  `colors` picks the revision palette: "author" (default, Word's By-author
 *  cycle) or "changeType" (Word's legacy fixed colors — insertions and
 *  deletions share the revision red; format-change bars go neutral gray).
 *  `balloons` is Word's Show Markup → Balloons: which annotations project
 *  margin balloons ("none"/absent keeps every mark inline, the default). */
export interface MarkupDisplay {
  view: "simple" | "all" | "none" | "original";
  authors?: readonly string[];
  colors?: "author" | "changeType";
  balloons?: "all" | "comments" | "revisions" | "none";
}

/** Per-document projection context, resolved once and threaded down. */
export interface ProjectContext {
  styles: StylesOptions | undefined;
  /** The active document theme's font pair (documentExtras.settings.theme →
   *  the editor's theme gallery). Runs with NO explicit font fall back to
   *  minorFont (body) / majorFont (heading styles). This is the minimal
   *  theme-font consumer: full `w:rFonts w:*Theme` resolution against a parsed
   *  `theme1.xml` (per-run major/minor/asymmetric slots) is the follow-up. */
  themeFonts?: { majorFont?: string; minorFont?: string };
  /** id → character style (w:style type="character") — a run's w:rStyle
   *  resolves its run props here (e.g. "Hyperlink" supplies the blue underline
   *  Word paints body links with, while a TOC entry's un-styled hyperlink
   *  stays plain). */
  characterStyles: Map<string, StyleEntry>;
  numberings: NumberingIndex;
  /** Live list counters per numbering reference (level → count), advanced in
   *  document order as numbered paragraphs project. */
  listCounters: Map<string, number[]>;
  /** Comment ranges open at the current document position (w:commentRangeStart
   *  opened, w:commentRangeEnd not yet seen) — ranges span paragraphs, so the
   *  set lives across the projection walk and every text atom inside tints. */
  openComments: Set<number>;
  /** Footnote id → displayed ordinal, assigned in first-reference order
   *  (Word's numbering: the Nth distinct note referenced shows N; the same id
   *  twice shows the same number). Lives across the whole projection walk. */
  footnoteOrdinals: Map<number, number>;
  footnoteNumFmt?: string;
  footnoteNumStart?: number;
  footnoteNumRestart?: "continuous" | "eachSect" | "eachPage";
  /** Endnote id → displayed ordinal — same first-reference-order rule as the
   *  footnotes; painted as lowercase Roman (Word's endnote default numFmt). */
  endnoteOrdinals: Map<number, number>;
  endnoteNumFmt?: string;
  endnoteNumStart?: number;
  endnoteNumRestart?: "continuous" | "eachSect" | "eachPage";
  /** The document's default tab-grid pitch in px (w:defaultTabStop, twips in
   *  settings.xml resolved once here) — every paragraph's tabs fall back to
   *  it past the last explicit stop. */
  defaultTabStopPx?: number;
  /** The ordinal of the footnote/endnote currently being projected, so its
   *  footnoteRef/endnoteRef mark runs pick up the matching note number. */
  currentNoteOrdinal?: number;
  /** The tracked-changes display state (Word's Display for Review); absent =
   *  every mark projects (the round-trip-faithful default). */
  markup?: MarkupDisplay;
  /** Revision author → "By author" palette slot, assigned on first encounter
   *  in document order (Word assigns reviewer colors in the order reviewers
   *  appear). Stable for the whole projection walk, so an author keeps one
   *  color across every paragraph; cycles at the palette length. */
  revisionAuthorColors: Map<string, number>;
  /** Comment id → balloon header/body data (word/comments.xml entries the
   *  compile pass spreads into DocumentOptions). Missing ids still anchor —
   *  the card just carries no author/text. */
  commentMeta?: Map<number, { author: string; initials: string; text: string }>;
  /** w:bookmarkStart names seen at BLOCK level (between paragraphs), waiting
   *  for the next projected paragraph to anchor them. OOXML keeps bookmarks
   *  inside paragraphs, so this is the defensive branch of the model; the
   *  paragraph projection drains it into `LayoutParagraph.bookmarks`. */
  pendingBookmarkNames?: string[];
  /** Word's field-code display (Alt+F9): every field projects its instruction
   *  verbatim instead of its cached result — no dynamic page atoms, no
   *  re-hydrated result runs. */
  showFieldCodes?: boolean;
  /** Word's "Show hidden text" (Options → Display): hidden runs (w:vanish)
   *  project as displayed text with their dotted marker; off (the default)
   *  suppresses them — no advance, no ink, the source characters kept on the
   *  atom for the caret lattice. */
  showHiddenText?: boolean;
  /** Headless renderers (Node canvas / PDF) cannot rasterize SVG: an SVG
   *  picture projects its raster fallback part as the renderer src instead of
   *  the vector source (see `projectedPictureSrc`). Off (the default) keeps
   *  the vector source for renderers that can draw it. */
  rasterFallbackImages?: boolean;
  /** Document auto-hyphenation settings (w:autoHyphenation, w:doNotHyphenateCaps, etc.) */
  autoHyphenation?: boolean;
  doNotHyphenateCaps?: boolean;
  hyphenationZoneTw?: number;
  consecutiveHyphenLimit?: number;
  /** Whether auto-hyphenation is suppressed for the paragraph being projected */
  suppressAutoHyphens?: boolean;
  /** Set when the projection reads or advances order-dependent state (list
   *  counters, note ordinals, open comment ranges, the revision author
   *  palette) — the incremental projection refuses to reuse such a child
   *  (see the layout adapter's projection cache). */
  stateful?: { hit: boolean };
  /** Effective paragraph / run styling cascaded from a table cell's conditional format. */
  tableCellDefaults?: {
    paragraph?: Record<string, unknown>;
    run?: Record<string, unknown>;
  };
}

/** Mark the current child projection stateful — call wherever a projection
 *  consults or mutates one of the order-dependent {@link ProjectContext}
 *  maps. Cheap: one boolean write behind an optional chain. */
export function markStateful(ctx: ProjectContext): void {
  if (ctx.stateful !== undefined) ctx.stateful.hit = true;
}
