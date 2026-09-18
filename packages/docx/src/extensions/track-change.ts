import type { ParagraphChild, RunOptions } from "@office-open/docx";

import type { RunPropMark } from "../converters/docx";
import { mergeTextNodes } from "../converters/styles";
import type { JSONContent } from "../core";
import { Mark } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";

/**
 * Track Changes marks (Word revision tracking).
 *
 * OOXML records inline revisions as `<w:ins>` / `<w:del>` containers wrapping
 * runs (carrying w:author / w:date / w:id metadata). office-open models these
 * as `ParagraphChild.insertion` / `.deletion` — `{ id, author, date, children }`
 * — structurally identical to `hyperlink` (an inline container with attrs +
 * child runs). docen mirrors that as two Tiptap marks applied to the contained
 * text:
 *
 *  - `insertion` — text added by a reviewer (Word renders colored + underlined)
 *  - `deletion`  — text marked for removal (Word renders colored + strikethrough;
 *    the text stays visible until the change is accepted/rejected)
 *
 * Container-level, NOT rPr-level: like `link`, these wrap child runs, so resolve
 * is declared via parseDocxInline (resolveTrackedChange) and compile via
 * compileTrackedChangeRun (compileTextRun pushes `{insertion|deletion:{...}}`).
 * They do NOT use the renderDocx/parseDocx mark hook — that is for rPr-level
 * marks like strike/bold. The attrs (id/author/date) are round-tripped via
 * resolve/compile and kept out of HTML (`rendered:false`): HTML paste loses
 * the metadata but keeps the native `<ins>`/`<del>` tag (the class is a CSS
 * hook); DOCX round-trip is byte-faithful.
 *
 * HTML tags: `<ins>`/`<del>` are HTML's native editorial-revision elements, so
 * they are used instead of a bare span — semantic, accessible, and matching
 * browser defaults (underlined / struck-through). `<ins>` has no competing
 * mark, so both the classed tag and a bare pasted `<ins>` are claimed. `<del>`
 * is also matched by the base Strike mark, so only the classed tag is claimed
 * to avoid shadowing strike on a bare `<del>`.
 *
 * Format revisions (w:rPrChange) ride the `formatChange` mark below: the old
 * run properties are stored verbatim (JSON, office-open keys) and compile back
 * into the run's `revision` option, so a Word file's format changes round-trip
 * and the editor's accept/reject can restore them. Block-level revisions
 * (paragraph splits/joins) and move tracking stay out of scope.
 */

// office-open ChangedProperties: { id:number; author:string; date:string }.
// `id` is a number on the OOXML side and is kept verbatim. All three are
// metadata — not rendered to HTML (the tag/class already identifies the mark).
const trackChangeAttrs = () => ({
  id: { default: null, rendered: false },
  author: { default: null, rendered: false },
  date: { default: null, rendered: false },
  name: { default: null, rendered: false },
  moveId: { default: null, rendered: false },
});

/** office-open's w:ins / w:del ParagraphChild branches, derived from the union
 *  (ChangedProperties & { children: TrackChangeChild[] } is not exported). */
type InsertionBranch = Extract<ParagraphChild, { insertion: unknown }>;
type DeletionBranch = Extract<ParagraphChild, { deletion: unknown }>;
type MovedFromBranch = Extract<ParagraphChild, { movedFrom: unknown }>;
type MoveFromBranch = Extract<ParagraphChild, { moveFrom: unknown }>;
type MovedToBranch = Extract<ParagraphChild, { movedTo: unknown }>;
type MoveToBranch = Extract<ParagraphChild, { moveTo: unknown }>;

/** ParagraphChild `{ insertion|deletion|movedFrom|movedTo: {...} }` → text[] carrying the mark.
 *  Mirrors the old DocxManager.resolveTrackedChange: recurse the container's
 *  runs via ctx, merge adjacent text, then stamp every text node with the
 *  revision mark alongside any existing rPr marks. Returns null for an empty
 *  container. */
function resolveTrackedChange(
  opts: Record<string, unknown> & {
    id?: unknown;
    author?: unknown;
    date?: unknown;
    name?: unknown;
    moveId?: unknown;
    children?: unknown[];
    wrap?: unknown[];
  },
  type: "insertion" | "deletion" | "moveFrom" | "moveTo",
  ctx: ResolveContext,
): JSONContent[] | null {
  // Track-change children (runs, strings, comment markers) are all valid
  // inline input — the ParagraphChild union admits every TrackChangeChild
  // shape as itself or its fallback member.
  const rawChildren = (opts.children ?? opts.wrap ?? []) as ParagraphChild[];
  const content = ctx.resolveInlineChildren(rawChildren);
  if (content.length === 0) return null;
  const merged = mergeTextNodes(content);
  const mark = {
    type,
    attrs: {
      id: opts.id ?? null,
      author: opts.author ?? null,
      date: opts.date ?? null,
      name: opts.name ?? null,
      moveId: opts.moveId ?? opts.id ?? null,
    },
  };
  for (const node of merged) {
    if (node.type === "text") {
      node.marks = [...(node.marks ?? []), mark];
    }
  }
  return merged;
}

// DOCX `<w:ins>` run → office-open ParagraphChild `{ insertion: {...} }`.
const insertionRule: ParseInlineRule<InsertionBranch> = {
  match: (child): child is InsertionBranch => "insertion" in child,
  convert: (child, ctx) => resolveTrackedChange(child.insertion as any, "insertion", ctx),
};

export const Insertion = Mark.create({
  name: "insertion",
  // Read-only render for now: a caret inside a revision range must not extend
  // the mark onto newly typed text. Re-enable inclusivity once accept/reject
  // (P1.2) makes revision editing first-class.
  inclusive: false,
  addAttributes() {
    return trackChangeAttrs();
  },
  parseHTML() {
    // Claim our <ins class="docen-insertion"> plus a bare <ins> from pasted
    // HTML — no other mark matches <ins>, so the bare tag is safe.
    return [{ tag: "ins.docen-insertion" }, { tag: "ins" }];
  },

  parseDocxInline: insertionRule,
});

// DOCX `<w:del>` run → office-open ParagraphChild `{ deletion: {...} }`.
const deletionRule: ParseInlineRule<DeletionBranch> = {
  match: (child): child is DeletionBranch => "deletion" in child,
  convert: (child, ctx) => resolveTrackedChange(child.deletion as any, "deletion", ctx),
};

export const Deletion = Mark.create({
  name: "deletion",
  inclusive: false,
  addAttributes() {
    return trackChangeAttrs();
  },
  parseHTML() {
    // Only claim <del class="docen-deletion">: the base Strike mark already
    // matches a bare <del>, so claiming all <del> would shadow strike. DOCX
    // round-trip runs through resolve/compile, not this parseHTML.
    return [{ tag: "del.docen-deletion" }];
  },

  parseDocxInline: deletionRule,
});

// DOCX `<w:moveFrom>` run → office-open ParagraphChild `{ movedFrom|moveFrom: {...} }`.
const moveFromRule: ParseInlineRule<MovedFromBranch | MoveFromBranch> = {
  match: (child): child is MovedFromBranch | MoveFromBranch =>
    "movedFrom" in child || "moveFrom" in child,
  convert: (child, ctx) =>
    resolveTrackedChange(
      ("movedFrom" in child ? child.movedFrom : (child as any).moveFrom) as any,
      "moveFrom",
      ctx,
    ),
};

export const MoveFrom = Mark.create({
  name: "moveFrom",
  inclusive: false,
  addAttributes() {
    return trackChangeAttrs();
  },
  parseHTML() {
    return [{ tag: "del.docen-move-from" }, { tag: "span.docen-move-from" }];
  },
  parseDocxInline: moveFromRule,
});

// DOCX `<w:moveTo>` run → office-open ParagraphChild `{ movedTo|moveTo: {...} }`.
const moveToRule: ParseInlineRule<MovedToBranch | MoveToBranch> = {
  match: (child): child is MovedToBranch | MoveToBranch => "movedTo" in child || "moveTo" in child,
  convert: (child, ctx) =>
    resolveTrackedChange(
      ("movedTo" in child ? child.movedTo : (child as any).moveTo) as any,
      "moveTo",
      ctx,
    ),
};

export const MoveTo = Mark.create({
  name: "moveTo",
  inclusive: false,
  addAttributes() {
    return trackChangeAttrs();
  },
  parseHTML() {
    return [{ tag: "ins.docen-move-to" }, { tag: "span.docen-move-to" }];
  },
  parseDocxInline: moveToRule,
});

/** One edit in a record's own-edit log — an exact mark-level delta. */
export interface RunFormatEdit {
  /** Monotonic order across the mark's records. Folded same-author edits can
   *  interleave with other authors, so replay sorts by this, never by record
   *  order (a folded older record may contain a later edit). */
  seq: number;
  /** The marks the edit removed (type + attrs, exactly). */
  removed: RunPropMark[];
  /** The marks the edit added. */
  added: RunPropMark[];
}

/** One tracked run format change (w:rPrChange): the revision identity plus the
 *  author's own edits. Different authors keep separate records on the same run
 *  (PM marks of one type cannot coexist on a node, so the list is the per-author
 *  storage). */
export interface RunFormatRecord {
  id: number;
  author: string;
  date: string;
  /** The record's OWN edits, oldest first — a mark-level op-log replayed over
   *  the run's base state. Empty for a DOCX-loaded record. */
  edits: RunFormatEdit[];
  /** The record's OLD run props (office-open keys) — the state before the
   *  record's first edit. Compiles to w:rPrChange for the record carrying the
   *  latest edit; also the reject target for a DOCX-loaded record (which has
   *  no base/edits to replay). */
  props: Record<string, unknown>;
}

const isRunEdit = (value: unknown): value is RunFormatEdit =>
  typeof value === "object" &&
  value !== null &&
  typeof (value as { seq?: unknown }).seq === "number" &&
  Array.isArray((value as { removed?: unknown }).removed) &&
  Array.isArray((value as { added?: unknown }).added);

const isRunRecord = (value: unknown): value is RunFormatRecord =>
  typeof value === "object" &&
  value !== null &&
  typeof (value as { id?: unknown }).id === "number" &&
  typeof (value as { author?: unknown }).author === "string" &&
  (value as { props?: unknown }).props !== null &&
  typeof (value as { props?: unknown }).props === "object";

/** Parse a formatChange mark's `records` attr; malformed entries degrade to an
 *  empty list (a corrupt record never throws the projection/compile). A record
 *  missing its op-log loads with no edits. */
export function parseFormatRecords(raw: unknown): RunFormatRecord[] {
  if (typeof raw !== "string") return [];
  try {
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(isRunRecord).map((record) => ({
      ...record,
      edits: Array.isArray(record.edits) ? record.edits.filter(isRunEdit) : [],
    }));
  } catch {
    return [];
  }
}

/** Parse a formatChange mark's `base` attr — the exact rPr mark set the mark's
 *  op-log replays over. Null when the mark was DOCX-loaded (no base snapshot). */
export function parseRunMarks(raw: unknown): RunPropMark[] | null {
  if (typeof raw !== "string") return null;
  try {
    const parsed: unknown = JSON.parse(raw);
    return Array.isArray(parsed)
      ? parsed.filter(
          (mark): mark is RunPropMark =>
            typeof mark === "object" &&
            mark !== null &&
            typeof (mark as { type?: unknown }).type === "string",
        )
      : null;
  } catch {
    return null;
  }
}

/**
 * Format-change mark (w:rPrChange) — the run-level companion of the
 * insertion/deletion containers.
 *
 * OOXML stores a run's previous properties in `<w:rPrChange>` (office-open:
 * `RunOptions.revision` = `{ id, author, date, …oldRunProps }`). The mark holds
 * the exact rPr mark set the revisions started from (`base`) plus a JSON list
 * of {@link RunFormatRecord} op-logs. Reject recomputes the run's marks by
 * replaying the surviving records' edits over `base` — order-independent, with
 * no props→marks reconstruction on the editor path. Compile emits the record
 * carrying the latest edit as `revision` (renderDocx) and resolve turns one
 * back into a DOCX-loaded record (parseDocx), so Word files round-trip.
 *
 * TextStyle declares a `revision` attr for its mirror guard but deliberately
 * skips it in render/parse: this mark owns the field (one mapping, once).
 */
export const FormatChange = Mark.create({
  name: "formatChange",
  // Read-only render (no HTML route): a caret inside a format revision must
  // not extend the mark onto newly typed text.
  inclusive: false,
  addAttributes() {
    return {
      /** JSON `RunPropMark[]` — the exact rPr marks before the first recorded
       *  edit. Null on a DOCX-loaded mark (no mark snapshot exists). */
      base: { default: null, rendered: false },
      /** JSON `RunFormatRecord[]` — one entry per author's tracked change on
       *  the marked text; replay orders the edits by their `seq`. */
      records: { default: null, rendered: false },
    };
  },
  parseDocx(opts: RunOptions) {
    const rev = opts.revision;
    if (!rev || typeof rev !== "object") return null;
    const { id, author, date, ...props } = rev;
    const record: RunFormatRecord = {
      id: typeof id === "number" ? id : 0,
      author: typeof author === "string" ? author : "",
      date: typeof date === "string" ? date : "",
      edits: [],
      props,
    };
    // A loaded record carries no mark snapshot: reject reconstructs its old
    // props (the only before-state OOXML keeps).
    return { base: null, records: JSON.stringify([record]) };
  },
  renderDocx(attrs: Record<string, unknown>) {
    const records = parseFormatRecords(attrs.records);
    const newest = latestEditedRecord(records);
    if (!newest) return {};
    return {
      revision: {
        id: newest.id,
        author: newest.author,
        date: newest.date,
        ...newest.props,
      },
    };
  },
});

/** The record carrying the mark's latest edit (a folded older record may hold
 *  a later edit than the list tail). Loaded records without edits tie at -1;
 *  the last such record wins. */
function latestEditedRecord(records: readonly RunFormatRecord[]): RunFormatRecord | undefined {
  let newest: RunFormatRecord | undefined;
  let best = -1;
  for (const record of records) {
    const seq = record.edits.reduce((max, edit) => Math.max(max, edit.seq), -1);
    if (seq >= best) {
      best = seq;
      newest = record;
    }
  }
  return newest;
}
