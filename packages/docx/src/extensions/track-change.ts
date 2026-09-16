import type { ParagraphChild, RunOptions } from "@office-open/docx";

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
});

/** office-open's w:ins / w:del ParagraphChild branches, derived from the union
 *  (ChangedProperties & { children: TrackChangeChild[] } is not exported). */
type InsertionBranch = Extract<ParagraphChild, { insertion: unknown }>;
type DeletionBranch = Extract<ParagraphChild, { deletion: unknown }>;

/** ParagraphChild `{ insertion|deletion: {...} }` → text[] carrying the mark.
 *  Mirrors the old DocxManager.resolveTrackedChange: recurse the container's
 *  runs via ctx, merge adjacent text, then stamp every text node with the
 *  revision mark alongside any existing rPr marks. Returns null for an empty
 *  container. */
function resolveTrackedChange(
  opts: InsertionBranch["insertion"] | DeletionBranch["deletion"],
  type: "insertion" | "deletion",
  ctx: ResolveContext,
): JSONContent[] | null {
  // Track-change children (runs, strings, comment markers) are all valid
  // inline input — the ParagraphChild union admits every TrackChangeChild
  // shape as itself or its fallback member.
  const content = ctx.resolveInlineChildren(opts.children ?? []);
  if (content.length === 0) return null;
  const merged = mergeTextNodes(content);
  const mark = {
    type,
    attrs: {
      id: opts.id ?? null,
      author: opts.author ?? null,
      date: opts.date ?? null,
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
  convert: (child, ctx) => resolveTrackedChange(child.insertion, "insertion", ctx),
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
  convert: (child, ctx) => resolveTrackedChange(child.deletion, "deletion", ctx),
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

/** Parse the stored old-rPr JSON into an options object; malformed entries
 *  degrade to no props (the revision metadata still round-trips). */
function parseRunProps(raw: unknown): Record<string, unknown> {
  if (typeof raw !== "string") return {};
  try {
    const parsed: unknown = JSON.parse(raw);
    return parsed !== null && typeof parsed === "object" && !Array.isArray(parsed)
      ? (parsed as Record<string, unknown>)
      : {};
  } catch {
    return {};
  }
}

/**
 * Format-change mark (w:rPrChange) — the run-level companion of the
 * insertion/deletion containers.
 *
 * OOXML stores a run's previous properties in `<w:rPrChange>` (office-open:
 * `RunOptions.revision` = `{ id, author, date, …oldRunProps }`). The mark holds
 * the revision metadata plus the old props verbatim as a JSON string; compile
 * merges them back into `revision` (renderDocx) and resolve extracts them from
 * it (parseDocx), so a Word file's format changes survive the JSON round-trip.
 * The current run properties stay on their own rPr marks — accept keeps them,
 * reject restores the old ones from `props`.
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
      id: { default: null, rendered: false },
      author: { default: null, rendered: false },
      date: { default: null, rendered: false },
      /** The old run props (office-open rPr keys) as JSON — the rPrChange body
       *  minus id/author/date. `null` for a bare revision with no old props. */
      props: { default: null, rendered: false },
    };
  },
  parseDocx(opts: RunOptions) {
    const rev = opts.revision;
    if (!rev || typeof rev !== "object") return null;
    const { id, author, date, ...props } = rev;
    return {
      id: id ?? null,
      author: author ?? null,
      date: date ?? null,
      props: JSON.stringify(props),
    };
  },
  renderDocx(attrs: Record<string, unknown>) {
    return {
      revision: {
        id: typeof attrs.id === "number" ? attrs.id : 0,
        author: typeof attrs.author === "string" ? attrs.author : "",
        date: typeof attrs.date === "string" ? attrs.date : "",
        ...parseRunProps(attrs.props),
      },
    };
  },
});
