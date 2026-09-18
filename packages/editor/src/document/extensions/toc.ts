import { detectHeadingLevel, type StylesOptions } from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { Node as PMNode } from "@tiptap/pm/model";

/**
 * Table of contents commands — the ribbon References tab's TOC and Update
 * Table buttons.
 *
 * The `tocField` node (engine) already round-trips a DOCX TOC's rendered
 * entries; these commands close the loop on the authoring side: scan the doc's
 * headings (the outline walk's rule), emit one entry paragraph per heading —
 * TOC1-3 style, the Word built-in per-level indent, a right dotted-leader tab,
 * and the heading's live page number from the canvas caret map — and insert or
 * refresh the `tocField` at the caret.
 *
 * Page numbers come from the host's bridge (`pageOf`): pagination lives in the
 * canvas layout, which the headless command chain can't see. The callback is a
 * command argument for the same reason — the extension is created before the
 * bridge exists, so the host wires it per dispatch (see #onCommand's local
 * branch). Inserting a TOC repaginates the doc, so the host re-runs the update
 * once the fresh layout lands; Word behaves the same (insert shows stale
 * numbers until the field updates).
 */

/** A heading's page lookup, wired by the host from the canvas caret map. */
export type PageOf = (pos: number) => number | null | undefined;

/** Word's built-in TOC styles indent 220 twips per level (TOC2 at 220, TOC3
 *  at 440). Stamped directly on the entry so the hierarchy reads even when the
 *  doc carries no TOC1-3 style definitions. */
const TOC_INDENT_TW = 220;

/** The \o switch's heading-level window ("1-3"). An absent or malformed
 *  range falls back to Word's default 1-3. */
function headingRangeOf(range: unknown): { min: number; max: number } {
  const m = /^(\d+)-(\d+)$/.exec(typeof range === "string" ? range : "");
  if (!m) return { min: 1, max: 3 };
  const min = Number(m[1]);
  const max = Number(m[2]);
  return min >= 1 && max >= min && max <= 9 ? { min, max } : { min: 1, max: 3 };
}

/** Parse \t switch custom styles mapping (e.g. "MyHeader,1,Subtitle,2" or { MyHeader: 1 }). */
export function parseCustomStyles(raw: unknown): Map<string, number> {
  const map = new Map<string, number>();
  if (!raw) return map;
  if (typeof raw === "object" && raw !== null && !Array.isArray(raw)) {
    for (const [k, v] of Object.entries(raw)) {
      const num = Number(v);
      if (num >= 1 && num <= 9) map.set(k.toLowerCase(), num);
    }
    return map;
  }
  if (typeof raw === "string") {
    const parts = raw
      .split(/[,;]/)
      .map((s) => s.trim())
      .filter(Boolean);
    for (let i = 0; i < parts.length; i += 2) {
      const name = parts[i];
      const lvl = Number(parts[i + 1]);
      if (name && lvl >= 1 && lvl <= 9) map.set(name.toLowerCase(), lvl);
    }
  }
  return map;
}

/** Max bookmark id already carried in the document passthroughs. */
function maxBookmarkIdOf(doc: PMNode): number {
  let max = -1;
  doc.descendants((child) => {
    if (child.type.name === "inlinePassthrough" || child.type.name === "passthrough") {
      try {
        const data = JSON.parse(String(child.attrs?.data ?? "{}")) as {
          bookmarkStart?: { id?: number };
          bookmarkEnd?: { id?: number };
        };
        for (const id of [data.bookmarkStart?.id, data.bookmarkEnd?.id]) {
          if (typeof id === "number" && id > max) max = id;
        }
      } catch {
        /* skip */
      }
    }
    return true;
  });
  return max;
}

/** Check whether a paragraph already has an inline _Toc... bookmarkStart. */
function existingTocBookmarkOf(node: PMNode): string | null {
  let name: string | null = null;
  node.descendants((child) => {
    if (name || child.type.name !== "inlinePassthrough") return true;
    try {
      const data = JSON.parse(String(child.attrs?.data ?? "{}")) as {
        bookmarkStart?: { name?: string };
      };
      if (data.bookmarkStart?.name && data.bookmarkStart.name.startsWith("_Toc")) {
        name = data.bookmarkStart.name;
      }
    } catch {
      /* skip */
    }
    return true;
  });
  return name;
}

export interface HeadingBookmarkInfo {
  pos: number;
  node: PMNode;
  bookmarkId: number;
  bookmarkName: string;
  needsInsert: boolean;
}

/** Find the [start, end] positions of a bookmark by name. */
export function findBookmarkRange(
  doc: PMNode,
  bookmarkName: string,
): { from: number; to: number } | null {
  let bookmarkId: number | null = null;
  let from: number | null = null;
  let to: number | null = null;

  doc.descendants((node, pos) => {
    if (node.type.name === "inlinePassthrough" || node.type.name === "passthrough") {
      try {
        const data = JSON.parse(String(node.attrs?.data ?? "{}")) as {
          bookmarkStart?: { id?: number; name?: string };
          bookmarkEnd?: { id?: number };
        };
        if (data.bookmarkStart && data.bookmarkStart.name === bookmarkName) {
          bookmarkId = data.bookmarkStart.id ?? -1;
          from = pos;
        }
        if (bookmarkId != null && data.bookmarkEnd && data.bookmarkEnd.id === bookmarkId) {
          to = pos + node.nodeSize;
        }
      } catch {
        /* skip */
      }
    }
    return true;
  });

  if (from != null && to != null) {
    return { from, to };
  }
  return null;
}

/** Entry paragraphs for the headings the TOC's level window covers. */
function buildTocEntries(
  doc: PMNode,
  pageOf?: PageOf,
  tabPositionTw = 9350,
  levels: { min: number; max: number } = { min: 1, max: 3 },
  opts: {
    leader?: string;
    showPageNumbers?: boolean;
    alignPageNumbers?: boolean;
    styles?: string;
    customStyles?: Record<string, number> | string;
    hyperlink?: boolean;
    hyperlinks?: boolean;
    bookmark?: string;
    useAppliedParagraphOutlineLevel?: boolean;
    outlineLevel?: boolean;
  } = {},
): {
  entries: { type: string; attrs?: Record<string, unknown>; content: unknown[] }[];
  headingBookmarks: HeadingBookmarkInfo[];
} {
  const styles = (doc.attrs as { styles?: StylesOptions }).styles;
  const { leader = "dot", showPageNumbers = true, alignPageNumbers = true } = opts;
  const useHyperlinks = opts.hyperlink !== false && opts.hyperlinks !== false;
  const useOutline = opts.outlineLevel ?? opts.useAppliedParagraphOutlineLevel ?? true;
  const customStyles = parseCustomStyles(opts.styles ?? opts.customStyles);
  const out: { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] = [];
  const headingBookmarks: HeadingBookmarkInfo[] = [];
  let nextBookmarkId = maxBookmarkIdOf(doc);

  let allowedRange: { from: number; to: number } | null = null;
  if (opts.bookmark) {
    allowedRange = findBookmarkRange(doc, opts.bookmark);
    if (!allowedRange) return { entries: [], headingBookmarks: [] };
  }

  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return true;
    if (allowedRange && (pos < allowedRange.from || pos > allowedRange.to)) return true;

    const customLevel = customStyles.get((node.attrs.style as string)?.toLowerCase());
    const level =
      customLevel ??
      detectHeadingLevel(
        {
          heading: (node.attrs.heading as string) || undefined,
          style: (node.attrs.style as string) || undefined,
          outlineLevel: useOutline ? (node.attrs.outlineLevel as number | undefined) : undefined,
        },
        styles,
      );
    if (level == null || level < levels.min || level > levels.max || node.textContent.length === 0)
      return true;

    const existingName = existingTocBookmarkOf(node);
    let bookmarkName: string;
    let bookmarkId = -1;
    let needsInsert = false;
    if (existingName) {
      bookmarkName = existingName;
    } else {
      nextBookmarkId++;
      bookmarkId = nextBookmarkId;
      bookmarkName = `_Toc${out.length + 1}`;
      needsInsert = true;
    }
    if (useHyperlinks) {
      headingBookmarks.push({ pos, node, bookmarkId, bookmarkName, needsInsert });
    }

    const page = showPageNumbers ? pageOf?.(pos + 1) : undefined;
    // Unaligned numbers trail the text after a space (Word's "Right align
    // page numbers" off); aligned ones ride the right leader tab.
    const numberRun =
      showPageNumbers && typeof page === "number"
        ? alignPageNumbers
          ? [{ type: "tab" }, { type: "text", text: String(page) }]
          : [{ type: "text", text: ` ${page}` }]
        : [];

    const marks = useHyperlinks
      ? [{ type: "link", attrs: { href: `#${bookmarkName}` } }]
      : undefined;

    out.push({
      type: "paragraph",
      attrs: {
        style: `TOC${level}`,
        ...(showPageNumbers && alignPageNumbers
          ? { tabStops: [{ type: "right", position: tabPositionTw, leader }] }
          : {}),
        ...(level > 1 ? { indent: { left: (level - 1) * TOC_INDENT_TW } } : {}),
      },
      content: [
        {
          type: "text",
          text: node.textContent,
          ...(marks ? { marks } : {}),
        },
        // A blank page (unmapped heading) omits the number run — an empty
        // text node is illegal in PM.
        ...numberRun,
      ],
    });
    return true;
  });
  return { entries: out, headingBookmarks };
}

/** The first tocField in the doc that is a TABLE OF CONTENTS (not the \c
 *  caption table — a figures-only document must not be rebuilt as a heading
 *  TOC), with its position (null when none). */
function findTocField(doc: PMNode): { node: PMNode; pos: number } | null {
  let found: { node: PMNode; pos: number } | null = null;
  doc.descendants((node, pos) => {
    if (found) return false;
    if (
      node.type.name === "tocField" &&
      !(node.attrs.options as { captionLabel?: string } | null)?.captionLabel
    ) {
      found = { node, pos };
      return false;
    }
    return true;
  });
  return found;
}

// ── Table of figures (the TOC field's \c switch) ──

/** The SEQ label a caption paragraph counts in — the `SEQ <label>` simple
 *  field the caption dialog inserts (attrs.data JSON inside an
 *  inlinePassthrough, invisible to textContent). Null for non-caption
 *  paragraphs and captions of another label. */
function captionLabelOf(node: PMNode): string | null {
  if (node.type.name !== "paragraph" || node.attrs.style !== "Caption") return null;
  let label: string | null = null;
  node.descendants((child) => {
    if (label || child.type.name !== "inlinePassthrough") return true;
    try {
      const data = JSON.parse(String(child.attrs.data ?? "{}")) as {
        simpleField?: { instruction?: string };
      };
      const m = /^SEQ (\S+)/.exec(data.simpleField?.instruction ?? "");
      if (m) label = m[1];
    } catch {
      /* opaque payload — not a SEQ field we can read */
    }
    return true;
  });
  return label;
}

/** Entry paragraphs for the caption paragraphs whose SEQ label matches. */
function buildTofEntries(
  doc: PMNode,
  pageOf: PageOf | undefined,
  tabPositionTw = 9350,
  label: string,
  options?: {
    leader?: string;
    showPageNumbers?: boolean;
    alignPageNumbers?: boolean;
  },
): { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] {
  const out: { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] = [];
  const leader = options?.leader ?? "dot";
  const showPageNumbers = options?.showPageNumbers ?? true;
  doc.descendants((node, pos) => {
    if (captionLabelOf(node) !== label || node.textContent.length === 0) return true;
    const page = pageOf?.(pos + 1);
    out.push({
      type: "paragraph",
      attrs: {
        style: "TOC1",
        tabStops: [{ type: "right", position: tabPositionTw, leader }],
      },
      content: [
        { type: "text", text: node.textContent },
        ...(showPageNumbers
          ? [
              { type: "tab" },
              ...(typeof page === "number" ? [{ type: "text", text: String(page) }] : []),
            ]
          : []),
      ],
    });
    return true;
  });
  return out;
}

/** The first figure-table tocField (a tocField carrying the \c captionLabel
 *  switch) — update-figures' target. */
function findTofField(doc: PMNode): { node: PMNode; pos: number } | null {
  let found: { node: PMNode; pos: number } | null = null;
  doc.descendants((node, pos) => {
    if (found) return false;
    if (
      node.type.name === "tocField" &&
      !!(node.attrs.options as { captionLabel?: string } | null)?.captionLabel
    ) {
      found = { node, pos };
      return false;
    }
    return true;
  });
  return found;
}

export const TocCommands = Extension.create({
  name: "docenTocCommands",

  addCommands() {
    return {
      // Insert a fresh TOC (Word's default: heading levels 1-3, hyperlinked)
      // at the caret. Entries build from the live headings; page numbers ride
      // the host's pageOf (null → blank until the post-insert update lands).
      // The insert options carry the custom dialog's choices (level window,
      // tab leader, page numbers); PM's nodeFromJSON (not Tiptap's
      // insertContentAt) keeps the command DOM-free — the viewless editor and
      // the headless tests have no window.
      toc:
        (
          pageOf?: PageOf,
          tabPositionTw?: number,
          insert?: {
            headingRange?: string;
            leader?: string;
            showPageNumbers?: boolean;
            alignPageNumbers?: boolean;
            styles?: string;
            customStyles?: Record<string, number> | string;
            hyperlink?: boolean;
            bookmark?: string;
            useAppliedParagraphOutlineLevel?: boolean;
          },
        ) =>
        ({ state, tr, dispatch }) => {
          const levels = headingRangeOf(insert?.headingRange);
          const { entries, headingBookmarks } = buildTocEntries(
            state.doc,
            pageOf,
            tabPositionTw,
            levels,
            insert,
          );
          if (entries.length === 0) return false;
          const node = state.schema.nodeFromJSON({
            type: "tocField",
            attrs: {
              options: {
                headingStyleRange: insert?.headingRange ?? "1-3",
                hyperlink: insert?.hyperlink !== false,
                ...(insert?.styles ? { styles: insert.styles } : {}),
                ...(insert?.bookmark ? { bookmark: insert.bookmark } : {}),
                ...(insert?.useAppliedParagraphOutlineLevel !== undefined
                  ? { useAppliedParagraphOutlineLevel: insert.useAppliedParagraphOutlineLevel }
                  : {}),
              },
            },
            content: entries,
          });
          if (!node) return false;
          if (dispatch) {
            tr.replaceSelectionWith(node).scrollIntoView();
            const seed = (data: object) => ({
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify(data) },
            });
            for (const hb of headingBookmarks.slice().reverse()) {
              if (!hb.needsInsert) continue;
              const insertPos = tr.mapping.map(hb.pos + hb.node.nodeSize - 1);
              tr.insert(insertPos, [
                state.schema.nodeFromJSON(
                  seed({ bookmarkStart: { id: hb.bookmarkId, name: hb.bookmarkName } }),
                ),
                state.schema.nodeFromJSON(seed({ bookmarkEnd: { id: hb.bookmarkId } })),
              ]);
            }
          }
          return true;
        },
      // Rebuild the first TOC's entries from the current headings — Word's F9.
      // The field switches (attrs.options) are preserved verbatim, INCLUDING
      // the \o level window: a document whose TOC covers "3-4" (skipped
      // levels) must not rebuild with the 1-3 default and lose entries.
      "update-toc":
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, tr, dispatch }) => {
          const found = findTocField(state.doc);
          if (!found) return false;
          const opts = found.node.attrs.options as {
            headingStyleRange?: string;
            styles?: string;
            customStyles?: unknown;
            leader?: string;
            showPageNumbers?: boolean;
            alignPageNumbers?: boolean;
            hyperlink?: boolean;
            bookmark?: string;
            useAppliedParagraphOutlineLevel?: boolean;
          } | null;
          const levels = headingRangeOf(opts?.headingStyleRange);
          const { entries, headingBookmarks } = buildTocEntries(
            state.doc,
            pageOf,
            tabPositionTw,
            levels,
            {
              leader: opts?.leader,
              styles: opts?.styles,
              customStyles: opts?.customStyles as Record<string, number> | string | undefined,
              showPageNumbers: opts?.showPageNumbers,
              alignPageNumbers: opts?.alignPageNumbers,
              hyperlink: opts?.hyperlink,
              bookmark: opts?.bookmark,
              useAppliedParagraphOutlineLevel: opts?.useAppliedParagraphOutlineLevel,
            },
          );
          if (entries.length === 0) return false;
          if (dispatch) {
            const seed = (data: object) => ({
              type: "inlinePassthrough",
              attrs: { data: JSON.stringify(data) },
            });
            for (const hb of headingBookmarks.slice().reverse()) {
              if (!hb.needsInsert) continue;
              const insertPos = tr.mapping.map(hb.pos + hb.node.nodeSize - 1);
              tr.insert(insertPos, [
                state.schema.nodeFromJSON(
                  seed({ bookmarkStart: { id: hb.bookmarkId, name: hb.bookmarkName } }),
                ),
                state.schema.nodeFromJSON(seed({ bookmarkEnd: { id: hb.bookmarkId } })),
              ]);
            }
            const nodes = entries.map((entry) => state.schema.nodeFromJSON(entry));
            const from = tr.mapping.map(found.pos + 1);
            const to = tr.mapping.map(found.pos + found.node.nodeSize - 1);
            tr.replaceWith(from, to, nodes);
          }
          return true;
        },
      // Word's "Update page numbers only": keep every entry's text and level,
      // re-deriving just the trailing number run from the live pagination.
      // Entries map to headings by their bookmark anchor or text (first match
      // in document order) — an entry whose heading vanished (or sits on an
      // unmapped page) keeps its number.
      "update-toc-page":
        (pageOf?: PageOf) =>
        ({ state, tr, dispatch }) => {
          const found = findTocField(state.doc);
          if (!found) return false;
          const opts = found.node.attrs.options as {
            headingStyleRange?: string;
            styles?: string;
            customStyles?: unknown;
            bookmark?: string;
            useAppliedParagraphOutlineLevel?: boolean;
          } | null;
          const levels = headingRangeOf(opts?.headingStyleRange);
          const customStyles = parseCustomStyles(opts?.styles ?? opts?.customStyles);
          const styles = (state.doc.attrs as { styles?: StylesOptions }).styles;
          let allowedRange: { from: number; to: number } | null = null;
          if (opts?.bookmark) {
            allowedRange = findBookmarkRange(state.doc, opts.bookmark);
          }
          const useOutline = opts?.useAppliedParagraphOutlineLevel ?? true;
          const bookmarkPages = new Map<string, number>();
          const textPages = new Map<string, number>();
          state.doc.descendants((node, pos) => {
            if (node.type.name !== "paragraph") return true;
            if (allowedRange && (pos < allowedRange.from || pos > allowedRange.to)) return true;
            const customLevel = customStyles.get((node.attrs.style as string)?.toLowerCase());
            const level =
              customLevel ??
              detectHeadingLevel(
                {
                  heading: (node.attrs.heading as string) || undefined,
                  style: (node.attrs.style as string) || undefined,
                  outlineLevel: useOutline
                    ? (node.attrs.outlineLevel as number | undefined)
                    : undefined,
                },
                styles,
              );
            if (
              level == null ||
              level < levels.min ||
              level > levels.max ||
              node.textContent.length === 0
            )
              return true;
            const page = pageOf?.(pos + 1);
            if (typeof page === "number") {
              if (!textPages.has(node.textContent)) textPages.set(node.textContent, page);
              const bm = existingTocBookmarkOf(node);
              if (bm && !bookmarkPages.has(bm)) bookmarkPages.set(bm, page);
            }
            return true;
          });
          if (textPages.size === 0 && bookmarkPages.size === 0) return false;
          if (!dispatch) return true;
          const docx = state.schema;
          found.node.content.forEach((entry, entryOffset) => {
            const head = entry.firstChild;
            if (entry.type.name !== "paragraph" || !head || !head.isText) return;
            const link = head.marks.find((m) => m.type.name === "link");
            const bmName =
              typeof link?.attrs?.href === "string" && link.attrs.href.startsWith("#")
                ? link.attrs.href.slice(1)
                : null;
            const page =
              (bmName ? bookmarkPages.get(bmName) : undefined) ?? textPages.get(head.textContent);
            if (page == null) return;
            const num = entry.lastChild;
            if (num && num.isText && num.textContent === String(page)) return;
            const base = found.pos + 1 + entryOffset + 1;
            if (num && num.isText) {
              tr.replaceWith(
                base + entry.content.size - num.nodeSize,
                base + entry.content.size,
                docx.text(String(page)),
              );
            } else {
              tr.insert(base + entry.content.size, docx.text(String(page)));
            }
          });
          return true;
        },
      // Word's "Remove Table of Contents": delete the heading TOC block (the
      // figures table stays — it is a different command's subject).
      "remove-toc":
        () =>
        ({ state, tr, dispatch }) => {
          const found = findTocField(state.doc);
          if (!found) return false;
          if (dispatch) tr.delete(found.pos, found.pos + found.node.nodeSize);
          return true;
        },
      // Insert a table of figures at the caret: one TOC1 entry per caption
      // counting the given SEQ label (Word's References → Insert Table of
      // Figures, the \c switch).
      "table-of-figures":
        (
          pageOf?: PageOf,
          tabPositionTw?: number,
          captionLabel = "Figure",
          insert?: {
            leader?: string;
            showPageNumbers?: boolean;
            alignPageNumbers?: boolean;
          },
        ) =>
        ({ state, dispatch }) => {
          const entries = buildTofEntries(state.doc, pageOf, tabPositionTw, captionLabel, insert);
          if (entries.length === 0) return false;
          const node = state.schema.nodeFromJSON({
            type: "tocField",
            attrs: {
              options: {
                captionLabel,
                ...(insert?.leader ? { leader: insert.leader } : {}),
                ...(insert?.showPageNumbers !== undefined
                  ? { showPageNumbers: insert.showPageNumbers }
                  : {}),
                ...(insert?.alignPageNumbers !== undefined
                  ? { alignPageNumbers: insert.alignPageNumbers }
                  : {}),
              },
            },
            content: entries,
          });
          if (!node) return false;
          if (dispatch) dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
          return true;
        },
      // Rebuild the first figure-table tocField's entries from the current
      // captions, keeping its \c label.
      "update-figures":
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, tr, dispatch }) => {
          const found = findTofField(state.doc);
          if (!found) return false;
          const opts = found.node.attrs.options as {
            captionLabel?: string;
            leader?: string;
            showPageNumbers?: boolean;
            alignPageNumbers?: boolean;
          } | null;
          const label = opts?.captionLabel ?? "Figure";
          const entries = buildTofEntries(
            state.doc,
            pageOf,
            tabPositionTw,
            label,
            opts ?? undefined,
          );
          if (entries.length === 0) return false;
          if (dispatch) {
            const nodes = entries.map((entry) => state.schema.nodeFromJSON(entry));
            tr.replaceWith(found.pos + 1, found.pos + found.node.nodeSize - 1, nodes);
          }
          return true;
        },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    docenTocCommands: {
      toc: (
        pageOf?: PageOf,
        tabPositionTw?: number,
        insert?: {
          headingRange?: string;
          leader?: string;
          showPageNumbers?: boolean;
          alignPageNumbers?: boolean;
          styles?: string;
          customStyles?: Record<string, number> | string;
          hyperlink?: boolean;
          bookmark?: string;
          useAppliedParagraphOutlineLevel?: boolean;
        },
      ) => ReturnType;
      "update-toc": (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
      "update-toc-page": (pageOf?: PageOf) => ReturnType;
      "remove-toc": () => ReturnType;
      "table-of-figures": (
        pageOf?: PageOf,
        tabPositionTw?: number,
        captionLabel?: string,
        insert?: {
          leader?: string;
          showPageNumbers?: boolean;
          alignPageNumbers?: boolean;
        },
      ) => ReturnType;
      "update-figures": (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
    };
  }
}
