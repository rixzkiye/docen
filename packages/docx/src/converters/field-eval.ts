/**
 * Generated-field cache pipeline — the builder half of the field contract.
 *
 * Word stores every field twice: the instruction (`SEQ Figure \* ARABIC`,
 * `REF _Ref1`, `PAGE`) and a cached result that is painted until the next
 * update. A DOCX generated from a Tiptap model carries whatever cache the
 * model had — a from-scratch caption gets a stale or empty `SEQ` cache, a moved
 * cross-reference keeps the old text, an empty TOC emits no field result at
 * all. This module re-derives the structural caches at generation time:
 *
 *  - SEQ: per-label ordinals in document order with `\s` heading restarts,
 *    `\*` number formats and chapter prefixes (the editor's Update All Fields
 *    semantics, shared verbatim).
 *  - REF/NOTEREF: the target bookmark's inner text.
 *  - PAGE/NUMPAGES/SECTION/SECTIONPAGES/PAGEREF: resolved through the caller's
 *    page context (`FieldCacheOptions`). The builder cannot paginate on its
 *    own — Word's layout is not ours — so without a context these keep the
 *    model's cache (never a freshly invented number).
 *  - TOC: a tocField whose cached entries are missing gets real entry
 *    paragraphs built from the document's headings (or captions for a `\c`
 *    table of figures), so the generated package renders a real table of
 *    contents immediately while the field structure stays updatable.
 *
 * The pass is non-mutating: it returns the same JSON object when nothing
 * changed, else a copy with only the patched paths rebuilt.
 */

import { formatNumber } from "@docen/layout";
import type { StylesOptions } from "@office-open/docx";

import type { JSONContent } from "../core";
import { detectHeadingLevel, paragraphStyleNames } from "../extensions/paragraph";

// ── Field instruction parsing (shared with the editor's update commands) ──

/** A parsed instruction: the field name plus its positional arguments and
 *  switches (`\* MERGEFORMAT`, `\@ "yyyy/M/d"`, `\h`). Quoted strings stay
 *  whole; a switch consumes the next token as its value. */
export interface ParsedFieldInstruction {
  /** The uppercase field name (first token). */
  name: string;
  /** The instruction verbatim, trimmed. */
  raw: string;
  /** Positional arguments after the name (`REF _Ref1`, `SEQ 图` → ["_Ref1"]),
   *  switch tokens and their values excluded. */
  args: string[];
  /** Switch values keyed without the backslash: `*` → "MERGEFORMAT",
   *  `@` → "yyyy/M/d", `#` → "0", `h` → "" (a flag switch). */
  switches: Record<string, string>;
}

export function parseFieldInstruction(instruction: string): ParsedFieldInstruction {
  const raw = instruction.trim();
  const tokens: string[] = [];
  let current = "";
  let quoted = false;
  for (const ch of raw) {
    if (ch === '"') {
      quoted = !quoted;
      continue;
    }
    if (!quoted && (ch === " " || ch === "\t")) {
      if (current) {
        tokens.push(current);
        current = "";
      }
      continue;
    }
    current += ch;
  }
  if (current) tokens.push(current);
  let name = (tokens[0] ?? "").toUpperCase();
  if (name.startsWith("=") && name.length > 1) {
    const exprRest = tokens[0]!.slice(1);
    tokens.splice(0, 1, "=", exprRest);
    name = "=";
  }
  const args: string[] = [];
  const switches: Record<string, string> = {};
  for (let i = 1; i < tokens.length; i++) {
    const token = tokens[i]!;
    if (!token.startsWith("\\") || token.length < 2) {
      args.push(token);
      continue;
    }
    const key = token[1]!;
    const next = tokens[i + 1];
    if (next && !next.startsWith("\\")) {
      switches[key] = next;
      i += 1;
    } else {
      switches[key] = "";
    }
  }
  return { name, raw, args, switches };
}

/** A field atom under the caret, resolved from its passthrough branch — the
 *  instruction, the cached result, and a checkbox's checked state. */
export interface FieldRef {
  kind: "simpleField" | "complexField" | "formField";
  instruction?: string;
  result?: string;
  /** formField checkBox only. */
  checked?: boolean;
}

/** The field a passthrough branch carries, or null (not a field). Both the
 *  flat shapes office-open parses (instruction/cachedValue, instruction/result)
 *  and the checkbox's `checked` flag are read. */
export function fieldRef(branch: Record<string, unknown>): FieldRef | null {
  const simple = branch.simpleField;
  if (simple && typeof simple === "object") {
    const s = simple as { instruction?: unknown; cachedValue?: unknown };
    return {
      kind: "simpleField",
      instruction: typeof s.instruction === "string" ? s.instruction : undefined,
      result: typeof s.cachedValue === "string" ? s.cachedValue : undefined,
    };
  }
  const complex = branch.complexField;
  if (complex && typeof complex === "object") {
    const c = complex as { instruction?: unknown; result?: unknown };
    return {
      kind: "complexField",
      instruction: typeof c.instruction === "string" ? c.instruction : undefined,
      result: typeof c.result === "string" ? c.result : undefined,
    };
  }
  const form = branch.formField;
  if (form && typeof form === "object") {
    const box = (form as { checkBox?: unknown }).checkBox;
    const checked =
      box && typeof box === "object" ? (box as { checked?: unknown }).checked : undefined;
    return { kind: "formField", checked: checked === true };
  }
  return null;
}

/** SEQ `\*` switch token → the number-format token {@link formatNumber}
 *  renders, keyed by the canonical tokens the Insert Caption dialog writes. */
export const SEQ_NUMBER_FORMATS: Readonly<Record<string, string>> = {
  ARABIC: "decimal",
  ROMAN: "upperRoman",
  roman: "lowerRoman",
  ALPHABETIC: "upperLetter",
  alphabetic: "lowerLetter",
};

/** Render one sequence ordinal under its `\*` switch (Word's Insert Caption
 *  format list): Roman/ROMAN → I, roman → i; Alphabetic/ALPHABETIC → A,
 *  alphabetic → a; an unknown switch falls back to ARABIC. */
export function formatSeqNumber(switchToken: string | undefined, ordinal: number): string {
  const token = switchToken ?? "";
  switch (token.toLowerCase()) {
    case "roman":
      return formatNumber(token === "roman" ? "lowerRoman" : "upperRoman", ordinal);
    case "alphabetic":
      return formatNumber(token === "alphabetic" ? "lowerLetter" : "upperLetter", ordinal);
    default:
      return formatNumber("decimal", ordinal);
  }
}

/** The `\s` switch as a valid chapter level: an integer 1-9, or undefined. */
export function seqChapterLevel(switchValue: string | undefined): number | undefined {
  if (switchValue == null || switchValue.trim() === "") return undefined;
  const level = Number(switchValue);
  return Number.isInteger(level) && level >= 1 && level <= 9 ? level : undefined;
}

/** w:caption@w:sep token → the character between a chapter number and a SEQ
 *  number (Word's "Use separator" list). */
export const CAPTION_SEPARATOR_CHARS: Readonly<Record<string, string>> = {
  hyphen: "-",
  period: ".",
  colon: ":",
  emDash: "\u2014",
  enDash: "\u2013",
};

// ── Generated-field cache pass ──

/** Page context for the page-dependent fields. Absent = leave the model's
 *  cache untouched (the builder never invents a page number). */
export interface FieldCacheOptions {
  /** 1-based displayed page for the field atom at `index` (document order of
   *  patchable field atoms; `bookmark` set for PAGEREF targets). */
  pageOf?: (field: { index: number; instruction: string; bookmark?: string }) => number | undefined;
  /** Total pages, for NUMPAGES. */
  pageCount?: number;
  /** 1-based section number / pages in that section, for SECTION(PAGES). */
  sectionOf?: (field: { index: number; instruction: string }) => number | undefined;
  sectionPages?: number;
  /** 1-based displayed page for the generated TOC entry at `index`
   *  (document-order collection index per `kind`). Only consulted for a TOC
   *  whose cached entries are missing — a non-empty TOC is never recomputed.
   *  Absent (or undefined for one entry) = the entry keeps Word's empty
   *  page-number slot; Word/LibreOffice fill it on update. */
  tocPageOf?: (entry: { index: number; kind: "heading" | "caption" }) => number | undefined;
}

interface HeadingEntry {
  /** Collection index (document order across every heading candidate) — the
   *  key `tocPageOf` addresses this entry by. */
  index: number;
  /** Walk-assigned paragraph ordinal — `\b` scope membership compares this. */
  paragraphIndex: number;
  style: string | undefined;
  heading: string | undefined;
  outlineLevel: number | undefined;
  text: string;
}

interface CaptionEntry {
  /** Collection index across captions, parallel to {@link HeadingEntry.index}. */
  index: number;
  paragraphIndex: number;
  label: string;
  text: string;
}

/** The paragraph-ordinal span a `\b` bookmark covers (inclusive). Paragraph
 *  granularity is Word's: a heading paragraph the bookmark merely touches is
 *  inside the scope. */
interface BookmarkRange {
  start: number;
  end: number | null;
}

interface TocTarget {
  node: JSONContent;
  options: Record<string, unknown>;
}

interface WalkState {
  chapterCounts: number[];
  resetLevels: Map<string, number>;
  counts: Map<string, number>;
  headings: HeadingEntry[];
  captions: CaptionEntry[];
  bookmarks: Map<string, string>;
  openBookmarks: Map<number, { name: string; text: string }>;
  bookmarkRanges: Map<string, BookmarkRange>;
  tocTargets: TocTarget[];
  fieldIndex: number;
  /** Monotonic ordinal of the paragraph being walked (0-based). */
  paragraphIndex: number;
  /** Ordinal of the paragraph currently being walked, or -1 outside one. */
  activeParagraph: number;
}

const paragraphTextOf = (node: JSONContent): string => {
  let text = "";
  for (const child of node.content ?? []) {
    if (typeof child.text === "string") {
      text += child.text;
      continue;
    }
    if (child.type === "inlinePassthrough" || child.type === "passthrough") {
      // A field atom has no text child; its cached result is the visible
      // text (caption entries must read "Figure 1: One", not "Figure : One").
      const data = child.attrs?.data;
      if (typeof data === "string") {
        try {
          const ref = fieldRef(JSON.parse(data) as Record<string, unknown>);
          if (ref?.result != null && ref.kind !== "formField") text += ref.result;
          continue;
        } catch {
          /* opaque payload */
        }
      }
    }
    text += paragraphTextOf(child);
  }
  return text;
};

/** `\o "1-3"` heading window; malformed/absent → Word's default 1-3. */
function headingRangeOf(range: unknown): { min: number; max: number } {
  const m = /^(\d+)-(\d+)$/.exec(typeof range === "string" ? range : "");
  if (!m) return { min: 1, max: 3 };
  const min = Number(m[1]);
  const max = Number(m[2]);
  return min >= 1 && max >= min && max <= 9 ? { min, max } : { min: 1, max: 3 };
}

/** The `\t` switch's style → level pairs in every shape the model carries them:
 *  office-open's parsed `stylesWithLevels` (`StyleLevel[]`, `{styleName,level}`),
 *  the editor dialog's string ("MyHeader,1,Other,2"), or a name→level map.
 *  Ordered as written, duplicates dropped (first wins). */
export function styleLevelsOf(raw: unknown): Array<{ styleName: string; level: number }> {
  const out: Array<{ styleName: string; level: number }> = [];
  const seen = new Set<string>();
  const push = (styleName: unknown, level: unknown): void => {
    if (typeof styleName !== "string" || !styleName) return;
    const num = Number(level);
    if (!Number.isInteger(num) || num < 1 || num > 9) return;
    const key = styleName.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ styleName, level: num });
  };
  if (Array.isArray(raw)) {
    for (const entry of raw) {
      if (!entry || typeof entry !== "object") continue;
      const { styleName, name, level } = entry as {
        styleName?: unknown;
        name?: unknown;
        level?: unknown;
      };
      push(styleName ?? name, level);
    }
    return out;
  }
  if (typeof raw === "object" && raw !== null) {
    for (const [name, level] of Object.entries(raw)) push(name, level);
    return out;
  }
  if (typeof raw === "string") {
    const parts = raw
      .split(/[,;]/)
      .map((s) => s.trim())
      .filter(Boolean);
    for (let i = 0; i + 1 < parts.length; i += 2) push(parts[i], parts[i + 1]);
  }
  return out;
}

/** `\t` switch custom styles mapping ("MyHeader,1", `{ MyHeader: 1 }`, or the
 *  parsed `stylesWithLevels` array). */
export function parseCustomStyles(raw: unknown): Map<string, number> {
  const map = new Map<string, number>();
  for (const { styleName, level } of styleLevelsOf(raw)) map.set(styleName.toLowerCase(), level);
  return map;
}

/** The `\t` level for a paragraph style: matched against the style's id and
 *  every resolved name in its `basedOn` chain (a `\t` switch lists names). */
function customStyleLevel(
  customStyles: Map<string, number>,
  styles: StylesOptions | undefined,
  styleId: string | undefined,
): number | undefined {
  for (const name of paragraphStyleNames(styles, styleId)) {
    const level = customStyles.get(name);
    if (level != null) return level;
  }
  return undefined;
}

/** The caption settings' label → separator characters (w:caption@w:sep). */
function captionSeparatorsOf(root: JSONContent): Map<string, string> {
  const attrs = (root.attrs ?? {}) as { documentExtras?: Record<string, unknown> };
  const settings = attrs.documentExtras?.settings as { captions?: unknown } | undefined;
  const list = (settings?.captions as { captions?: unknown } | undefined)?.captions;
  const out = new Map<string, string>();
  if (!Array.isArray(list)) return out;
  for (const entry of list) {
    if (!entry || typeof entry !== "object") continue;
    const { name, sep } = entry as { name?: unknown; sep?: unknown };
    const char = typeof sep === "string" ? CAPTION_SEPARATOR_CHARS[sep] : undefined;
    if (typeof name === "string" && char) out.set(name, char);
  }
  return out;
}

/** The SEQ label a field atom's `data` counts in, or null (shared with the
 *  editor's caption scans, whose PM nodes carry the same serialized branch). */
export function seqLabelOfData(data: string): string | null {
  try {
    const ref = fieldRef(JSON.parse(data) as Record<string, unknown>);
    const m = /^SEQ\s+(\S+)/.exec(ref?.instruction?.trim() ?? "");
    return m ? m[1]! : null;
  } catch {
    return null;
  }
}

/** The SEQ label a paragraph's caption fields count in (first SEQ field). */
function seqLabelOf(node: JSONContent): string | null {
  for (const child of node.content ?? []) {
    if (child.type !== "inlinePassthrough") continue;
    const data = child.attrs?.data;
    if (typeof data !== "string") continue;
    const label = seqLabelOfData(data);
    if (label) return label;
  }
  return null;
}

/** Walk the document in order, computing SEQ caches inline and collecting the
 *  facts the other fields need (bookmarks, headings, TOC targets). */
function walk(
  node: JSONContent,
  state: WalkState,
  patches: Map<object, string>,
  styles: StylesOptions | undefined,
  options: FieldCacheOptions | undefined,
  captionSeparators: Map<string, string>,
  inTocEntry = false,
): void {
  if (typeof node.text === "string" && node.text !== "") {
    // Bookmark inner text accumulates across every open range.
    for (const open of state.openBookmarks.values()) open.text += node.text;
    return;
  }
  const type = node.type;
  if (type === "inlinePassthrough" || type === "passthrough") {
    const data = node.attrs?.data;
    if (typeof data === "string") {
      handlePassthrough(node, data, state, patches, styles, options, captionSeparators);
    }
    return;
  }
  const paragraphIndex = type === "paragraph" ? state.paragraphIndex++ : -1;
  if (paragraphIndex >= 0) state.activeParagraph = paragraphIndex;
  if (!inTocEntry && type === "paragraph") {
    const attrs = (node.attrs ?? {}) as {
      heading?: string;
      style?: string;
      outlineLevel?: number;
    };
    // Heading candidates: paragraphs a heading style/outline level marks, plus
    // every explicitly styled paragraph — a TOC's `\t` switch can map an
    // arbitrary style name, and the switch is only known per target (collected
    // here, filtered in `headingEntries`).
    const level = detectHeadingLevel(
      {
        heading: attrs.heading ?? undefined,
        style: attrs.style ?? undefined,
        outlineLevel: attrs.outlineLevel,
      },
      styles,
    );
    if (level != null || (typeof attrs.style === "string" && attrs.style !== "")) {
      if (level != null) {
        state.chapterCounts[level] = (state.chapterCounts[level] ?? 0) + 1;
        for (let l = level + 1; l <= 9; l++) state.chapterCounts[l] = 0;
        for (const [label, resetLevel] of state.resetLevels) {
          if (resetLevel >= level) state.counts.delete(label);
        }
      }
      state.headings.push({
        index: state.headings.length,
        paragraphIndex,
        style: attrs.style,
        heading: attrs.heading,
        outlineLevel: attrs.outlineLevel,
        text: paragraphTextOf(node),
      });
    }
    const seqLabel = attrs.style === "Caption" ? seqLabelOf(node) : null;
    if (seqLabel) {
      state.captions.push({
        index: state.captions.length,
        paragraphIndex,
        label: seqLabel,
        text: paragraphTextOf(node),
      });
    }
  }
  if (type === "tocField") {
    const opts = (node.attrs?.options as Record<string, unknown> | undefined) ?? {};
    state.tocTargets.push({ node, options: opts });
    // Entry paragraphs are cached results, not headings — but their fields
    // (PAGEREF page numbers) are still live and get patched.
    for (const child of node.content ?? []) {
      walk(child, state, patches, styles, options, captionSeparators, true);
    }
    return;
  }
  for (const child of node.content ?? []) {
    walk(child, state, patches, styles, options, captionSeparators, inTocEntry);
  }
  if (paragraphIndex >= 0) state.activeParagraph = -1;
}

/** One passthrough atom: field, bookmark marker, or neither. */
function handlePassthrough(
  node: JSONContent,
  data: string,
  state: WalkState,
  patches: Map<object, string>,
  styles: StylesOptions | undefined,
  options: FieldCacheOptions | undefined,
  captionSeparators: Map<string, string>,
): void {
  let branch: Record<string, unknown>;
  try {
    branch = JSON.parse(data) as Record<string, unknown>;
  } catch {
    return;
  }
  const start = branch.bookmarkStart as { id?: unknown; name?: unknown } | undefined;
  if (start && typeof start === "object") {
    const id = typeof start.id === "number" ? start.id : -1;
    const name = typeof start.name === "string" ? start.name : "";
    if (name) {
      state.openBookmarks.set(id, { name, text: "" });
      // First range wins: bookmark names are unique in Word. The current
      // paragraph ordinal (not the candidate count) makes a bookmark that
      // starts inside a heading's paragraph include that heading.
      if (!state.bookmarkRanges.has(name)) {
        state.bookmarkRanges.set(name, {
          start: state.activeParagraph >= 0 ? state.activeParagraph : state.paragraphIndex,
          end: null,
        });
      }
    }
    return;
  }
  const end = branch.bookmarkEnd as { id?: unknown; name?: unknown } | undefined;
  if (end && typeof end === "object") {
    const id = typeof end.id === "number" ? end.id : -1;
    const open = state.openBookmarks.get(id);
    if (open) {
      state.bookmarks.set(open.name, open.text);
      state.openBookmarks.delete(id);
      const range = state.bookmarkRanges.get(open.name);
      if (range) {
        range.end ??= state.activeParagraph >= 0 ? state.activeParagraph : state.paragraphIndex;
      }
    }
    // A block bookmark whose end arrives without its start (a split range):
    // nothing to close.
    return;
  }
  const ref = fieldRef(branch);
  if (!ref?.instruction || ref.kind === "formField") return;
  const field = parseFieldInstruction(ref.instruction);
  const index = state.fieldIndex++;
  const value = evaluateGeneratedField(field, index, state, options, captionSeparators);
  if (value == null || value === ref.result) return;
  const kind = ref.kind;
  const current = branch[kind];
  if (!current || typeof current !== "object") return;
  const next = { ...(current as Record<string, unknown>) };
  // The cached runs XML outranks the flat cache in office-open's writer; a
  // recomputed value must clear it or the stale runs keep painting.
  delete next.cachedRunsXml;
  delete next.resultRunsXml;
  if (kind === "simpleField") next.cachedValue = value;
  else next.result = value;
  patches.set(node, JSON.stringify({ ...branch, [kind]: next }));
}

/** Re-derive one field's value at generation time, or null to keep the cache. */
function evaluateGeneratedField(
  field: ParsedFieldInstruction,
  index: number,
  state: WalkState,
  options: FieldCacheOptions | undefined,
  captionSeparators: Map<string, string>,
): string | null {
  switch (field.name) {
    case "SEQ": {
      const label = field.args[0];
      if (!label) return null;
      const level = seqChapterLevel(field.switches.s);
      if (level != null) state.resetLevels.set(label, level);
      const ordinal = (state.counts.get(label) ?? 0) + 1;
      state.counts.set(label, ordinal);
      const value = formatSeqNumber(field.switches["*"], ordinal);
      const chapter = level != null ? state.chapterCounts[level] : undefined;
      if (chapter == null || chapter === 0) return value;
      return `${chapter}${captionSeparators.get(label) ?? "-"}${value}`;
    }
    case "REF":
    case "NOTEREF": {
      // `\p` (above/below) needs reading-order positions and localized terms —
      // out of scope for generation; `\n` needs paragraph numbering. Both keep
      // the model's cache rather than claiming a wrong value.
      if ("p" in field.switches || "n" in field.switches) return null;
      const bookmark = field.args[0] ? state.bookmarks.get(field.args[0]) : undefined;
      return bookmark ?? null;
    }
    case "PAGEREF": {
      const bookmark = field.args[0];
      if (!bookmark || !options?.pageOf) return null;
      const page = options.pageOf({ index, instruction: field.raw, bookmark });
      return page != null ? String(page) : null;
    }
    case "PAGE": {
      const page = options?.pageOf?.({ index, instruction: field.raw });
      return page != null ? String(page) : null;
    }
    case "NUMPAGES":
      return options?.pageCount != null ? String(options.pageCount) : null;
    case "SECTION": {
      const section = options?.sectionOf?.({ index, instruction: field.raw });
      return section != null ? String(section) : null;
    }
    case "SECTIONPAGES":
      return options?.sectionPages != null ? String(options.sectionPages) : null;
    default:
      return null;
  }
}

/** The `\b` scope of a TOC's options: the Word-faithful `entriesFromBookmark`
 *  key plus the legacy editor `bookmark` one. */
function tocBookmarkScope(options: Record<string, unknown>): string | undefined {
  for (const key of ["entriesFromBookmark", "bookmark"] as const) {
    const value = options[key];
    if (typeof value === "string" && value !== "") return value;
  }
  return undefined;
}

/** True when the entry's paragraph falls inside the bookmark's span. A range
 *  that never closed runs to the end of the document (Word's behavior). */
function inBookmarkRange(range: BookmarkRange | undefined, paragraphIndex: number): boolean {
  if (!range) return false;
  return paragraphIndex >= range.start && (range.end == null || paragraphIndex <= range.end);
}

/** Entry paragraphs for a heading TOC — Word's TOC1-3 shape: the heading text
 *  on a right-leader tab, the caller's page number when a `tocPageOf` context
 *  knows it (Word/LO fill a missing one on update). */
function headingEntries(
  headings: HeadingEntry[],
  options: Record<string, unknown>,
  styles: StylesOptions | undefined,
  bookmarkRanges: Map<string, BookmarkRange>,
  tocPageOf: FieldCacheOptions["tocPageOf"],
): JSONContent[] {
  const { min, max } = headingRangeOf(options.headingStyleRange);
  const customStyles = parseCustomStyles(
    options.styles ?? options.customStyles ?? options.stylesWithLevels,
  );
  const useOutline = options.useAppliedParagraphOutlineLevel !== false;
  const scope = tocBookmarkScope(options);
  const range = scope ? bookmarkRanges.get(scope) : undefined;
  if (scope && !range) return [];
  const entries: JSONContent[] = [];
  for (const heading of headings) {
    if (scope && !inBookmarkRange(range, heading.paragraphIndex)) continue;
    const level =
      customStyleLevel(customStyles, styles, heading.style) ??
      detectHeadingLevel(
        {
          heading: heading.heading,
          style: heading.style,
          outlineLevel: useOutline ? heading.outlineLevel : undefined,
        },
        styles,
      );
    if (level == null || level < min || level > max || heading.text === "") continue;
    entries.push(
      tocEntry(
        "TOC" + level,
        heading.text,
        level,
        options.showPageNumbers === false
          ? undefined
          : tocPageOf?.({ index: heading.index, kind: "heading" }),
        options.alignPageNumbers === false,
      ),
    );
  }
  return entries;
}

function tocEntry(
  style: string,
  text: string,
  level: number,
  page?: number,
  unaligned = false,
): JSONContent {
  const attrs: Record<string, unknown> = {
    style,
    tabStops: [{ type: "right", position: 9350, leader: "dot" }],
  };
  if (level > 1) attrs.indent = { left: (level - 1) * 220 };
  const content: JSONContent[] = [{ type: "text", text }];
  if (page != null && unaligned) content.push({ type: "text", text: ` ${page}` });
  else {
    content.push({ type: "tab" });
    if (page != null) content.push({ type: "text", text: String(page) });
  }
  return { type: "paragraph", attrs, content };
}

/** True when a tocField's cached entries are missing (the placeholder shape a
 *  from-scratch field carries): no content, or a single empty paragraph. */
function tocCacheEmpty(node: JSONContent): boolean {
  const content = node.content ?? [];
  if (content.length === 0) return true;
  if (content.length > 1) return false;
  const only = content[0];
  return (
    only?.type === "paragraph" &&
    (only.content ?? []).every((c) => !c.text && !c.type?.includes("Field"))
  );
}

/** Build the cached entries for every empty tocField from the collected
 *  headings/captions, replacing their content. */
function fillTocs(
  state: WalkState,
  styles: StylesOptions | undefined,
  tocPatches: Map<object, JSONContent[]>,
  options: FieldCacheOptions | undefined,
): void {
  for (const target of state.tocTargets) {
    if (!tocCacheEmpty(target.node)) continue;
    // A `\c` table of figures carries the parsed `captionLabelIncludingNumbers`
    // key; the editor's own command historically stamped `captionLabel`.
    const captionLabel =
      typeof target.options.captionLabelIncludingNumbers === "string"
        ? target.options.captionLabelIncludingNumbers
        : typeof target.options.captionLabel === "string"
          ? target.options.captionLabel
          : null;
    const scope = tocBookmarkScope(target.options);
    const range = scope ? state.bookmarkRanges.get(scope) : undefined;
    if (scope && !range) continue;
    const entries = captionLabel
      ? state.captions
          .filter(
            (c) =>
              c.label === captionLabel &&
              c.text !== "" &&
              (!scope || inBookmarkRange(range, c.paragraphIndex)),
          )
          .map((c) =>
            tocEntry(
              "TOC1",
              c.text,
              1,
              target.options.showPageNumbers === false
                ? undefined
                : options?.tocPageOf?.({ index: c.index, kind: "caption" }),
              target.options.alignPageNumbers === false,
            ),
          )
      : headingEntries(
          state.headings,
          target.options,
          styles,
          state.bookmarkRanges,
          options?.tocPageOf,
        );
    if (entries.length === 0) continue;
    tocPatches.set(target.node, entries);
  }
}

/** Rebuild the tree with the patched field data / TOC content, preserving
 *  object identity for untouched subtrees. */
function applyPatches(
  node: JSONContent,
  fieldPatches: Map<object, string>,
  tocPatches: Map<object, JSONContent[]>,
): JSONContent {
  const nextData = fieldPatches.get(node);
  const nextToc = tocPatches.get(node);
  let content = node.content;
  let contentChanged = false;
  if (content) {
    const mapped = content.map((child) => applyPatches(child, fieldPatches, tocPatches));
    contentChanged = mapped.some((child, i) => child !== content![i]);
    if (contentChanged) content = mapped;
  }
  if (nextData === undefined && nextToc === undefined && !contentChanged) return node;
  const next: JSONContent = { ...node };
  if (nextData !== undefined) next.attrs = { ...node.attrs, data: nextData };
  if (nextToc !== undefined) next.content = nextToc;
  else if (contentChanged) next.content = content;
  return next;
}

/**
 * Re-derive generated-field caches (SEQ/REF/NOTEREF/page fields) and empty
 * TOC entry caches over a Tiptap document. Returns the same object when
 * nothing changed, else a copy with only the touched paths rebuilt.
 */
export function fillGeneratedFields(json: JSONContent, options?: FieldCacheOptions): JSONContent {
  const styles = ((json.attrs ?? {}) as { styles?: StylesOptions }).styles;
  const state: WalkState = {
    chapterCounts: Array.from({ length: 10 }, () => 0),
    resetLevels: new Map(),
    counts: new Map(),
    headings: [],
    captions: [],
    bookmarks: new Map(),
    openBookmarks: new Map(),
    bookmarkRanges: new Map(),
    tocTargets: [],
    fieldIndex: 0,
    paragraphIndex: 0,
    activeParagraph: -1,
  };
  const fieldPatches = new Map<object, string>();
  const tocPatches = new Map<object, JSONContent[]>();
  const captionSeparators = captionSeparatorsOf(json);
  walk(json, state, fieldPatches, styles, options, captionSeparators);
  // REF targets are collected across the whole document (a field may precede
  // its bookmark); the deferred pass re-evaluates the bookmark-reading fields
  // with the complete table so forward references resolve.
  const forwardPatches = new Map<object, string>();
  evaluateDeferredRefs(json, state, forwardPatches);
  for (const [node, data] of forwardPatches) fieldPatches.set(node, data);
  fillTocs(state, styles, tocPatches, options);
  if (fieldPatches.size === 0 && tocPatches.size === 0) return json;
  return applyPatches(json, fieldPatches, tocPatches);
}

/** Second pass for the bookmark-reading fields (forward references). */
function evaluateDeferredRefs(
  node: JSONContent,
  state: WalkState,
  patches: Map<object, string>,
): void {
  const type = node.type;
  if (type === "inlinePassthrough" || type === "passthrough") {
    const data = node.attrs?.data;
    if (typeof data !== "string") return;
    let branch: Record<string, unknown>;
    try {
      branch = JSON.parse(data) as Record<string, unknown>;
    } catch {
      return;
    }
    const ref = fieldRef(branch);
    if (!ref?.instruction || ref.kind === "formField") return;
    const field = parseFieldInstruction(ref.instruction);
    if (field.name !== "REF" && field.name !== "NOTEREF") return;
    if ("p" in field.switches || "n" in field.switches) return;
    const bookmark = field.args[0] ? state.bookmarks.get(field.args[0]) : undefined;
    if (bookmark == null || bookmark === ref.result) return;
    const kind = ref.kind;
    const current = branch[kind];
    if (!current || typeof current !== "object") return;
    const next = { ...(current as Record<string, unknown>) };
    delete next.cachedRunsXml;
    delete next.resultRunsXml;
    if (kind === "simpleField") next.cachedValue = bookmark;
    else next.result = bookmark;
    patches.set(node, JSON.stringify({ ...branch, [kind]: next }));
    return;
  }
  for (const child of node.content ?? []) {
    evaluateDeferredRefs(child, state, patches);
  }
}
