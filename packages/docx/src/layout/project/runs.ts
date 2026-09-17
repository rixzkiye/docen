// Inline run projection: text runs (rPr resolved over the paragraph
// default), hard breaks, tabs, pictures, footnote/endnote references, fields
// (PAGE/NUMPAGES dynamic, complex-field results re-hydrated), and the
// container children (hyperlink / tracked insertion / deletion).

import {
  emuToPx,
  hyphenateText,
  layoutMathTree,
  ptToPx,
  twipToPx,
  type LayoutBalloonAnchor,
  type LayoutCombine,
  type LayoutInline,
  type LayoutTextStyle,
} from "@docen/layout";

import { mergeStyleChain } from "../../style-cascade";
import type { MarkupDisplay, ProjectContext } from "./context";
import { markStateful } from "./context";
import { cropOf, outlineOf, pictureAdjustOf } from "./drawing";
import { isRecord, measureEmu, num, str, unescapeXml, type Rec } from "./guards";
import { metafileMembers, pictureSrc } from "./media";
import { formatNumber } from "./numbering";
import { fontAttr, normalizeScalePct, toFamily, runStyleOf } from "./styles";

/** Word's "By author" revision palette — slot 0 is the red Word's first
 *  reviewer gets (the canvas default before per-author colors existed, so
 *  single-author documents render exactly as before). Word assigns colors in
 *  reviewer order; the projection does the same on first encounter. Colors
 *  are OOXML bare hex (no #). */
export const REVISION_AUTHOR_COLORS = [
  "FF0000",
  "2E74B5",
  "2F7D32",
  "9C27B0",
  "E36C0A",
  "00838F",
  "C2185B",
  "6D4C41",
] as const;

/** "By change type" mode: insertions and deletions share Word's legacy
 *  revision red; format-change bars stay neutral. */
const REVISION_CHANGE_TYPE_COLOR = "FF0000";
const FORMAT_CHANGE_TYPE_COLOR = "808080";

/** The author's stable "By author" palette slot — assigned on first encounter
 *  (Word colors reviewers, not views), shared by revision marks, format bars
 *  and comment balloons. */
export function authorColorOf(ctx: ProjectContext, author: string): string {
  markStateful(ctx);
  let slot = ctx.revisionAuthorColors.get(author);
  if (slot == null) {
    slot = ctx.revisionAuthorColors.size;
    ctx.revisionAuthorColors.set(author, slot);
  }
  return REVISION_AUTHOR_COLORS[slot % REVISION_AUTHOR_COLORS.length]!;
}

/** The revision's display color: the author's stable palette slot in
 *  "By author" mode (the default), the fixed revision red in "By change type"
 *  mode. `mark` is the w:ins/w:del/rPrChange record carrying w:author. */
function revisionColor(ctx: ProjectContext, mark: Rec): string {
  if (ctx.markup?.colors === "changeType") return REVISION_CHANGE_TYPE_COLOR;
  return authorColorOf(ctx, str(mark.author) ?? "");
}

/** Word's format-change bar color: the author's palette slot, or the neutral
 *  change-type gray when the display is "By change type". */
function formatChangeColor(ctx: ProjectContext, mark: Rec): string {
  return ctx.markup?.colors === "changeType" ? FORMAT_CHANGE_TYPE_COLOR : revisionColor(ctx, mark);
}

/** Which balloon kinds the display state asks for (comments / revisions /
 *  deletions in the full view). Absent markup or "none" = all inline. */
export function balloonKinds(ctx: ProjectContext): {
  comments: boolean;
  revisions: boolean;
  deletions: boolean;
} {
  const mode = ctx.markup?.balloons;
  const view = ctx.markup?.view ?? "all";
  const shows = view !== "none" && view !== "original";
  const revisions = shows && (mode === "all" || mode === "revisions");
  return {
    comments: shows && (mode === "all" || mode === "comments"),
    revisions,
    deletions: revisions && view === "all",
  };
}

/** Flatten a container's child runs to plain text (a deletion balloon's
 *  excerpt) — strings, run shapes and nested paragraph shapes all reduce. */
function containerText(children: readonly unknown[]): string {
  let text = "";
  for (const child of children) {
    if (typeof child === "string") {
      text += child;
      continue;
    }
    if (!isRecord(child)) continue;
    if (typeof child.text === "string") {
      text += child.text;
      continue;
    }
    if (Array.isArray(child.children)) text += containerText(child.children);
  }
  return text;
}

/** One revision mark's effective view under the display state: a mark whose
 *  author sits outside the filter renders as accepted ("none") no matter the
 *  view; simple/no-markup share the accepted rendering. */
function effectiveView(markup: MarkupDisplay | undefined, mark: Rec): "all" | "none" | "original" {
  const view = markup?.view ?? "all";
  const authors = markup?.authors;
  if (view === "simple" || view === "none") return "none";
  if (!authors || authors.length === 0) return view;
  const author = str(mark.author) ?? "";
  return author !== "" && authors.includes(author) ? view : "none";
}

/** The paint metadata for a run carrying w:rPrChange (formatChange mark in the
 *  runtime model), or undefined when the run has no shown format change. The
 *  author's palette slot is assigned even when the mark is filtered out, so
 *  toggling Specific People never recolors the remaining authors. */
export function formatIndicatorOf(
  ctx: ProjectContext,
  mark: Rec,
): { formatChange: { color: string } } | Record<string, never> {
  const color = formatChangeColor(ctx, mark);
  return effectiveView(ctx.markup, mark) === "all" ? { formatChange: { color } } : {};
}

/** Memo for complex-field result XML tokenization, keyed by the verbatim
 *  string. The projection re-runs per transaction and a TOC-dense document
 *  re-tokenizes the same cached results every keystroke; the XML only changes
 *  when the field updates. Cleared wholesale at the cap — distinct result
 *  strings per document are few. */
const FIELD_TOKEN_CACHE_CAP = 64;
const fieldTokenCache = new Map<string, string[]>();

function fieldResultTokens(xml: string): string[] {
  const cached = fieldTokenCache.get(xml);
  if (cached) return cached;
  if (fieldTokenCache.size >= FIELD_TOKEN_CACHE_CAP) fieldTokenCache.clear();
  const tokens = xml.match(/<w:(fldChar|instrText|tab|t)\b[^>]*(?:\/>|>([\s\S]*?)<\/w:\1>)/g) ?? [];
  fieldTokenCache.set(xml, tokens);
  return tokens;
}

/** A footnote/endnote reference's note id — the bare number form (`{
 *  footnoteReference: 1 }` / `{ endnoteReference: 1 }`) or the option object
 *  form (`{ id }`); anything else is not one. */
function noteRefId(child: Rec, key: "footnoteReference" | "endnoteReference"): number | undefined {
  const ref = child[key];
  if (typeof ref === "number") return ref;
  if (isRecord(ref)) return num(ref.id);
  return undefined;
}

/** The displayed ordinal for a note id — assign the next number on first
 *  reference, reuse it afterward (the Nth distinct note referenced shows N). */
function noteOrdinal(ordinals: Map<number, number>, id: number, start = 1): number {
  let ordinal = ordinals.get(id);
  if (ordinal == null) {
    ordinal = ordinals.size + start;
    ordinals.set(id, ordinal);
  }
  return ordinal;
}

/** A short structural label for a formula placeholder box — the outermost
 *  structure's glyph approximation (√□, □/□, ∑□), not a rendering of it. */
function mathLabelOf(math: Rec): string {
  const first = Array.isArray(math.children)
    ? math.children.find((c) => isRecord(c) && !("text" in c))
    : undefined;
  if (!isRecord(first)) return "fx";
  if ("fraction" in first) return "□/□";
  if ("superScript" in first) return "□^□";
  if ("subScript" in first) return "□_□";
  if ("radical" in first) return "√□";
  if ("sum" in first) return "∑□";
  if ("integral" in first) return "∫□";
  return "fx";
}

const COMBINE_BRACKETS: Record<string, "round" | "square" | "angle" | "curly"> = {
  round: "round",
  square: "square",
  angle: "angle",
  curly: "curly",
};

/** Two-lines-in-one metadata (w:eastAsianLayout @w:combine) for a run's text:
 *  the two lines the run packs into and its optional bracket pair. Word's
 *  dialog lets spaces mark the split — folded away here for an even split of
 *  the remaining characters. */
function combineOf(rPr: Rec, text: string): LayoutCombine | undefined {
  const layout = isRecord(rPr.eastAsianLayout) ? rPr.eastAsianLayout : {};
  const on = layout.combine === true || layout.combine === "1" || layout.combine === 1;
  if (!on) return undefined;
  // Word's dialog marks the split with a space — honor it; an unsplit text
  // folds to an even split of the whitespace-free run.
  const at = text.indexOf(" ");
  const [first, second] =
    at >= 0 ? [text.slice(0, at), text.slice(at + 1)] : evenSplit(text.replace(/\s+/g, ""));
  if (!first && !second) return undefined;
  const bracket = str(layout.combineBrackets);
  return {
    first,
    second,
    ...(bracket && COMBINE_BRACKETS[bracket] ? { bracket: COMBINE_BRACKETS[bracket] } : {}),
  };
}

/** Even two-lines-in-one split — the shorter second line takes the tail. */
function evenSplit(text: string): [string, string] {
  const half = Math.ceil(text.length / 2);
  return [text.slice(0, half), text.slice(half)];
}

/** Inline content: text runs (rPr resolved over the paragraph default), hard
 *  breaks, pictures (paragraph-child or run-child slot), and the container
 *  children (hyperlink / insertion / deletion — their runs project with the
 *  Word display preset above) as atoms. The members arrive as unknown — the
 *  ParagraphChild union is wide and its runtime shapes are looser still
 *  (compile pushes `{text, …rPr}` run forms), so each leg is validated rather
 *  than trusted.
 *  The remaining inline atoms project as their own kinds: tab/break as zero
 *  width, charts through the picture member-replay channel, fields dynamic
 *  or re-hydrated from cached results, math as a labeled placeholder slot
 *  (no OMML layout engine yet). */
export function projectRuns(
  runs: readonly unknown[],
  chainRPr: Rec,
  docRPr: Rec,
  defRun: LayoutTextStyle,
  ctx: ProjectContext,
  anchors?: LayoutBalloonAnchor[],
): LayoutInline[] {
  const { openComments } = ctx;
  const out: LayoutInline[] = [];
  // ── Margin balloons (Word's Show Markup → Balloons) ──────────────────────
  // Anchors ride the paragraph's atoms; the flow resolves them to lines and
  // packs the page margin. Emission is display-only: measurement and wrapping
  // never read the list, so body geometry is untouched.
  const kinds = balloonKinds(ctx);
  const wantsComments = anchors != null && kinds.comments;
  const wantsRevisions = anchors != null && kinds.revisions;
  const wantsDeletions = anchors != null && kinds.deletions;
  /** Comment ids opened in this paragraph but not yet anchored. */
  const pendingComments: number[] = [];
  const pushCommentAnchors = (inlineIndex: number): void => {
    if (!anchors || pendingComments.length === 0) return;
    for (const id of pendingComments.splice(0)) {
      const meta = ctx.commentMeta?.get(id);
      if (ctx.markup?.authors && ctx.markup.authors.length > 0) {
        const author = meta?.author ?? "";
        if (author === "" || !ctx.markup.authors.includes(author)) continue;
      }
      anchors.push({
        id,
        kind: "comment",
        color: authorColorOf(ctx, meta?.author ?? ""),
        label: meta?.initials || meta?.author || "",
        text: meta?.text,
        inlineIndex,
      });
    }
  };
  const revisionBalloon = (
    revision: Rec,
    inlineIndex: number,
    color: string,
    text?: string,
  ): void => {
    if (!anchors || !wantsRevisions) return;
    if (effectiveView(ctx.markup, revision) !== "all") return;
    anchors.push({
      id: num(revision.id) ?? 0,
      kind: "revision",
      color,
      label: str(revision.author) ?? "",
      ...(text ? { text } : {}),
      inlineIndex,
    });
  };
  const textStyleOf = (rPr: Rec): LayoutTextStyle => {
    // A run's character style (w:rStyle, e.g. a body link's "Hyperlink") slots
    // between the paragraph-style chain and direct formatting: its props beat
    // chainRPr/docRPr (they land in `own`), and an explicit rPr field beats it
    // (the spread keeps rPr's keys on top). Word paints w:hyperlink content
    // solely through this style — the container itself carries no look.
    const styleId = str(rPr.style);
    const charRun = styleId ? mergeStyleChain(ctx.characterStyles, styleId).run : undefined;
    const own = runStyleOf(charRun ? { ...charRun, ...rPr } : rPr);
    const isNoteRun = ctx.currentNoteOrdinal != null;
    const chainSizeRaw = num(chainRPr.size);
    const chainSize = isNoteRun && chainSizeRaw === 20 ? 10 : chainSizeRaw;
    const effectiveSizePt = own.sizePt ?? chainSize ?? num(docRPr.size) ?? (isNoteRun ? 10 : 12);
    // w:caps wins over w:smallCaps when both apply; an explicit false on
    // either cancels the inherited caps token (direct beats the chain).
    const caps: LayoutTextStyle["caps"] =
      own.allCaps === true
        ? "all"
        : own.smallCaps === true
          ? "small"
          : own.allCaps === false || own.smallCaps === false
            ? undefined
            : defRun.caps;
    const scaleRaw = own.scalePct ?? defRun.scalePct;
    // ST_TextScale: only 1-600 is meaningful; 100 is the identity (and an
    // explicit 100 must cancel an inherited scale, so it maps to undefined).
    const scalePct = normalizeScalePct(scaleRaw);
    // Remaining character effects cascade field-by-field (direct/character
    // style beats the paragraph default): an explicit value wins, absent
    // falls through to defRun.
    const baselineShiftPx =
      own.positionPt != null ? -ptToPx(own.positionPt) : defRun.baselineShiftPx;
    const hidden = own.vanish ?? defRun.hidden;
    const kernPt = own.kernPt ?? defRun.kernPt;
    const border = own.border ?? defRun.border;
    const emphasisMark = own.emphasisMark ?? defRun.emphasisMark;
    const outline = own.outline ?? defRun.outline;
    const shadow = own.shadow ?? defRun.shadow;
    const emboss = own.emboss ?? defRun.emboss;
    const imprint = own.imprint ?? defRun.imprint;
    const glow = own.glow ?? defRun.glow;
    const reflection = own.reflection ?? defRun.reflection;
    return {
      family: toFamily(own.font, fontAttr(chainRPr.font) ?? fontAttr(docRPr.font)) ?? defRun.family,
      sizePx: ptToPx(effectiveSizePt),
      bold: own.bold ?? defRun.bold,
      italic: own.italic ?? defRun.italic,
      color: own.color ?? defRun.color,
      highlight: own.highlight ?? str(chainRPr.highlight) ?? str(docRPr.highlight),
      shadingFill:
        own.shadingFill ??
        str((isRecord(chainRPr.shading) ? chainRPr.shading : undefined)?.fill) ??
        str((isRecord(docRPr.shading) ? docRPr.shading : undefined)?.fill) ??
        defRun.shadingFill,
      underline: own.underline ?? defRun.underline,
      underlineStyle: own.underlineStyle ?? defRun.underlineStyle,
      underlineColor: own.underlineColor ?? defRun.underlineColor,
      strikethrough: own.strikethrough ?? defRun.strikethrough,
      letterSpacingPx:
        own.characterSpacingTw != null ? twipToPx(own.characterSpacingTw) : defRun.letterSpacingPx,
      verticalAlign: own.verticalAlign ?? defRun.verticalAlign,
      direction: own.rtl === true ? "rtl" : own.rtl === false ? "ltr" : defRun.direction,
      caps,
      ...(scalePct != null ? { scalePct } : {}),
      ...(baselineShiftPx != null ? { baselineShiftPx } : {}),
      ...(hidden ? { hidden: true } : {}),
      ...(kernPt != null ? { kernPt } : {}),
      ...(border ? { border } : {}),
      ...(emphasisMark ? { emphasisMark } : {}),
      ...(outline ? { outline } : {}),
      ...(shadow ? { shadow } : {}),
      ...(emboss ? { emboss: true } : {}),
      ...(imprint ? { imprint: true } : {}),
      ...(glow ? { glow } : {}),
      ...(reflection ? { reflection } : {}),
      ...((own.ligatures ?? defRun.ligatures)
        ? { ligatures: own.ligatures ?? defRun.ligatures }
        : {}),
      ...((own.numForm ?? defRun.numForm) ? { numForm: own.numForm ?? defRun.numForm } : {}),
      ...((own.numSpacing ?? defRun.numSpacing)
        ? { numSpacing: own.numSpacing ?? defRun.numSpacing }
        : {}),
      ...((own.stylisticSet ?? defRun.stylisticSet)
        ? { stylisticSet: own.stylisticSet ?? defRun.stylisticSet }
        : {}),
      ...((own.fontFeatures ?? defRun.fontFeatures)
        ? { fontFeatures: own.fontFeatures ?? defRun.fontFeatures }
        : {}),
      ...((own.fontVariations ?? defRun.fontVariations)
        ? { fontVariations: own.fontVariations ?? defRun.fontVariations }
        : {}),
      ...((own.fontWeight ?? defRun.fontWeight)
        ? { fontWeight: own.fontWeight ?? defRun.fontWeight }
        : {}),
    };
  };
  const pushText = (text: string, rPr: Rec): void => {
    if (!text) return;
    const runLang = str(rPr.language) || "en";
    const runText =
      ctx.autoHyphenation && !ctx.suppressAutoHyphens
        ? hyphenateText(text, runLang, {
            doNotHyphenateCaps: ctx.doNotHyphenateCaps,
            hyphenationZoneTw: ctx.hyphenationZoneTw,
            consecutiveHyphenLimit: ctx.consecutiveHyphenLimit,
          })
        : text;
    // Read per atom — openComments mutates as the walk opens/closes ranges.
    if (openComments && openComments.size > 0) markStateful(ctx);
    const commentIds =
      openComments && openComments.size > 0 ? [...openComments].sort((a, b) => a - b) : undefined;
    // Two-lines-in-one (双行合一 / 合并字符): the run packs into two
    // half-size lines; the dialog's spaces mark the split in Word, here
    // folded away for an even split.
    const combine = combineOf(rPr, runText);
    // A run whose rPrChange is shown carries the format-change paint marker
    // (Word's change bar rides the run's line — the painter draws it) and,
    // when balloons are on, a margin anchor.
    const revision = isRecord(rPr.revision) ? rPr.revision : undefined;
    const indicator = revision ? formatIndicatorOf(ctx, revision) : {};
    const link = isRecord(rPr.link)
      ? {
          url: typeof rPr.link.url === "string" ? rPr.link.url : undefined,
          anchor: typeof rPr.link.anchor === "string" ? rPr.link.anchor : undefined,
          tooltip: typeof rPr.link.tooltip === "string" ? rPr.link.tooltip : undefined,
        }
      : undefined;
    out.push({
      kind: "text",
      text: runText,
      style: textStyleOf(rPr),
      commentIds,
      ...(combine ? { combine } : {}),
      ...(link ? { link } : {}),
      ...indicator,
    });
    const inlineIndex = out.length - 1;
    pushCommentAnchors(inlineIndex);
    const change = (indicator as { formatChange?: { color: string } }).formatChange;
    if (revision && change) revisionBalloon(revision, inlineIndex, change.color, text);
  };
  /** A field (w:fldSimple / complexField): PAGE/NUMPAGES become dynamic atoms
   *  (the painter resolves the number per page — `text` is a measuring
   *  placeholder); anything else renders its cached result. Every field atom
   *  carries its descriptor (instruction verbatim + cached result) so the
   *  render pass can re-resolve it and Alt+F9 can show the code. A structured
   *  result (resultRunsXml — present when the result runs hold anything but
   *  plain text, e.g. a TOC hyperlink's tab + nested PAGEREF) is re-hydrated
   *  item by item; only then does the flat `result` string stand in. */
  const pushField = (field: Rec, rPr: Rec): void => {
    const raw = typeof field.instruction === "string" ? field.instruction.trim() : "";
    const instr = raw.toUpperCase();
    const cached =
      typeof field.result === "string" ? field.result : (field.cachedValue as string | undefined);
    const style = textStyleOf(rPr);
    const descriptor = raw
      ? { instruction: raw, ...(typeof cached === "string" ? { result: cached } : {}) }
      : {};
    // A run carrying rPrChange keeps the format-change marker on every atom
    // it emits (field results included).
    const revision = isRecord(rPr.revision) ? rPr.revision : undefined;
    const format = revision ? formatIndicatorOf(ctx, revision) : {};
    if (ctx.showFieldCodes && raw) {
      // Alt+F9: the instruction verbatim in place of every result.
      out.push({ kind: "text", text: raw, style, ...descriptor, ...format });
    } else if (
      instr.startsWith("PAGE") &&
      !instr.startsWith("PAGES") &&
      !instr.startsWith("PAGEREF")
    ) {
      out.push({ kind: "text", text: "0", style, field: "page", ...descriptor, ...format });
    } else if (instr.startsWith("NUMPAGES")) {
      out.push({ kind: "text", text: "0", style, field: "numPages", ...descriptor, ...format });
    } else if (typeof field.resultRunsXml === "string") {
      // A structured result (TOC entry, nested field) re-hydrates item by
      // item and outranks the flat-result placeholder branches.
      pushFieldResultRuns(field.resultRunsXml, rPr);
    } else if (/^SECTION(PAGES)?\b/.test(instr)) {
      // The render pass re-resolves the section numbering live; keep a
      // measuring placeholder when the document cached no result so the atom
      // stays on a line the resolve walk visits.
      out.push({ kind: "text", text: cached || "0", style, ...descriptor, ...format });
    } else if (raw) {
      // Emit the atom even for an empty cache: the render pass and the update
      // commands resolve the field from its descriptor, and Word keeps the
      // (empty) field mark too.
      out.push({ kind: "text", text: cached ?? "", style, ...descriptor, ...format });
    } else if (cached) {
      pushText(cached, rPr);
    }
  };
  /** Walk a complex field's verbatim result-run XML: text runs become text,
   *  w:tab atoms become tab jumps, and a nested field (a TOC entry's PAGEREF)
   *  contributes its separated result — instruction runs are skipped. */
  const pushFieldResultRuns = (xml: string, rPr: Rec): void => {
    // Stack of nested fields, each "instr" until its separate, then "result".
    const stack: ("instr" | "result")[] = [];
    const tokens = fieldResultTokens(xml);
    for (const tk of tokens) {
      if (tk.startsWith("<w:fldChar")) {
        const type = /w:fldCharType="(\w+)"/.exec(tk)?.[1];
        if (type === "begin") stack.push("instr");
        else if (type === "separate" && stack.length > 0) stack[stack.length - 1] = "result";
        else if (type === "end") stack.pop();
      } else if (tk.startsWith("<w:instrText")) {
        continue;
      } else if (stack[stack.length - 1] === "instr") {
        continue;
      } else if (tk.startsWith("<w:tab")) {
        out.push({ kind: "tab" });
      } else {
        const text = unescapeXml(/>([\s\S]*?)<\/w:t>$/.exec(tk)?.[1] ?? "");
        pushText(text, rPr);
      }
    }
  };
  /** Flatten a w:rt / w:rubyBase content's runs to plain text — the
   *  annotation has no formatting of its own beyond its font size. */
  const rubyContentText = (content: unknown): string => {
    if (!isRecord(content) || !Array.isArray(content.children)) return "";
    let text = "";
    for (const c of content.children) {
      if (typeof c === "string") text += c;
      else if (isRecord(c) && typeof c.text === "string") text += c.text;
    }
    return text;
  };
  const pushRuby = (ruby: Rec, rPr: Rec): void => {
    const baseText = rubyContentText(ruby.base);
    if (!baseText) return;
    const props = isRecord(ruby.properties) ? ruby.properties : {};
    const style = textStyleOf(rPr);
    // Word's default annotation is half the base size (w:hps absent).
    const sizePt = num(props.fontSize);
    out.push({
      kind: "text",
      text: baseText,
      style,
      ruby: {
        text: rubyContentText(ruby.text),
        alignment: str(props.alignment),
        fontSizePx: sizePt != null ? ptToPx(sizePt) : style.sizePx / 2,
      },
    });
  };
  const pushPicture = (pic: Rec): void => {
    // A floating picture is an anchored drawing (projectDrawings), not an
    // inline atom — projecting it here too would double-render it.
    if (isRecord(pic.floating)) return;
    const tr = isRecord(pic.transformation) ? pic.transformation : {};
    const w = measureEmu(tr.width);
    const h = measureEmu(tr.height);
    if (w != null && h != null) {
      const widthPx = emuToPx(w);
      const heightPx = emuToPx(h);
      // The metafile replay (WMF vector layers) is the main battlefield for
      // inline pictures — the flat DIB src only fills in when replay fails.
      // A flat src carries its a:srcRect crop; a replay folds the same crop
      // into its frame mapping — dropping it stretches the WHOLE source into
      // the extent box.
      const members = metafileMembers(pic, widthPx, heightPx, cropOf(pic));
      const adjust = pictureAdjustOf(pic);
      const line = outlineOf(pic.outline);
      out.push({
        kind: "picture",
        widthPx,
        heightPx,
        src: members ? undefined : pictureSrc(pic),
        crop: members ? undefined : cropOf(pic),
        members,
        // a:xfrm @rot (degrees) — Word tilts inline pictures about the
        // extent's center just like floating ones.
        ...(typeof tr.rotation === "number" && tr.rotation !== 0 ? { rotation: tr.rotation } : {}),
        ...(tr.flipHorizontal === true ? { flipH: true } : {}),
        ...(tr.flipVertical === true ? { flipV: true } : {}),
        ...(adjust?.filter ? { filter: adjust.filter } : {}),
        ...(adjust?.opacity != null ? { opacity: adjust.opacity } : {}),
        ...(adjust?.shadow ? { shadow: adjust.shadow } : {}),
        ...(line ? { line } : {}),
      });
    }
  };
  /** One nesting level's walk. `preset` carries the enclosing containers'
   *  display fields (outermost first, an inner leg overrides) merged under
   *  each run's own rPr — explicit run props always win per field. */
  const pushRuns = (items: readonly unknown[], preset: Rec): void => {
    for (const child of items) {
      if (typeof child === "string") {
        pushText(child, preset);
        continue;
      }
      if (!isRecord(child)) continue;
      // A content control is transparent in the run walk: its children project
      // as plain runs under the same preset (the control's properties are
      // metadata, not run formatting — Word's default view shows the content
      // bare).
      if (isRecord(child.sdt) && Array.isArray(child.sdt.children)) {
        pushRuns(child.sdt.children as readonly unknown[], preset);
        continue;
      }
      // Comment range markers are zero-width: a start opens tinting for every
      // text atom after it, an end closes it. The set lives across paragraphs
      // (the caller's walk), matching Word's range semantics.
      if (isRecord(child.commentRangeStart) && num(child.commentRangeStart.id) != null) {
        const commentId = num(child.commentRangeStart.id)!;
        // The range's balloon anchors at the next atom pushed (Word pins the
        // card to the range's first line).
        if (openComments) markStateful(ctx);
        if (wantsComments && !openComments?.has(commentId)) pendingComments.push(commentId);
        openComments?.add(commentId);
      }
      if (isRecord(child.commentRangeEnd) && num(child.commentRangeEnd.id) != null) {
        if (openComments) markStateful(ctx);
        openComments.delete(num(child.commentRangeEnd.id)!);
      }
      const rPr: Rec = { ...preset, ...child };
      // A footnote/endnote reference is a superscript ordinal (Word's
      // FootnoteReference/EndnoteReference style look) — numbered by
      // first-reference order, the same id twice showing the same number;
      // endnotes paint lowercase Roman (Word's endnote default numFmt). The
      // reference run's own rPr still applies.
      const fnRefId = noteRefId(child, "footnoteReference");
      if (fnRefId != null) {
        markStateful(ctx);
        const ordinal = noteOrdinal(ctx.footnoteOrdinals, fnRefId, ctx.footnoteNumStart ?? 1);
        const text = formatNumber(ctx.footnoteNumFmt ?? "decimal", ordinal);
        out.push({
          kind: "text",
          text,
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
          noteRef: { kind: "footnote", id: fnRefId, ordinal },
        });
      }
      const enRefId = noteRefId(child, "endnoteReference");
      if (enRefId != null) {
        markStateful(ctx);
        const ordinal = noteOrdinal(ctx.endnoteOrdinals, enRefId, ctx.endnoteNumStart ?? 1);
        const text = formatNumber(ctx.endnoteNumFmt ?? "lowerRoman", ordinal);
        out.push({
          kind: "text",
          text,
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
          noteRef: { kind: "endnote", id: enRefId, ordinal },
        });
      }
      if (child.footnoteRef === true) {
        out.push({
          kind: "text",
          text: formatNumber(ctx.footnoteNumFmt ?? "decimal", ctx.currentNoteOrdinal ?? 1),
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
        });
      }
      if (child.endnoteRef === true) {
        out.push({
          kind: "text",
          text: formatNumber(ctx.endnoteNumFmt ?? "lowerRoman", ctx.currentNoteOrdinal ?? 1),
          style: { ...textStyleOf(rPr), verticalAlign: "superscript" },
        });
      }
      if (typeof child.text === "string") pushText(child.text, rPr);
      // A phonetic guide (w:ruby — Word's 拼音指南): the base text stays the
      // atom's own text (the guide rides as paint metadata); consumed here so
      // the children walk below does not re-emit the base runs verbatim.
      if (isRecord(child.ruby)) pushRuby(child.ruby, rPr);
      if (child.break != null) out.push({ kind: "break" });
      if (child.tab != null) out.push({ kind: "tab" });
      if (isRecord(child.math)) {
        const style = { ...textStyleOf(rPr), italic: true, color: "#808080" };
        const label = mathLabelOf(child.math);
        const mathData = layoutMathTree(child.math, style.sizePx);
        out.push({
          kind: "math",
          label,
          widthPx: mathData?.widthPx ?? label.length * style.sizePx * 0.7 + 8,
          heightPx: mathData?.heightPx ?? style.sizePx * 1.6,
          data: mathData,
        });
      }
      if (isRecord(child.picture)) pushPicture(child.picture);
      if (isRecord(child.chart)) {
        // An inline chart paints through the picture box's member replay
        // channel (same contract as a metafile replay): the box measures from
        // the transformation, the renderer's chart painter draws the member.
        // A floating chart is an anchored drawing (projectDrawings) — skipped
        // here to avoid double-rendering.
        const chart = child.chart;
        const tr = isRecord(chart.transformation) ? chart.transformation : {};
        const w = measureEmu(tr.width);
        const h = measureEmu(tr.height);
        if (w != null && h != null && !isRecord(chart.floating)) {
          const widthPx = emuToPx(w);
          const heightPx = emuToPx(h);
          out.push({
            kind: "picture",
            widthPx,
            heightPx,
            members: [{ kind: "chart", x: 0, y: 0, width: widthPx, height: heightPx, chart }],
            ...(typeof tr.rotation === "number" && tr.rotation !== 0
              ? { rotation: tr.rotation }
              : {}),
          });
        }
      }
      if (isRecord(child.complexField)) pushField(child.complexField, rPr);
      if (isRecord(child.simpleField)) pushField(child.simpleField, rPr);
      if (isRecord(child.hyperlink) && Array.isArray(child.hyperlink.children)) {
        const link = {
          url: typeof child.hyperlink.url === "string" ? child.hyperlink.url : undefined,
          anchor: typeof child.hyperlink.anchor === "string" ? child.hyperlink.anchor : undefined,
          tooltip:
            typeof child.hyperlink.tooltip === "string" ? child.hyperlink.tooltip : undefined,
        };
        pushRuns(child.hyperlink.children, { ...preset, link });
      }
      if (isRecord(child.insertion) && Array.isArray(child.insertion.children)) {
        // Word's Display for Review: a revision outside the author filter (or
        // in simple/no-markup view) shows as accepted — plain text; the
        // original view drops shown insertions entirely. The author's palette
        // slot is assigned regardless (Word colors reviewers, not views).
        const color = revisionColor(ctx, child.insertion);
        const eff = effectiveView(ctx.markup, child.insertion);
        if (eff !== "original") {
          pushRuns(
            child.insertion.children,
            eff === "all" ? { ...preset, underline: { type: "single" }, color } : preset,
          );
        }
      }
      if (isRecord(child.deletion) && Array.isArray(child.deletion.children)) {
        const color = revisionColor(ctx, child.deletion);
        const eff = effectiveView(ctx.markup, child.deletion);
        if (eff === "all") {
          // Word moves a deletion's text into its balloon in the full markup
          // view; the anchor rides the first atom the struck runs push.
          if (wantsDeletions) {
            revisionBalloon(
              child.deletion,
              out.length,
              color,
              containerText(child.deletion.children) || undefined,
            );
          }
          pushRuns(child.deletion.children, { ...preset, strike: true, color });
        } else if (eff === "original") pushRuns(child.deletion.children, preset);
      }
      if (Array.isArray(child.children)) pushRuns(child.children, preset);
    }
  };
  pushRuns(runs, {});
  // A comment range that opened with no following text still gets its card,
  // pinned to the paragraph's last atom (or its first line when empty).
  if (pendingComments.length > 0) {
    pushCommentAnchors(out.length > 0 ? out.length - 1 : -1);
  }
  // Hidden formatting (w:vanish) with the host's Show Hidden Text off: the
  // text is neither measured nor painted, so every atom it projected is
  // suppressed — the atom keeps its source characters (the caret lattice
  // stays aligned) while the packer charges a zero advance. With the setting
  // on, only the style's `hidden` flag survives and the painter adds its
  // dotted marker.
  //
  // w:fitText (compress-to-width) stays unprojected: it needs a per-line
  // target-width compression the packer's CJK-only advance squeeze cannot
  // carry — deferred to a follow-up.
  if (!ctx.showHiddenText) {
    for (const atom of out) {
      if (atom.kind === "text" && atom.style.hidden) atom.suppressed = true;
    }
  }
  return out;
}
