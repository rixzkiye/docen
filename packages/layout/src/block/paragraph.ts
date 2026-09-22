// Paragraph layout — OOXML line-height semantics over the packer's output.
// Ported from the editor's measure.ts (its full feature list is P1's
// acceptance bar), with the PmNode/style-cascade inputs replaced by the
// already-resolved LayoutParagraph projection.
//
// Line-height model (ECMA-376, verified against Word in the measure.ts era):
//  1. spacing.line applies to every paragraph, table cells included: exact →
//     fixed; atLeast → max(natural, spec) — the DOM route pinned atLeast to
//     its spec value as a CSS approximation; the engine does the true max;
//     multiple → factor × single line (the docGrid pitch when a grid is
//     defined, else the font natural).
//  2. No spacing.line + snapToGrid on + a grid pitch:
//     - table cell → max(natural, pitch) (the row's trHeight governs)
//     - CJK line → ceil(natural / pitch) × pitch (chars snap to the grid)
//     - picture-sized line → same ceil (the box spans whole rows; the
//       painter half-leads it — gridPadOf)
//     - Latin line → max(natural, pitch)
//  3. Otherwise → natural.
//
// The ¶-mark strut: an empty paragraph's line (and a picture row's minimum)
// is the paragraph-mark line height — spacing.line first, then the ¶-mark
// size (an absolute height), then the default run's natural metric. The
// empty line takes NO grid pitch (verified vs Word).

import {
  type LayoutBlockContext,
  type LayoutFloatZone,
  type LayoutInline,
  type LayoutLineHeight,
  type LayoutParagraph,
  wrapEffectsOf,
} from "../layout-doc";
import type { LaidOutLine, LaidOutLineItem, LaidOutParagraph } from "../layout-result";
import { packLines, type PackedLine } from "../text/line-break";
import { splitFirstGrapheme, type TextMeasurer } from "../text/measure";

/** Lay out a paragraph at `width` (its container's content width; indents
 *  shrink the usable width inside).
 *
 *  Memoized: a keystroke re-lays the whole flow, but only the edited
 *  paragraph's inputs change — the projection hands structurally identical
 *  (fresh) objects for every other block, so identity cannot drive reuse. The
 *  memo keys on a structural hash of the paragraph plus the context fields
 *  that shape it, and returns a fresh shallow copy (callers key cell boxes by
 *  paragraph identity — two identical paragraphs must not alias). A
 *  position-dependent paragraph (an anchored drawing, or a flow with
 *  absolute float zones) additionally keys on its flow position and those
 *  zones, so an unchanged paragraph still reuses its layout. */
export function layoutParagraph(
  para: LayoutParagraph,
  width: number,
  ctx: LayoutBlockContext | undefined,
  measurer: TextMeasurer,
): LaidOutParagraph {
  const key = memoKey(para, width, ctx);
  const memo = memoFor(measurer);
  const hit = memo.get(key);
  if (hit && mediaSame(para, hit)) return { ...hit };
  const laid = layoutParagraphUncached(para, width, ctx, measurer);
  if (memo.size >= MEMO_CAP) {
    const oldest = memo.keys().next().value;
    if (oldest !== undefined) memo.delete(oldest);
  }
  memo.set(key, laid);
  return { ...laid };
}

/** The memo key omits paint-only media sources (megabyte data URLs), so a hit
 *  must still prove them unchanged — a replaced image with identical geometry
 *  re-lays instead of painting the old bitmap. String comparison is a memcmp
 *  (or O(1) when the projection's identity cache hands back the same string),
 *  far below hashing every base64 char. The member/array structure is part of
 *  the key, so the walks below line up. */
function mediaSame(para: LayoutParagraph, laid: LaidOutParagraph): boolean {
  const a = para.inline;
  const b = laid.inline;
  for (let i = 0; i < a.length; i++) {
    const x = a[i]!;
    if (x.kind !== "picture") continue;
    const y = b[i];
    if (y?.kind !== "picture" || x.src !== y.src) return false;
  }
  const da = para.drawings;
  const db = laid.drawings;
  if (da === db) return true;
  if (!da || !db || da.length !== db.length) return false;
  for (let i = 0; i < da.length; i++) {
    const ma = da[i]!.members;
    const mb = db[i]!.members;
    if (ma === mb) continue;
    if (ma.length !== mb.length) return false;
    for (let k = 0; k < ma.length; k++) {
      const x = ma[k]!;
      const y = mb[k]!;
      if (x.kind !== "picture" || y.kind !== "picture") continue;
      if (x.src !== y.src) return false;
    }
  }
  return true;
}

/** Line breaking is the flow's dominant cost on large documents (see
 *  flow-scale.bench.ts) — one map per measurer, FIFO-capped so a long editing
 *  session cannot grow it without bound. */
const MEMO_CAP = 2048;
const memoByMeasurer = new WeakMap<TextMeasurer, Map<string, LaidOutParagraph>>();

function memoFor(measurer: TextMeasurer): Map<string, LaidOutParagraph> {
  let memo = memoByMeasurer.get(measurer);
  if (!memo) {
    memo = new Map();
    memoByMeasurer.set(measurer, memo);
  }
  return memo;
}

/** Structural FNV-1a pair over the layout inputs — a memo key without
 *  JSON.stringify's string allocation, which the typing path would pay for
 *  every paragraph on every render. Two independent hashes make a collision
 *  (a stale layout) practically impossible; both sides of a comparison come
 *  from the same projection code, so property order is deterministic.
 *
 *  Key names repeat across every paragraph and style object, so their hash
 *  is computed once and cached — hashing ~30 names per object was most of
 *  the per-keystroke key cost on large documents. */
const keyNameHashes = new Map<string, number>();
function hashKeyName(name: string): number {
  let h = keyNameHashes.get(name);
  if (h === undefined) {
    h = 0x811c9dc5;
    for (let i = 0; i < name.length; i++) h = Math.imul(h ^ name.charCodeAt(i), 0x01000193);
    keyNameHashes.set(name, h);
  }
  return h;
}

/** Streaming structural FNV-1a pair — the memo key's hasher. Two
 *  independent hashes make a collision (a stale layout) practically
 *  impossible; both sides of a comparison come from the same projection
 *  code, so property order is deterministic. Key names repeat across every
 *  paragraph and style object, so their hash is computed once and cached —
 *  hashing ~30 names per object was most of the per-keystroke key cost on
 *  large documents. */
class KeyMixer {
  h1: number;
  h2: number;

  constructor(h1 = 0x811c9dc5, h2 = 0x9e3779b9) {
    this.h1 = h1;
    this.h2 = h2;
  }

  mix(v: number): void {
    this.h1 = Math.imul(this.h1 ^ v, 0x01000193);
    this.h2 = Math.imul(this.h2 ^ v, 0x85ebca6b);
  }

  mixStr(s: string): void {
    for (let i = 0; i < s.length; i++) this.mix(s.charCodeAt(i));
  }

  walk(v: unknown): void {
    if (v === null || v === undefined) {
      this.mix(v === null ? 0x11 : 0x12);
      return;
    }
    switch (typeof v) {
      case "string":
        this.mix(0x13);
        this.mixStr(v);
        return;
      case "number":
        this.mix(0x14);
        // Sub-thousandth differences never change layout — quantize so tiny
        // float noise (e.g. a re-derived percentage) cannot churn the memo.
        this.mix(Math.round(v * 1000));
        return;
      case "boolean":
        this.mix(v ? 0x15 : 0x16);
        return;
      case "object": {
        if (Array.isArray(v)) {
          this.mix(0x17);
          for (const item of v) this.walk(item);
          return;
        }
        this.mix(0x18);
        for (const key of Object.keys(v as Record<string, unknown>)) {
          // Renderer-only media source (a megabyte data URL): it never shapes
          // the layout — the picture atom's box is widthPx/heightPx alone (see
          // LayoutInline) — and a hit is re-checked for it separately, so the
          // memo does not walk every base64 char on every render.
          if (key === "src") continue;
          this.mix(hashKeyName(key));
          this.walk((v as Record<string, unknown>)[key]);
        }
        return;
      }
      default:
        this.mixStr(String(v));
    }
  }

  key(): string {
    return `${(this.h1 >>> 0).toString(36)}:${(this.h2 >>> 0).toString(36)}`;
  }
}

/** Position-free keys depend only on the paragraph, the width, and a handful
 *  of flow scalars — all stable across a rerun of the same paragraph object
 *  (the flow re-lays overflow blocks from its page replay). One entry per
 *  object; the scalars are compared, not just the width, so a cell paragraph
 *  and a body paragraph cannot share a key. The raw hash pair is kept too:
 *  a position-dependent paragraph mixes its flow context into the pair
 *  without re-walking the paragraph body on every render. */
const keyCache = new WeakMap<
  LayoutParagraph,
  {
    width: number;
    inTable: boolean;
    adjust: boolean;
    onGrid: boolean;
    pitch: number;
    h1: number;
    h2: number;
    key: string;
  }
>();

function memoKey(
  para: LayoutParagraph,
  width: number,
  ctx: LayoutBlockContext | undefined,
): string {
  const inTable = ctx?.inTable === true;
  const adjust = ctx?.adjustLinesInTable === true;
  const onGrid = ctx?.onGrid === true;
  const pitch = Math.round((ctx?.linePitchPx ?? 0) * 1000);
  // Position-dependent paragraphs (an anchored drawing, or a page carrying
  // absolute float zones) add their flow context below.
  const positional = (para.drawings?.length ?? 0) > 0 || (ctx?.floatZones?.length ?? 0) > 0;
  let cached = keyCache.get(para);
  if (
    !cached ||
    cached.width !== width ||
    cached.inTable !== inTable ||
    cached.adjust !== adjust ||
    cached.onGrid !== onGrid ||
    cached.pitch !== pitch
  ) {
    const mixer = new KeyMixer();
    mixer.walk(para);
    mixer.mix(0x19);
    mixer.mix(Math.round(width * 1000));
    mixer.mix(inTable ? 1 : 0);
    mixer.mix(adjust ? 1 : 0);
    mixer.mix(onGrid ? 1 : 0);
    mixer.mix(pitch);
    cached = {
      width,
      inTable,
      adjust,
      onGrid,
      pitch,
      h1: mixer.h1,
      h2: mixer.h2,
      key: "",
    };
    keyCache.set(para, cached);
  }
  if (!positional) {
    if (cached.key === "") cached.key = formatKey(cached.h1, cached.h2);
    return cached.key;
  }
  // The flow context: the zones scan from the block's flow Y, and a
  // paragraph-anchored drawing resolves against the page box — an unchanged
  // paragraph must not reuse a layout from another position.
  const mixer = new KeyMixer(cached.h1, cached.h2);
  mixer.mix(0x1a);
  mixer.mix(Math.round((ctx?.startY ?? 0) * 1000));
  if (ctx?.wrapPage) {
    mixer.mix(0x1b);
    mixer.walk(ctx.wrapPage);
  }
  if (ctx?.floatZones?.length) {
    mixer.mix(0x1c);
    mixer.walk(ctx.floatZones);
  }
  return mixer.key();
}

function formatKey(h1: number, h2: number): string {
  return `${(h1 >>> 0).toString(36)}:${(h2 >>> 0).toString(36)}`;
}

function layoutParagraphUncached(
  para: LayoutParagraph,
  width: number,
  ctx: LayoutBlockContext | undefined,
  measurer: TextMeasurer,
): LaidOutParagraph {
  const spec = para.spacing?.lineHeight;
  const inTable = ctx?.inTable ?? false;
  // Cell lines join the document grid only under the adjustLineHeightInTable
  // compat (CT_Compat: absence leaves cells grid-free — CJK-Word documents
  // carry it, most Western ones don't).
  const pitch =
    para.snapToGrid === false || (inTable && ctx?.adjustLinesInTable !== true)
      ? 0
      : (ctx?.linePitchPx ?? 0);

  // The ¶-mark strut line: spacing wins, then the mark size, then the default
  // run's natural metric (no grid pitch — see the module doc).
  const strutNatural = para.defaultTextStyle ? measurer.naturalOf(para.defaultTextStyle) : 0;
  let strutPx: number;
  if (spec) {
    strutPx = resolveLine(spec, strutNatural, pitch);
  } else {
    strutPx = para.markSizePx ?? strutNatural;
  }
  if (strutPx <= 0) {
    const firstText = para.inline.find((i) => i.kind === "text");
    strutPx = firstText?.style.sizePx ?? 16;
  }

  const usable = Math.max(0, width - (para.indent?.leftPx ?? 0) - (para.indent?.rightPx ?? 0));

  // The anchor paragraph wraps beside its own square floats: paragraph/
  // column anchors are paragraph-relative by definition, so their zones feed
  // the packer in paragraph-relative Y (the flow's — or the table cell's —
  // absolute zones cover every later paragraph); page/margin anchors resolve
  // through ctx.wrapPage with −startY translating flow Y into the local
  // space. The box grows by the anchor's wrap distances first
  // (distL/R/T/B), matching the flow's zone padding. Full-column boxes and
  // topAndBottom clears are bands — the packer cannot skip a mid-paragraph
  // band, so they stay flow-only. Table cells qualify too (Word's
  // layoutInCell): a cell-anchored float wraps the cell's text, and inside a
  // cell `column` IS the cell's column (wrapPage never reaches a cell).
  const selfZones: LayoutFloatZone[] = wrapEffectsOf(
    para.drawings,
    0,
    width,
    inTable,
    ctx?.wrapPage && ctx.startY != null ? { ...ctx.wrapPage, flowZeroPx: -ctx.startY } : undefined,
  ).zones;

  // A drop cap renders as one enlarged glyph in its own box; the flow must
  // not repeat it. Split the leading grapheme off the first text atom before
  // packing and hand it to the painter through `dropCapGlyph`.
  let inline = para.inline;
  let dropCapGlyph: LaidOutParagraph["dropCapGlyph"];
  if (para.dropCap) {
    const first = para.inline[0];
    if (first && first.kind === "text" && first.text.length > 0) {
      const [glyph, rest] = splitFirstGrapheme(first.text);
      dropCapGlyph = { text: glyph, style: first.style };
      inline = rest ? [{ ...first, text: rest }, ...para.inline.slice(1)] : para.inline.slice(1);
    }
  }

  if (para.dropCap && para.inline.length > 0) {
    const lines = para.dropCap.lines ?? 3;
    const h = lines * strutPx;
    const capWidth = Math.round(h * 0.75) + (para.dropCap.distancePx ?? 0);
    selfZones.push({
      x0Px: 0,
      widthPx: capWidth,
      topPx: 0,
      bottomPx: h,
      textAfter: true,
    });
  }

  // Word never applies a first-line indent to centered or right-aligned
  // paragraphs: suppress a dirty firstLinePx at the layout level too, so a
  // caller-built LayoutParagraph cannot skew the axis the projection fixed.
  const isCenteredOrRight = para.align === "center" || para.align === "right";
  const packed = packLines(inline, {
    measurer,
    width: usable,
    firstLineIndentPx: isCenteredOrRight ? undefined : para.indent?.firstLinePx,
    tabStops: para.tabStops,
    defaultTabStopPx: para.defaultTabStopPx,
    strutPx,
    lineHeight: ({ naturalPx }) =>
      spec ? resolveLine(spec, naturalPx, pitch) : snapLine(naturalPx, pitch),
    startY: ctx?.startY,
    bidi: para.bidi,
    kinsoku: para.kinsoku,
    overflowPunct: para.overflowPunctuation,
    compressPunctuation: para.compressPunctuation,
    wordWrap: para.wordWrap,
    autoSpaceDE: para.autoSpaceDE,
    textAlignment: para.textAlignment,
    // Absolute zones come from whoever stacks the blocks: the flow passes the
    // page's float zones, the table cell stacker accumulates the cell's own
    // (a cell's width is its column — page floats never reach it because the
    // cellCtx starts empty, not because the paragraph drops them here).
    floatZones: ctx?.floatZones,
    selfZones: selfZones.length > 0 ? selfZones : undefined,
  });

  // An empty paragraph still occupies one strut line (the ¶ glyph's).
  if (packed.length === 0) {
    packed.push({
      items: [],
      endInlineIndex: 0,
      final: true,
      maxWidthPx: usable,
      heightPx: strutPx,
      naturalPx: strutNatural,
    });
  }

  // w:jc — horizontal alignment, one pass over the packed lines (the layout
  // owns the geometry; the painter places what it is given):
  //  - both/distribute stretch inter-character gaps to the line's content
  //    width. both skips the paragraph's last line and hard-break lines (they
  //    are that logical line's natural end); distribute stretches every line.
  //  - center/right shift the whole line's items by the slack after the
  //    content (trailing whitespace hangs and never counts).
  const justifyGapPx: (number | undefined)[] = packed.map(() => undefined);
  const stretchAll = para.align === "distribute";
  if (para.align === "both" || stretchAll) {
    for (let i = 0; i < packed.length; i++) {
      const line = packed[i];
      if (line.items.length === 0) continue;
      if (!stretchAll && i === packed.length - 1) continue;
      // A hard break ends this line — it is that logical line's last line.
      if (para.inline[line.endInlineIndex]?.kind === "break") continue;
      justifyGapPx[i] = justifyLine(line, para.inline, measurer);
    }
  } else if (para.align === "center" || para.align === "right") {
    for (const line of packed) {
      const items = line.items;
      if (items.length === 0) continue;
      const last = items[items.length - 1];
      const contentEnd =
        last.xPx +
        last.widthPx -
        Math.max(trailingHang(items, para.inline, measurer), line.hangPx ?? 0);
      const slack = line.maxWidthPx - contentEnd;
      if (slack <= 0) continue;
      const shift = para.align === "center" ? slack / 2 : slack;
      for (const it of items) it.xPx += shift;
    }
  }

  const lines: LaidOutLine[] = [];
  let heightPx = 0;
  // Grid lattice: only the body flow (ctx.onGrid) centers grid-height lines;
  // exact/atLeast spacing overrides the grid's height and with it the centering.
  const gridLine =
    pitch > 0 && ctx?.onGrid === true && spec?.rule !== "exact" && spec?.rule !== "atLeast";
  for (let i = 0; i < packed.length; i++) {
    const line = packed[i];
    lines.push({
      yPx: heightPx,
      heightPx: line.heightPx,
      naturalPx: line.naturalPx,
      textEmPx: line.textEmPx,
      grid: gridLine || undefined,
      spacingRule: spec?.rule,
      pictureFloored: line.pictureFloored,
      advanceScale: line.advanceScale,
      firstLineIndentPx: i === 0 && !isCenteredOrRight ? para.indent?.firstLinePx : undefined,
      endInlineIndex: line.endInlineIndex,
      final: line.final,
      items: line.items,
      // Carried on every line (not just justified ones) — the wrap width is
      // the selection highlight's right edge (Word highlights to the wrap
      // edge, not the last glyph) and hit-testing's line-end boundary.
      maxWidthPx: line.maxWidthPx,
      justifyGapPx: justifyGapPx[i],
      hangPx: line.hangPx,
      xOffsetPx: line.xOffsetPx,
    });
    heightPx += line.heightPx;
  }

  return {
    kind: "paragraph",
    heightPx,
    beforePx: para.spacing?.beforePx ?? 0,
    afterPx: para.spacing?.afterPx ?? 0,
    lines,
    inline: para.inline,
    align: para.align,
    keepLines: para.keepLines,
    suppressLineNumbers: para.suppressLineNumbers,
    keepNext: para.keepNext,
    widowControl: para.widowControl,
    borders: para.borders,
    shadingFill: para.shadingFill,
    indent: para.indent,
    tabStops: para.tabStops,
    drawings: para.drawings,
    dropCap: para.dropCap,
    dropCapGlyph,
    markSizePx: para.markSizePx,
    preserveSpaces: true,
    sectionEnd: para.sectionEnd,
    bidi: para.bidi,
    textDirection: para.textDirection,
    formatChange: para.formatChange,
    balloons: para.balloons,
  };
}

/** The width of the last text item's trailing whitespace — it hangs past the
 *  line's right edge and never counts toward justification or alignment. */
function trailingHang(
  items: readonly LaidOutLineItem[],
  inline: readonly LayoutInline[],
  measurer: TextMeasurer,
): number {
  const last = items[items.length - 1];
  if (last.kind !== "text") return 0;
  const src = inline[last.inlineIndex];
  if (src?.kind !== "text") return 0;
  const trail = /\s+$/.exec(last.text)?.[0];
  // Measured like the breaker charges it (same whitespace mode): under
  // pre-wrap a line-trailing run keeps every space, and its true advance —
  // not a strut-height stand-in — is what justification must not stretch.
  return trail ? measurer.widthOf(trail, src.style, "pre-wrap") : 0;
}

/** Stretch one wrapped line to its packed width: the slack after the last
 *  item's content (trailing whitespace AND an overflow-punct hang — both
 *  reach past the right edge, both excluded) spread over justify units —
 *  CJK items stretch per inter-grapheme gap (Word's CJK justification),
 *  Latin items per word gap (spaces absorb the slack; letter-spreading
 *  English is the letter-mode tell-tale). Re-spaces every item's x in place
 *  and returns the per-unit stretch (0 when there is nothing to spread). */
function justifyLine(
  line: PackedLine,
  inline: readonly LayoutInline[],
  measurer: TextMeasurer,
): number {
  const items = line.items;
  const last = items[items.length - 1];
  const hang = Math.max(trailingHang(items, inline, measurer), line.hangPx ?? 0);
  const itemUnits = items.map((it) => {
    if (it.kind !== "text") return 1;
    if (CJK_ITEM.test(it.text)) {
      let points = 0;
      for (const _ of it.text) points++;
      return points - 1;
    }
    const indices = leaferWordIndices(it.text);
    return indices[indices.length - 1] ?? 0;
  });
  let units = 0;
  for (const u of itemUnits) units += Math.max(u, 0);
  const delta =
    units > 0 ? Math.max(0, line.maxWidthPx - (last.xPx + last.widthPx - hang)) / units : 0;
  if (delta === 0) return 0;
  let before = 0;
  items.forEach((it, i) => {
    it.xPx += delta * before;
    before += itemUnits[i]!;
  });
  return delta;
}

/** Items stretch per grapheme when they hold any CJK glyph (the painter's
 *  "both-letter" trigger — painter and caret map consume this same test). */
const CJK_ITEM = /[一-鿿぀-ヿ가-힯]/;

/** Whether a text item justifies per grapheme (any CJK glyph) rather than
 *  per word gap — Leafer's both-letter vs both-justify choice, shared by
 *  the painter's textAlign and the caret map's boundary distribution. */
export function justifyPerGrapheme(text: string): boolean {
  return CJK_ITEM.test(text);
}

/** Leafer's word split (its justify denominator): a space or one of its
 *  break chars, or a CJK glyph, stands alone as one word; other runs
 *  coalesce. Maps each grapheme to its word index — the count is the last
 *  index + 1, and a grapheme's justify shift is its index × the per-gap
 *  stretch (the painter's "both-justify" Text applies exactly this). */
export function leaferWordIndices(text: string): number[] {
  const indices: number[] = [];
  let word = -1;
  let inRun = false;
  for (const ch of text) {
    if (ch === " " || LEAFER_BREAK_CHARS.has(ch) || CJK_ITEM.test(ch)) {
      word++;
      inRun = false;
    } else {
      if (!inRun) word++;
      inRun = true;
    }
    indices.push(word);
  }
  return indices;
}

const LEAFER_BREAK_CHARS = new Set(["-", "—", "／", "～", "｜", "┆", "·"]);

/** A spacing.line spec against a line's natural height. On a grid, a body CJK
 *  line's spec'd height never falls below the line's natural height and snaps
 *  up to whole rows; a cell line (w:adjustLineHeightInTable, on in every
 *  CJK-Word document) adds grid pitch only as far as the line's own natural
 *  height demands — the multiple's demand is factor × pitch, and the natural
 *  height snaps up to whole rows *before* the comparison. Corpus-verified
 *  (honor table, 340-twip grid): the 14pt header cell (natural 24.2px >
 *  pitch 22.7px) takes 2 rows (45.3px) while every 12pt/10.5pt row stays at
 *  1.5 × pitch (34px); shrinking the header run to 12pt moves its bottom
 *  border from y215 to y204 (= 1.5 × pitch), confirming natural — not factor
 *  — drives the snap. */

/** Round a height up to whole grid rows — the lattice snap every on-grid
 *  branch applies (a CJK/picture/cell line always spans whole rows). */
function snapUpToPitch(px: number, pitch: number): number {
  return Math.ceil(px / pitch) * pitch;
}

function resolveLine(spec: LayoutLineHeight, naturalPx: number, pitch: number): number {
  if (spec.rule === "exact") return spec.px;
  if (spec.rule === "atLeast") return Math.max(naturalPx, spec.px);
  // multiple: 240ths of a single line — the grid pitch when defined, else the
  // font natural (verified vs Word). On a grid every line takes the larger of
  // the multiple's demand (factor × pitch) and its own natural height snapped
  // up to whole rows — body CJK, Latin, cell and picture lines alike.
  // Word-COM verified: an 11pt body line stays at factor × pitch (18.04pt on a
  // 15.6pt grid, never snapped to the lattice) while a 24pt heading line spans
  // 2 rows (31.2pt).
  if (pitch > 0) return Math.max(spec.factor * pitch, snapUpToPitch(naturalPx, pitch));
  return spec.factor * naturalPx;
}

/** No spacing.line: snap the natural height to the document grid. */
function snapLine(naturalPx: number, pitch: number): number {
  if (pitch <= 0) return naturalPx;
  return snapUpToPitch(naturalPx, pitch);
}
