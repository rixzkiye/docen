// The per-grapheme glyph geometry of one laid text item — the ONE derived
// model behind the painter's formatting marks/emphasis marks and the editor's
// caret map, so a mark, a caret and a selection edge can never disagree with
// what the painter's Text actually renders. It folds in the run's display
// transform (w:caps / w:smallCaps), the piece's painted size, and the w:w
// horizontal scale: the item's glyph layout is measured in natural space and
// scaled — matching a painted element whose glyphs ride at the piece size with
// `scaleX = w:w` and the run's letter spacing (which the scale stretches too).

import { justifyPerGrapheme, leaferWordIndices } from "../block/paragraph";
import { isCjkCodeUnit } from "../font";
import type { LayoutTextStyle } from "../layout-doc";
import type { LaidOutGlyphRun } from "../layout-result";
import { itemGlyphLayout, type ItemGlyphLayout } from "./advance";
import { characterScaleOf, cssFontAtSize, familyOfSlot, vertAlignedSizePx } from "./measure";

const SEGMENTER = new Intl.Segmenter(undefined, { granularity: "grapheme" });

/** The painted (and measured, at scale 1) glyph size of one laid text item:
 *  the smallCaps piece override, else the run's vertAlign-scaled size. */
export function itemSizePxOf(style: LayoutTextStyle, fontSizePx?: number): number {
  return fontSizePx ?? vertAlignedSizePx(style);
}

/** One laid text item's font shorthand for an explicit size — the paint font
 *  (piece size) for ink metrics; the advance font multiplies the size by the
 *  w:w scale when the caller measures rather than paints. */
export function itemFontOf(
  style: LayoutTextStyle,
  item: { text: string; displayText?: string },
  sizePx?: number,
): string {
  const text = item.displayText ?? item.text;
  return cssFontAtSize(
    style,
    familyOfSlot(style.family, isCjkCodeUnit(text, 0)),
    itemSizePxOf(style, sizePx),
  );
}

/**
 * Derives grapheme-exact caret and selection geometry directly from a shaped OpenType glyph run.
 * Accurately divides ligature clusters and handles complex script clusters.
 */
export function itemGlyphLayoutFromRun(
  text: string,
  glyphRun: LaidOutGlyphRun,
  style: LayoutTextStyle,
  intervalPx?: number,
): ItemGlyphLayout {
  const scale = characterScaleOf(style);
  const graphemes: { segment: string; index: number; length: number }[] = [];
  for (const { segment, index } of SEGMENTER.segment(text)) {
    graphemes.push({ segment, index, length: segment.length });
  }

  if (graphemes.length === 0) {
    return { xs: [], widths: [], lens: [], endX: 0 };
  }

  // Build map from UTF-8 byte offset to UTF-16 code unit offset
  const utf8ToUtf16: number[] = [];
  let u16 = 0;
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!;
    const u16Len = cp > 0xffff ? 2 : 1;
    const u8Len = cp < 0x80 ? 1 : cp < 0x800 ? 2 : cp < 0x10000 ? 3 : 4;
    for (let b = 0; b < u8Len; b++) {
      utf8ToUtf16.push(u16);
    }
    u16 += u16Len;
    i += u16Len;
  }
  utf8ToUtf16.push(text.length);

  // Group glyphs by distinct cluster boundaries
  interface ClusterInfo {
    charStart: number;
    charEnd: number;
    startX: number;
    endX: number;
  }

  const clusters: ClusterInfo[] = [];
  const glyphs = glyphRun.glyphs;

  for (let i = 0; i < glyphs.length; i++) {
    const g = glyphs[i]!;
    const charStart = utf8ToUtf16[g.cluster] ?? g.cluster;
    const startX = g.xPx;
    // Find where this cluster ends in character space and pixel space
    let nextCharStart = text.length;
    let nextX = glyphRun.totalAdvancePx;

    for (let j = i + 1; j < glyphs.length; j++) {
      const nextG = glyphs[j]!;
      const c = utf8ToUtf16[nextG.cluster] ?? nextG.cluster;
      if (c !== charStart) {
        nextCharStart = c;
        nextX = nextG.xPx;
        break;
      }
    }

    // Only record new cluster or merge glyphs in same cluster
    const last = clusters[clusters.length - 1];
    if (!last || last.charStart !== charStart) {
      clusters.push({
        charStart,
        charEnd: nextCharStart,
        startX,
        endX: nextX,
      });
    }
  }

  // Assign widths and xs to each grapheme based on cluster interpolation
  const xs: number[] = [];
  const widths: number[] = [];
  const lens: number[] = [];

  for (const g of graphemes) {
    lens.push(g.length);
    const gStart = g.index;

    // Find cluster matching this grapheme
    const cluster =
      clusters.find((c) => gStart >= c.charStart && gStart < c.charEnd) ??
      clusters[clusters.length - 1];
    if (cluster) {
      const clusterSpan = Math.max(1, cluster.charEnd - cluster.charStart);
      const clusterWidth = Math.max(0, cluster.endX - cluster.startX);
      const fracStart = (gStart - cluster.charStart) / clusterSpan;
      const fracLen = g.length / clusterSpan;

      const gx = cluster.startX + fracStart * clusterWidth;
      const gw = fracLen * clusterWidth;
      xs.push(gx);
      widths.push(gw);
    } else {
      const fallbackW = glyphRun.totalAdvancePx / graphemes.length;
      xs.push(xs.length > 0 ? xs[xs.length - 1]! + widths[widths.length - 1]! : 0);
      widths.push(fallbackW);
    }
  }

  let endX = glyphRun.totalAdvancePx;

  // Apply justification / squeezing if intervalPx is specified
  if (intervalPx != null && graphemes.length > 1) {
    const natural = endX;
    if (justifyPerGrapheme(text)) {
      const delta = (intervalPx - natural) / (graphemes.length - 1);
      for (let g = 0; g < graphemes.length; g++) xs[g]! += g * delta;
      endX = intervalPx;
    } else {
      const indices = leaferWordIndices(text);
      const words = (indices[graphemes.length - 1] ?? 0) + 1;
      if (words > 1) {
        const gap = (intervalPx - natural) / (words - 1);
        for (let g = 0; g < graphemes.length; g++) xs[g]! += (indices[g] ?? 0) * gap;
        endX = intervalPx;
      }
    }
  }

  if (scale === 1) {
    return { xs, widths, lens, endX };
  }

  return {
    xs: xs.map((x) => x * scale),
    widths: widths.map((w) => w * scale),
    lens,
    endX: endX * scale,
  };
}

/** Glyph placement of one laid text item: the item's painted text (caps
 *  display form) measured in the piece's font, with the run's w:w scale applied
 *  to every advance and to the justify/compress interval — `intervalPx` is in
 *  painted space, so the natural-space layout fills `intervalPx / scale` before
 *  scaling back. */
export function itemGlyphLayoutOf(
  item: { text: string; displayText?: string; fontSizePx?: number; glyphRun?: LaidOutGlyphRun },
  style: LayoutTextStyle,
  intervalPx?: number,
): ItemGlyphLayout {
  if (item.glyphRun && item.glyphRun.glyphs.length > 0) {
    return itemGlyphLayoutFromRun(item.displayText ?? item.text, item.glyphRun, style, intervalPx);
  }

  const scale = characterScaleOf(style);
  const layout = itemGlyphLayout(
    item.displayText ?? item.text,
    itemFontOf(style, item, item.fontSizePx),
    style.letterSpacingPx,
    intervalPx == null ? undefined : intervalPx / scale,
  );
  if (scale === 1) return layout;
  return {
    xs: layout.xs.map((x) => x * scale),
    widths: layout.widths.map((w) => w * scale),
    lens: layout.lens,
    endX: layout.endX * scale,
  };
}
