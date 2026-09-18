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

  // Group glyphs by their cluster start character index
  interface ClusterBounds {
    charStart: number;
    charEnd: number;
    startX: number;
    endX: number;
    isRtl: boolean;
  }

  const clusterMap = new Map<number, { minX: number; maxX: number }>();
  const isRtl = glyphRun.direction === "rtl";

  for (let i = 0; i < glyphRun.glyphs.length; i++) {
    const g = glyphRun.glyphs[i]!;
    const cStart = utf8ToUtf16[g.cluster] ?? g.cluster;
    const gNextX =
      i + 1 < glyphRun.glyphs.length ? glyphRun.glyphs[i + 1]!.xPx : glyphRun.totalAdvancePx;
    const gMinX = Math.min(g.xPx, gNextX);
    const gMaxX = Math.max(g.xPx, gNextX);

    const existing = clusterMap.get(cStart);
    if (!existing) {
      clusterMap.set(cStart, { minX: gMinX, maxX: gMaxX });
    } else {
      existing.minX = Math.min(existing.minX, gMinX);
      existing.maxX = Math.max(existing.maxX, gMaxX);
    }
  }

  // Sort distinct character starts
  const sortedStarts = Array.from(clusterMap.keys()).sort((a, b) => a - b);
  const clusters: ClusterBounds[] = [];

  for (let i = 0; i < sortedStarts.length; i++) {
    const cStart = sortedStarts[i]!;
    const cEnd = i + 1 < sortedStarts.length ? sortedStarts[i + 1]! : text.length;
    const bounds = clusterMap.get(cStart)!;
    clusters.push({
      charStart: cStart,
      charEnd: cEnd,
      startX: bounds.minX,
      endX: bounds.maxX,
      isRtl,
    });
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
      const fracStart = cluster.isRtl
        ? (cluster.charEnd - gStart - g.length) / clusterSpan
        : (gStart - cluster.charStart) / clusterSpan;
      const fracLen = g.length / clusterSpan;

      const gx = cluster.startX + Math.max(0, fracStart) * clusterWidth;
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
