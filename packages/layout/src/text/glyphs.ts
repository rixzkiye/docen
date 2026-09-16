// The per-grapheme glyph geometry of one laid text item — the ONE derived
// model behind the painter's formatting marks/emphasis marks and the editor's
// caret map, so a mark, a caret and a selection edge can never disagree with
// what the painter's Text actually renders. It folds in the run's display
// transform (w:caps / w:smallCaps), the piece's painted size, and the w:w
// horizontal scale: the item's glyph layout is measured in natural space and
// scaled — matching a painted element whose glyphs ride at the piece size with
// `scaleX = w:w` and the run's letter spacing (which the scale stretches too).

import { isCjkCodeUnit } from "../font";
import type { LayoutTextStyle } from "../layout-doc";
import { itemGlyphLayout, type ItemGlyphLayout } from "./advance";
import { characterScaleOf, cssFontAtSize, familyOfSlot, vertAlignedSizePx } from "./measure";

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

/** Glyph placement of one laid text item: the item's painted text (caps
 *  display form) measured in the piece's font, with the run's w:w scale applied
 *  to every advance and to the justify/compress interval — `intervalPx` is in
 *  painted space, so the natural-space layout fills `intervalPx / scale` before
 *  scaling back. */
export function itemGlyphLayoutOf(
  item: { text: string; displayText?: string; fontSizePx?: number },
  style: LayoutTextStyle,
  intervalPx?: number,
): ItemGlyphLayout {
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
