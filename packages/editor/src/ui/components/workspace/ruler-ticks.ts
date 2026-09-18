// Shared ruler tick math — the interactive horizontal `<docen-ruler>` and the
// canvas stage's vertical strip render the same Word tick hierarchy from one
// computation: a labelled major tick every unit, half ticks, quarter ticks,
// and the smallest eighth ticks. Keeping the geometry here means the two
// rulers can never drift apart.

export type RulerUnit = "in" | "cm";

/** Tick hierarchy level: 0 = unit (labelled), 1 = half, 2 = quarter, 3 = eighth. */
export type RulerTickLevel = 0 | 1 | 2 | 3;

export interface RulerTick {
  /** Position in CSS px along the strip (from the screen-start edge). */
  pos: number;
  level: RulerTickLevel;
  /** Integer unit number for level-0 ticks — signed, so the margin gutter
   *  reads -2, -1, 0, 1, … like Word's ruler. Omitted where labels are
   *  thinned out (low zoom) and on non-major ticks. */
  label?: number;
}

export interface RulerTickOptions {
  /** Strip length in CSS px (already zoom-scaled). */
  lengthPx: number;
  /** The content-origin offset in CSS px — Word's margin line, where 0 sits. */
  zeroPx: number;
  unit: RulerUnit;
  /** Zoom factor (1 = 100%). */
  scale: number;
  /** Mirror the scale for RTL: positions grow leftwards from `zeroPx`. */
  mirror?: boolean;
}

/** Tick lengths in CSS px, indexed by hierarchy level (unit → eighth). */
export const RULER_TICK_LEN: Record<RulerTickLevel, number> = { 0: 7, 1: 5, 2: 3, 3: 2 };

/** One screen px per measurement unit at the given zoom. */
export function rulerUnitPx(unit: RulerUnit, scale: number): number {
  return (unit === "cm" ? 96 / 2.54 : 96) * scale;
}

/** Word's numbers stop fitting below ~22px per unit; every second unit is
 *  labelled there instead of overlapping digits. */
export function rulerLabelEvery(unit: RulerUnit, scale: number): number {
  return rulerUnitPx(unit, scale) < 22 ? 2 : 1;
}

/**
 * Every tick of the four-level hierarchy covering the whole strip, including
 * the margin gutter (negative offsets). Positions are rounded to integer CSS
 * px so a 1px stroke lands on the pixel grid without blur.
 */
export function rulerTicks(o: RulerTickOptions): RulerTick[] {
  const unitPx = rulerUnitPx(o.unit, o.scale);
  const step = unitPx / 8;
  if (!(step > 0) || !Number.isFinite(step)) return [];

  // Offsets run from the negative margin gutter to the strip's far edge; in
  // RTL the strip flips around the origin (the right-margin line).
  const gutter = o.mirror ? o.lengthPx - o.zeroPx : o.zeroPx;
  const far = o.mirror ? o.zeroPx : o.lengthPx - o.zeroPx;
  const first = Math.ceil(-gutter / step - 1e-6);
  const last = Math.floor(far / step + 1e-6);
  const labelEvery = rulerLabelEvery(o.unit, o.scale);
  const ticks: RulerTick[] = [];
  for (let e = first; e <= last; e++) {
    const eighths = Math.abs(e) % 8;
    const level: RulerTickLevel = eighths === 0 ? 0 : eighths === 4 ? 1 : eighths % 2 === 0 ? 2 : 3;
    const tick: RulerTick = { pos: Math.round(o.zeroPx + (o.mirror ? -e : e) * step), level };
    if (level === 0 && e % (8 * labelEvery) === 0) tick.label = e / 8;
    ticks.push(tick);
  }
  return ticks;
}
