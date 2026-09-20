import { describe, expect, it } from "vitest";

import { RULER_TICK_LEN, rulerLabelEvery, rulerTicks, rulerUnitPx } from "./ruler-ticks";

describe("ruler tick hierarchy", () => {
  it("subdivides an inch into eighths with the Word tick levels", () => {
    const ticks = rulerTicks({ lengthPx: 288, zeroPx: 96, unit: "in", scale: 1 });
    // One inch from the origin (96 → 192) = 8 ticks: unit, eighth, quarter,
    // eighth, half, eighth, quarter, eighth.
    const inFirstInch = ticks.filter((t) => t.pos >= 96 && t.pos < 192);
    expect(inFirstInch.map((t) => t.level)).toEqual([0, 3, 2, 3, 1, 3, 2, 3]);
    expect(inFirstInch.map((t) => t.pos)).toEqual([96, 108, 120, 132, 144, 156, 168, 180]);
    expect(inFirstInch[0]!.label).toBe(0);
    expect(inFirstInch[4]!.label).toBeUndefined();
    // Tick lengths are strictly ordered major > half > quarter > eighth.
    expect(RULER_TICK_LEN[0]).toBeGreaterThan(RULER_TICK_LEN[1]);
    expect(RULER_TICK_LEN[1]).toBeGreaterThan(RULER_TICK_LEN[2]);
    expect(RULER_TICK_LEN[2]).toBeGreaterThan(RULER_TICK_LEN[3]);
  });

  it("labels the margin gutter with negative unit numbers (Word's ruler)", () => {
    const ticks = rulerTicks({ lengthPx: 816, zeroPx: 96, unit: "in", scale: 1 });
    const majors = ticks.filter((t) => t.level === 0).map((t) => t.label);
    expect(majors).toContain(-1);
    expect(majors).toContain(0);
    expect(majors).toContain(1);
    expect(ticks[0]!.pos).toBeGreaterThanOrEqual(0);
    expect(Math.max(...ticks.map((t) => t.pos))).toBeLessThanOrEqual(816);
  });

  it("uses centimetre units and thins labels out at low zoom", () => {
    expect(rulerUnitPx("cm", 1)).toBeCloseTo(37.795, 2);
    const ticks = rulerTicks({ lengthPx: 200, zeroPx: 37.8, unit: "cm", scale: 1 });
    // 10 cm-ish strip: every unit labelled at 100%.
    expect(rulerLabelEvery("cm", 1)).toBe(1);
    expect(ticks.filter((t) => t.label !== undefined).length).toBeGreaterThan(3);

    // At 20% zoom a unit is ~7.6px — every second unit is labelled.
    expect(rulerLabelEvery("cm", 0.2)).toBe(2);
    const zoomed = rulerTicks({ lengthPx: 160, zeroPx: 7.56, unit: "cm", scale: 0.2 });
    const labels = zoomed.filter((t) => t.label !== undefined).map((t) => t.label);
    expect(labels.every((n) => n! % 2 === 0)).toBe(true);
    // Integer pixel positions keep the 1px strokes crisp at any zoom.
    expect(zoomed.every((t) => Number.isInteger(t.pos))).toBe(true);
  });

  it("mirrors the scale for RTL (positions grow leftwards from the origin)", () => {
    const ltr = rulerTicks({ lengthPx: 400, zeroPx: 100, unit: "in", scale: 1 });
    const rtl = rulerTicks({ lengthPx: 400, zeroPx: 300, unit: "in", scale: 1, mirror: true });
    // The RTL tick at +1 unit sits left of the origin; the margin gutter runs
    // to the right.
    const rtlPlus = rtl.find((t) => t.label === 1)!;
    expect(rtlPlus.pos).toBe(204);
    const rtlGutter = rtl.find((t) => t.label === -1)!;
    expect(rtlGutter.pos).toBe(396);
    // Same tick count on both sides of the origin as LTR.
    expect(rtl.length).toBe(ltr.length);
  });

  it("returns no ticks for a degenerate strip", () => {
    expect(rulerTicks({ lengthPx: 0, zeroPx: 0, unit: "in", scale: 0 })).toEqual([]);
  });

  it("covers the whole strip when the origin sits off-strip (fixed vertical ruler)", () => {
    // The page's content-top origin can lie above the pane (scrolled down) or
    // below it (scrolled up); the ticks must still span the visible strip.
    const above = rulerTicks({ lengthPx: 600, zeroPx: -150, unit: "in", scale: 1 });
    expect(above.length).toBeGreaterThan(40);
    expect(above[0]!.pos).toBeGreaterThanOrEqual(0);
    expect(above[0]!.pos).toBeLessThan(12);
    expect(Math.max(...above.map((t) => t.pos))).toBeLessThanOrEqual(600);
    expect(Math.max(...above.map((t) => t.pos))).toBeGreaterThan(588);
    expect(above.filter((t) => t.label !== undefined).length).toBeGreaterThan(3);

    const below = rulerTicks({ lengthPx: 600, zeroPx: 900, unit: "in", scale: 1 });
    expect(below[0]!.pos).toBeGreaterThanOrEqual(0);
    expect(Math.min(...below.map((t) => t.pos))).toBeLessThan(12);
    expect(Math.max(...below.map((t) => t.pos))).toBeLessThanOrEqual(600);
    expect(Math.max(...below.map((t) => t.pos))).toBeGreaterThan(588);
  });
});
