// Canvas font-kerning activation (docen-local): `fontKerning: "normal"` must
// be applied to the measurement context for the whole preparation when a run's
// w:kern threshold is met, and the kerned widths must never be served from the
// default-mode segment cache. The fake context below is the only canvas in
// this file's module registry; it records every `fontKerning` write and narrows
// "AV" under "normal" so the option's effect on measured widths is observable.

import { beforeEach, describe, expect, it } from "vitest";

import { measureNaturalWidth, prepareWithSegments } from "./layout";
import {
  clearMeasurementCaches,
  fontKerningSupported,
  getFontMeasurementState,
  getSegmentMetrics,
} from "./measurement";

const FONT = "16px serif";

const kernWrites: (string | undefined)[] = [];

class FakeKernContext {
  private _font = FONT;
  private _fontKerning: string | undefined;

  set font(v: string) {
    this._font = v;
  }
  get font(): string {
    return this._font;
  }
  set fontKerning(v: string | undefined) {
    this._fontKerning = v;
    kernWrites.push(v);
  }
  get fontKerning(): string | undefined {
    return this._fontKerning;
  }
  measureText(text: string): { width: number } {
    const em = Number(/(\d+(?:\.\d+)?)px/.exec(this._font)?.[1] ?? 16);
    let width = 0;
    for (const ch of text) width += ch === " " || ch === "\t" ? em / 4 : em / 2;
    // Kerned "AV": the classic pair narrows by 2px — proof the mode reached
    // the measurement call, not just the context property.
    if (this._fontKerning === "normal" && text.includes("AV")) width -= 2;
    return { width };
  }
}

if (typeof OffscreenCanvas === "undefined") {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- the minimal duck-typed surface pretext touches
  (globalThis as any).OffscreenCanvas = class {
    getContext(): unknown {
      return new FakeKernContext();
    }
  };
}

describe.sequential("canvas font kerning", () => {
  beforeEach(() => {
    clearMeasurementCaches();
    kernWrites.length = 0;
  });

  it("probes the engine for fontKerning support", () => {
    expect(fontKerningSupported()).toBe(true);
  });

  it("measures with kerning when the preparation asks for it", () => {
    const kerned = prepareWithSegments("AV", FONT, { fontKerning: true });
    // The fake narrows "AV" by 2px only while fontKerning === "normal".
    expect(measureNaturalWidth(kerned)).toBe(14);
    expect(kernWrites).toContain("normal");
  });

  it("restores the default mode when kerning is not requested", () => {
    prepareWithSegments("AV", FONT, { fontKerning: true });
    const plain = prepareWithSegments("AV", FONT);
    expect(measureNaturalWidth(plain)).toBe(16);
    expect(kernWrites.at(-1)).toBe("auto");
  });

  it("keeps kerned and default segment caches apart", () => {
    const plain = getFontMeasurementState(FONT, false, false);
    // Measure under the mode the state just applied, before switching.
    expect(getSegmentMetrics("AV", plain.cache).width).toBe(16);
    const kerned = getFontMeasurementState(FONT, false, false, true);
    expect(kerned.cache).not.toBe(plain.cache);
    expect(getSegmentMetrics("AV", kerned.cache).width).toBe(14);
  });
});
