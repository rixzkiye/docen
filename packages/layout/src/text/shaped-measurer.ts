import type { PrepareOptions } from "@docen/pretext";
import { FontManager, type FontRef } from "@docen/shaping";

import type { FontMetrics } from "../font";
import type { LayoutTextStyle } from "../layout-doc";
import type { LaidOutGlyphRun } from "../layout-result";
import { TextMeasurer, familyOfSlot, kerningActive, vertAlignedSizePx } from "./measure";

export function resolveOpenTypeFeatures(style: LayoutTextStyle): { tag: string; value: number }[] {
  const features: { tag: string; value: number }[] = [];

  if (style.kernPt !== undefined) {
    if (style.kernPt === 0) {
      features.push({ tag: "kern", value: 0 });
    } else {
      const active = kerningActive(style);
      features.push({ tag: "kern", value: active ? 1 : 0 });
    }
  }

  if (style.caps === "small") {
    features.push({ tag: "smcp", value: 1 });
  } else if (style.caps === "all") {
    features.push({ tag: "c2sc", value: 1 });
  }

  if (style.ligatures === "none") {
    features.push({ tag: "liga", value: 0 });
    features.push({ tag: "clig", value: 0 });
    features.push({ tag: "calt", value: 0 });
  } else if (style.ligatures === "all") {
    features.push({ tag: "liga", value: 1 });
    features.push({ tag: "dlig", value: 1 });
    features.push({ tag: "hlig", value: 1 });
    features.push({ tag: "calt", value: 1 });
  } else if (style.ligatures === "discretionary") {
    features.push({ tag: "dlig", value: 1 });
  } else if (style.ligatures === "historical") {
    features.push({ tag: "hlig", value: 1 });
  }

  if (style.numForm === "oldStyle") {
    features.push({ tag: "onum", value: 1 });
  } else if (style.numForm === "lining") {
    features.push({ tag: "lnum", value: 1 });
  }

  if (style.numSpacing === "tabular") {
    features.push({ tag: "tnum", value: 1 });
  } else if (style.numSpacing === "proportional") {
    features.push({ tag: "pnum", value: 1 });
  }

  if (style.stylisticSet && style.stylisticSet >= 1 && style.stylisticSet <= 20) {
    const tag = `ss${String(style.stylisticSet).padStart(2, "0")}`;
    features.push({ tag, value: 1 });
  }

  if (style.fontFeatures && style.fontFeatures.length > 0) {
    for (const f of style.fontFeatures) {
      features.push({ tag: f.tag, value: f.value ?? 1 });
    }
  }

  return features;
}

export function resolveFontVariations(
  style: LayoutTextStyle,
  fontRef: FontRef,
): { tag: string; value: number }[] {
  const variations: { tag: string; value: number }[] = [];

  const axes = fontRef.getAxes();
  if (axes && axes.length > 0) {
    const hasWght = axes.some((a) => a.tag === "wght");
    if (hasWght) {
      if (style.fontWeight !== undefined) {
        variations.push({ tag: "wght", value: style.fontWeight });
      } else if (style.bold) {
        variations.push({ tag: "wght", value: 700 });
      }
    }
  }

  if (style.fontVariations && style.fontVariations.length > 0) {
    for (const v of style.fontVariations) {
      variations.push({ tag: v.tag, value: v.value });
    }
  }

  return variations;
}

let globalShapingOptIn = false;

/**
 * Opt-in flag controlling deterministic OpenType shaping across docen.
 */
export function setShapingEnabled(enabled: boolean): void {
  globalShapingOptIn = enabled;
}

export function isShapingEnabled(): boolean {
  return globalShapingOptIn;
}

export interface ShapedMeasurerOptions {
  readonly fontManager?: FontManager;
  readonly optIn?: boolean;
}

/**
 * TextMeasurer backed by @docen/shaping (rustybuzz + fontations) with
 * content-addressed caching, exact font design unit advances, and glyph-run generation.
 */
export class ShapedMeasurer extends TextMeasurer {
  readonly fontManager: FontManager;
  private readonly optIn: boolean;
  private readonly runCache = new Map<string, LaidOutGlyphRun>();
  private readonly fontMap = new Map<string, FontRef>();

  constructor(metrics: FontMetrics, options?: ShapedMeasurerOptions) {
    super(metrics);
    this.fontManager = options?.fontManager ?? new FontManager({ enableOpfs: false });
    this.optIn = options?.optIn ?? false;
  }

  get enabled(): boolean {
    return this.optIn || globalShapingOptIn;
  }

  registerFont(family: string, fontRef: FontRef): void {
    this.fontMap.set(family.toLowerCase(), fontRef);
  }

  getFont(family: string): FontRef | undefined {
    return this.fontMap.get(family.toLowerCase()) ?? this.fontManager.getActiveFont(family);
  }

  override widthOf(
    text: string,
    style: LayoutTextStyle,
    whiteSpace?: PrepareOptions["whiteSpace"],
  ): number {
    if (!this.enabled || !text) {
      return super.widthOf(text, style, whiteSpace);
    }

    const glyphRun = this.shapeRun(text, style);
    if (glyphRun) {
      return glyphRun.totalAdvancePx;
    }

    return super.widthOf(text, style, whiteSpace);
  }

  /**
   * Shape a run of text and produce a LaidOutGlyphRun.
   */
  shapeRun(text: string, style: LayoutTextStyle): LaidOutGlyphRun | undefined {
    if (!this.enabled || !text) return undefined;

    const rawSize =
      style.sizePx ??
      ((style as unknown as { sizePt?: number }).sizePt
        ? (style as unknown as { sizePt: number }).sizePt * (96 / 72)
        : undefined);
    if (rawSize === undefined || isNaN(rawSize) || rawSize <= 0) {
      return undefined;
    }

    const effectiveStyle: LayoutTextStyle =
      style.sizePx !== undefined ? style : { ...style, sizePx: rawSize };
    const sizePx = vertAlignedSizePx(effectiveStyle);
    const family = familyOfSlot(style.family, false);
    const direction = style.vertical ? "ttb" : (style.direction ?? "auto");

    const fontRef = this.getFont(family);
    if (!fontRef) {
      return undefined;
    }

    const features = resolveOpenTypeFeatures(style);
    const variations = resolveFontVariations(style, fontRef);

    const featStr = features.map((f) => `${f.tag}:${f.value}`).join(",");
    const varStr = variations.map((v) => `${v.tag}:${v.value}`).join(",");
    const cacheKey = `${text}|${family}|${sizePx}|${style.bold ? "b" : ""}|${style.italic ? "i" : ""}|${direction}|${style.script ?? ""}|${style.language ?? ""}|${featStr}|${varStr}`;

    const cached = this.runCache.get(cacheKey);
    if (cached) return cached;

    try {
      const shapingRes = fontRef.shape(text, {
        direction,
        script: style.script,
        language: style.language,
        features: features.length > 0 ? features : undefined,
        variations: variations.length > 0 ? variations : undefined,
      });

      const metrics = variations.length > 0 ? fontRef.getMetrics(variations) : fontRef.metrics;
      const scale = sizePx / (metrics.unitsPerEm || 1000);
      const isVertical = direction === "ttb";

      let currentX = 0;
      let currentY = 0;
      const glyphs = shapingRes.glyphs.map((g) => {
        const xAdvPx = g.xAdvance * scale;
        const yAdvPx = Math.abs(g.yAdvance) * scale;
        const xOffPx = g.xOffset * scale;
        const yOffPx = g.yOffset * scale;
        const xPx = isVertical ? xOffPx : currentX + xOffPx;
        const yPx = isVertical ? currentY + yOffPx : yOffPx;
        currentX += xAdvPx;
        currentY += yAdvPx;

        return {
          glyphId: g.glyphId,
          cluster: g.cluster,
          xAdvance: g.xAdvance,
          yAdvance: g.yAdvance,
          xOffset: g.xOffset,
          yOffset: g.yOffset,
          xPx,
          yPx,
        };
      });

      const totalAdvancePx = shapingRes.totalAdvance * scale;
      const run: LaidOutGlyphRun = {
        fontId: fontRef.id,
        fontName: family,
        fontSizePx: sizePx,
        direction,
        script: style.script,
        language: style.language,
        glyphs,
        totalAdvancePx,
      };

      if (this.runCache.size >= 500) {
        const oldest = this.runCache.keys().next().value;
        if (oldest !== undefined) this.runCache.delete(oldest);
      }
      this.runCache.set(cacheKey, run);

      return run;
    } catch {
      return undefined;
    }
  }

  override clearCache(): void {
    super.clearCache();
    this.runCache.clear();
  }
}
