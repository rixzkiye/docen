import type { PrepareOptions } from "@docen/pretext";
import { FontManager, createFontRefSync, type FontRef } from "@docen/shaping";

import type { FontMetrics } from "../font";
import { isCjkCodePoint, isCjkText, setRegisteredFontRatio } from "../font";
import type { LayoutTextStyle } from "../layout-doc";
import type { LaidOutGlyphRun } from "../layout-result";
import {
  TextMeasurer,
  characterScaleOf,
  familyOfSlot,
  kerningActive,
  vertAlignedSizePx,
} from "./measure";

const GRAPHEME_SEGMENTER = new Intl.Segmenter(undefined, { granularity: "grapheme" });

function countGraphemes(text: string): number {
  let count = 0;
  for (const _ of GRAPHEME_SEGMENTER.segment(text)) count++;
  return count;
}

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

let globalShapingDefault = false;

/**
 * Flag controlling deterministic OpenType shaping across docen.
 *
 * Default is `false` (canvas measureText — R6.4 audit: the laid-out glyph
 * pipeline had no registered fonts, so a "default" shaped measurer was a
 * no-op facade; the default returns to canvas until document-level parity,
 * per-keystroke and cross-environment determinism gates are green).
 * `setShapingEnabled(true)` opts in; `DOCEN_SHAPING_ENABLED=1` forces it on
 * and `DOCEN_SHAPING_DISABLED=1` forces it off regardless of the flag.
 */
export function setShapingEnabled(enabled: boolean): void {
  globalShapingDefault = enabled;
}

export function isShapingEnabled(): boolean {
  if (typeof process !== "undefined") {
    if (process.env?.DOCEN_SHAPING_DISABLED === "1") return false;
    if (process.env?.DOCEN_SHAPING_ENABLED === "1") return true;
  }
  return globalShapingDefault;
}

/** The process-wide font manager every ShapedMeasurer shares unless it is
 *  given its own — so a font registered once (editor host, tests) is visible
 *  to the body, furniture and drawing measurers alike. */
let defaultFontManager: FontManager | undefined;

export function getShapingFontManager(): FontManager {
  defaultFontManager ??= new FontManager({ enableOpfs: false });
  return defaultFontManager;
}

/** Word's single-line ratio from a face's own OS/2 metrics: winAscent +
 *  winDescent + 2 × round(0.15 × (A + D)) over upem — the font-metrics-data
 *  formula, computed from the face itself instead of the DOM probe. */
function wordLineRatioOf(fontRef: FontRef): number {
  const m = fontRef.metrics;
  const upem = m.unitsPerEm || 1000;
  const winA = m.winAscent ?? m.ascender;
  const winD = m.winDescent ?? -m.descender;
  const sum = winA + winD;
  return (sum + 2 * Math.round(0.15 * sum)) / upem;
}

/**
 * Register font bytes for shaping in the shared manager. The WASM runtime
 * must be initialized first (`await initShapingWasm()`).
 */
export function registerShapingFont(family: string, fontData: Uint8Array): FontRef {
  const fontRef = createFontRefSync(fontData);
  getShapingFontManager().registerActiveFont(family, fontRef);
  setRegisteredFontRatio(family, wordLineRatioOf(fontRef));
  return fontRef;
}

export interface ShapedMeasurerOptions {
  readonly fontManager?: FontManager;
  readonly optIn?: boolean;
  readonly enabled?: boolean;
}

/**
 * Factory creating a ShapedMeasurer when shaping is explicitly opted in
 * (`enabled`/`optIn`) or the global flag is set, else the standard
 * canvas-backed TextMeasurer.
 */
export function createMeasurer(
  metrics: FontMetrics,
  options?: ShapedMeasurerOptions,
): TextMeasurer {
  if (options?.enabled === false || options?.optIn === false) {
    return new TextMeasurer(metrics);
  }
  if (
    options?.enabled === true ||
    options?.optIn === true ||
    options?.fontManager !== undefined ||
    isShapingEnabled()
  ) {
    return new ShapedMeasurer(metrics, { ...options, enabled: true });
  }
  return new TextMeasurer(metrics);
}

/**
 * TextMeasurer backed by @docen/shaping (rustybuzz + fontations) with
 * content-addressed caching, exact font design unit advances, and glyph-run generation.
 */
export class ShapedMeasurer extends TextMeasurer {
  readonly fontManager: FontManager;
  private readonly forcedEnabled?: boolean;
  private readonly runCache = new Map<string, LaidOutGlyphRun>();
  private readonly fontMap = new Map<string, FontRef>();

  constructor(metrics: FontMetrics, options?: ShapedMeasurerOptions) {
    super(metrics);
    this.fontManager = options?.fontManager ?? getShapingFontManager();
    this.forcedEnabled = options?.enabled ?? options?.optIn;
  }

  get enabled(): boolean {
    if (this.forcedEnabled !== undefined) {
      return this.forcedEnabled;
    }
    return isShapingEnabled();
  }

  registerFont(family: string, fontRef: FontRef): void {
    this.fontMap.set(family.toLowerCase(), fontRef);
    setRegisteredFontRatio(family, wordLineRatioOf(fontRef));
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
      // Match pretext's width math: letter spacing adds between graphemes
      // (never after the last) and the w:w scale stretches the sum.
      const spacing = style.letterSpacingPx ?? 0;
      const spacingSum = spacing === 0 ? 0 : Math.max(0, countGraphemes(text) - 1) * spacing;
      return (glyphRun.totalAdvancePx + spacingSum) * characterScaleOf(style);
    }

    return super.widthOf(text, style, whiteSpace);
  }

  /**
   * The shaped run for a laid-out run of text, when the measurer can shape it.
   * `undefined` = the canvas measurer (or shaping off / no registered font) and
   * the painter should keep using fillText.
   */
  override glyphRunOf(text: string, style: LayoutTextStyle): LaidOutGlyphRun | undefined {
    if (!this.enabled || !text) return undefined;
    return this.shapeRun(text, style);
  }

  /** Shape a run of text and produce a LaidOutGlyphRun. */
  shapeRun(text: string, style: LayoutTextStyle): LaidOutGlyphRun | undefined {
    if (!this.enabled || !text) return undefined;
    // The glyph painter places natural advances (plus justification stretch);
    // letter spacing is not representable yet, so those runs stay canvas.
    if (style.letterSpacingPx) return undefined;

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
    // A slot family resolves by script: a mixed Latin+CJK run cannot be
    // shaped faithfully with one face, so it stays canvas (the painter then
    // paints the same mixed run with fillText).
    if (typeof style.family !== "string") {
      let hasCjk = false;
      let hasNonCjk = false;
      for (const ch of text) {
        if (isCjkCodePoint(ch)) hasCjk = true;
        else hasNonCjk = true;
        if (hasCjk && hasNonCjk) return undefined;
      }
    }
    const family = familyOfSlot(style.family, isCjkText(text));
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
      const isRtl = direction === "rtl";
      const totalAdvancePx = shapingRes.totalAdvance * scale;

      // HarfBuzz/rustybuzz returns RTL glyphs in visual order; the pen walks
      // right-to-left so the first output glyph lands at the run's right edge.
      let pen = 0;
      let currentY = 0;
      const glyphs = shapingRes.glyphs.map((g) => {
        const xAdvPx = g.xAdvance * scale;
        const advPx = isVertical ? Math.abs(g.yAdvance) * scale : xAdvPx;
        const xOffPx = g.xOffset * scale;
        const yOffPx = g.yOffset * scale;
        const xPx = isVertical ? xOffPx : (isRtl ? totalAdvancePx - pen - advPx : pen) + xOffPx;
        const yPx = isVertical ? currentY + yOffPx : yOffPx;
        pen += advPx;
        if (isVertical) currentY += advPx;

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

      const run: LaidOutGlyphRun = {
        fontId: fontRef.id,
        fontName: family,
        fontSizePx: sizePx,
        unitsPerEm: metrics.unitsPerEm || 1000,
        direction,
        script: style.script,
        language: style.language,
        ...(variations.length > 0 ? { variations } : {}),
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
