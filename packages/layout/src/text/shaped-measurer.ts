import type { PrepareOptions } from "@docen/pretext";
import type { FontRef } from "@docen/shaping";
import { FontManager, createFontRefSync, isShapingWasmInitialized } from "@docen/shaping";

import type { FontMetrics, ResolvedFaceMetrics } from "../font";
import {
  isCjkCodePoint,
  isCjkText,
  setFaceMetricsResolver,
  setRegisteredFontMetrics,
} from "../font";
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

let globalShapingDefault = true;

/**
 * Flag controlling deterministic OpenType shaping across docen.
 *
 * Default is `true`: registered fonts shape (rustybuzz + fontations) and feed
 * the line breaker; families without a registered face fall back to the
 * canvas measurer per run, so a document never shaped by its host is
 * unaffected. `setShapingEnabled(false)` (or `DOCEN_SHAPING_DISABLED=1`)
 * rolls back to canvas; `DOCEN_SHAPING_ENABLED=1` forces shaping on.
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

/** Word's single-line ratio and the baseline share from a face's own tables:
 *  winAscent + winDescent + 2 × round(0.15 × (A + D)) over upem — the
 *  font-metrics-data formula, computed from the face itself instead of the
 *  DOM probe. */
function faceMetricsOf(fontRef: FontRef): ResolvedFaceMetrics {
  const m = fontRef.metrics;
  const upem = m.unitsPerEm || 1000;
  const winA = m.winAscent ?? Math.max(m.ascender, 0);
  const winD = m.winDescent ?? Math.max(-m.descender, 0);
  const sum = winA + winD;
  return {
    ratio: (sum + 2 * Math.round(0.15 * sum)) / upem,
    baselineShare: winA / upem,
  };
}

// Any face the shaping font manager holds resolves from its own tables — the
// resolver makes that true even for FontManager.registerActiveFont /
// resolveFont callers that never went through registerShapingFont. `font.ts`
// stays dependency-free; this module injects the bridge.
setFaceMetricsResolver((family) => {
  if (!isShapingWasmInitialized()) return undefined;
  const fontRef = getShapingFontManager().getActiveFont(family);
  return fontRef ? faceMetricsOf(fontRef) : undefined;
});

/**
 * Register font bytes for shaping in the shared manager. The WASM runtime
 * must be initialized first (`await initShapingWasm()`). `fontIndex` selects
 * a face inside a `.ttc`/`.otc` collection (0 = plain sfnt). `style` selects
 * the face's weight/slant slot: a bold/italic run first looks for its own
 * registered face and falls back to the family's regular face when none was
 * registered.
 */
export function registerShapingFont(
  family: string,
  fontData: Uint8Array,
  fontIndex = 0,
  style?: ShapingFontStyle,
): FontRef {
  const fontRef = createFontRefSync(fontData, { index: fontIndex });
  getShapingFontManager().registerActiveFont(shapingFaceKey(family, style), fontRef);
  setRegisteredFontMetrics(family, faceMetricsOf(fontRef));
  clearMissingFontWarning(family);
  return fontRef;
}

/** Weight/slant slot of a registered shaping face. */
export interface ShapingFontStyle {
  readonly bold?: boolean;
  readonly italic?: boolean;
}

/** The manager/fontMap key for a family+style face: the regular face keeps
 *  the bare family key (back-compatible with `registerActiveFont(name)`
 *  callers), styled faces get a `|b`/`|i`/`|bi` suffix. */
export function shapingFaceKey(family: string, style?: ShapingFontStyle): string {
  const base = family.trim().toLowerCase();
  const suffix = `${style?.bold ? "b" : ""}${style?.italic ? "i" : ""}`;
  return suffix ? `${base}|${suffix}` : base;
}

/** Families already warned about a missing face — one warning per family, so
 *  a paragraph of canvas-fallback runs does not flood the console. */
const warnedMissingFonts = new Set<string>();

/** Drop a family's warned flag (called on registration — a late
 *  `registerFont` after a warning must be able to warn again if it is
 *  replaced by an unregistered family). */
export function clearMissingFontWarning(family: string): void {
  warnedMissingFonts.delete(family.trim().toLowerCase());
}

/** Test/app helper: forget every missing-face warning. */
export function clearMissingFontWarnings(): void {
  warnedMissingFonts.clear();
}

/** Loud fallback: shaping is on, but the run's family has no registered face,
 *  so the measurer/painter degrade to canvas — the user-visible equivalent of
 *  "this document may not measure like Word". The `DOCEN_SHAPING_DISABLED=1`
 *  rollback is an explicit opt-out and stays silent. */
function warnMissingShapingFont(family: string): void {
  const key = family.trim().toLowerCase();
  if (!key || warnedMissingFonts.has(key)) return;
  warnedMissingFonts.add(key);
  console.warn(
    `[@docen/layout] shaping is enabled but no face is registered for "${family}" — ` +
      "falling back to canvas measurement. Register the production faces with " +
      "registerDefaultFonts() (or registerShapingFont) for deterministic metrics; " +
      "DOCEN_SHAPING_DISABLED=1 rolls back to canvas for the whole process.",
  );
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

  registerFont(family: string, fontRef: FontRef, style?: ShapingFontStyle): void {
    this.fontMap.set(shapingFaceKey(family, style), fontRef);
    setRegisteredFontMetrics(family, faceMetricsOf(fontRef));
    clearMissingFontWarning(family);
  }

  /** The face for a family+style: the exact bold/italic slot first, then the
   *  family's regular face (CSS-style synthetic fallback), then the shared
   *  manager. */
  getFont(family: string, style?: ShapingFontStyle): FontRef | undefined {
    const key = shapingFaceKey(family, style);
    const found = this.fontMap.get(key) ?? this.fontManager.getActiveFont(key);
    if (found) return found;
    if (style?.bold || style?.italic) {
      const base = family.trim().toLowerCase();
      return this.fontMap.get(base) ?? this.fontManager.getActiveFont(base);
    }
    return undefined;
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

  /** Feed the line breaker shaped advances: without this the packer measures
   *  every segment with canvas measureText even when shaping is enabled, so
   *  kern/liga widths never reach the wrap decisions. */
  override segmentMeasurer(style: LayoutTextStyle): ((segment: string) => number) | undefined {
    if (!this.enabled) return undefined;
    const family =
      typeof style.family === "string"
        ? style.family
        : (style.family.latin ?? style.family.eastAsia);
    if (!family || !this.getFont(family, { bold: style.bold, italic: style.italic })) {
      if (family) warnMissingShapingFont(family);
      return undefined;
    }
    // Natural advances only: pretext applies the item's letterSpacing and
    // widthScale itself, so shape the run with spacing zeroed (shapeRun
    // refuses styled runs — the painter cannot represent spacing yet).
    const shapeStyle: LayoutTextStyle = style.letterSpacingPx
      ? { ...style, letterSpacingPx: undefined }
      : style;
    const scale = characterScaleOf(shapeStyle);
    return (segment: string): number => {
      if (!segment) return 0;
      const run = this.shapeRun(segment, shapeStyle);
      if (run) return run.totalAdvancePx;
      // A segment the shaper cannot represent (mixed script, degenerate size)
      // falls back to the canvas advance, de-scaled to the raw contract.
      const raw = super.widthOf(segment, shapeStyle);
      return scale === 1 ? raw : raw / scale;
    };
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

    const fontRef = this.getFont(family, { bold: style.bold, italic: style.italic });
    if (!fontRef) {
      warnMissingShapingFont(family);
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
