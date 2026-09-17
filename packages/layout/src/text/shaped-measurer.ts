import type { PrepareOptions } from "@docen/pretext";
import { FontManager, type FontRef } from "@docen/shaping";

import type { FontMetrics } from "../font";
import type { LayoutTextStyle } from "../layout-doc";
import type { LaidOutGlyphRun } from "../layout-result";
import { TextMeasurer, familyOfSlot, vertAlignedSizePx } from "./measure";

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
    const cacheKey = `${text}|${family}|${sizePx}|${style.bold ? "b" : ""}|${style.italic ? "i" : ""}`;

    const cached = this.runCache.get(cacheKey);
    if (cached) return cached;

    const fontRef = this.getFont(family);
    if (!fontRef) {
      return undefined;
    }

    try {
      const shapingRes = fontRef.shape(text, {
        direction: "ltr",
      });

      const scale = sizePx / (fontRef.unitsPerEm || 1000);

      let currentX = 0;
      const glyphs = shapingRes.glyphs.map((g) => {
        const xAdvPx = g.xAdvance * scale;
        const xOffPx = g.xOffset * scale;
        const yOffPx = g.yOffset * scale;
        const xPx = currentX + xOffPx;
        const yPx = yOffPx;
        currentX += xAdvPx;

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
