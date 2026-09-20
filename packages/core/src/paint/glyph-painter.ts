import type { LaidOutGlyphRun } from "@docen/layout";
import { getShapingBackend, hasShapingBackend, type PathCommand } from "@docen/shaping";
import { Path, type IGroup } from "leafer-ui";

/**
 * Converts unscaled OpenType glyph outline commands (with Y pointing up)
 * into standard SVG path data (with Y pointing down).
 */
export function commandsToSvgPath(commands: readonly PathCommand[]): string {
  if (!commands || commands.length === 0) return "";
  const parts: string[] = [];

  for (const cmd of commands) {
    switch (cmd.type) {
      case "M":
        parts.push(`M ${cmd.x} ${-cmd.y}`);
        break;
      case "L":
        parts.push(`L ${cmd.x} ${-cmd.y}`);
        break;
      case "Q":
        parts.push(`Q ${cmd.cx} ${-cmd.cy} ${cmd.x} ${-cmd.y}`);
        break;
      case "C":
        parts.push(`C ${cmd.cx1} ${-cmd.cy1} ${cmd.cx2} ${-cmd.cy2} ${cmd.x} ${-cmd.y}`);
        break;
      case "Z":
        parts.push("Z");
        break;
    }
  }

  return parts.join(" ");
}

/**
 * Content-addressed LRU cache for unscaled glyph outlines as SVG path strings.
 */
export class GlyphOutlineCache {
  private readonly cache = new Map<string, string>();
  private readonly maxSize: number;

  constructor(maxSize = 4000) {
    this.maxSize = maxSize;
  }

  getGlyphPath(
    fontId: number,
    glyphId: number,
    variations?: readonly { readonly tag: string; readonly value: number }[],
  ): string | undefined {
    const varKey = variations?.length ? variations.map((v) => `${v.tag}=${v.value}`).join(",") : "";
    const key = varKey ? `${fontId}:${glyphId}@${varKey}` : `${fontId}:${glyphId}`;
    const cached = this.cache.get(key);
    if (cached !== undefined) {
      return cached;
    }

    if (!hasShapingBackend("rustybuzz")) {
      return undefined;
    }

    try {
      const backend = getShapingBackend("rustybuzz");
      if (!backend.getGlyphOutline) {
        return undefined;
      }
      const commands = backend.getGlyphOutline(fontId, glyphId, variations);
      const pathStr = commandsToSvgPath(commands);

      if (this.cache.size >= this.maxSize) {
        const oldest = this.cache.keys().next().value;
        if (oldest !== undefined) this.cache.delete(oldest);
      }
      this.cache.set(key, pathStr);
      return pathStr;
    } catch {
      return undefined;
    }
  }

  setGlyphPath(fontId: number, glyphId: number, path: string): void {
    const key = `${fontId}:${glyphId}`;
    if (this.cache.size >= this.maxSize) {
      const oldest = this.cache.keys().next().value;
      if (oldest !== undefined) this.cache.delete(oldest);
    }
    this.cache.set(key, path);
  }

  clear(): void {
    this.cache.clear();
  }
}

export const defaultGlyphOutlineCache = new GlyphOutlineCache();

export interface PaintGlyphRunOptions {
  x: number;
  y: number; // Baseline Y
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  /** One shadow or an effect stack (shadow + bevel edges) — Leafer paints the
   *  array in order, so a bevel's light and dark edges both render. */
  shadow?:
    | { x: number; y: number; blur?: number; color: string }
    | { x: number; y: number; blur?: number; color: string }[];
  scaleX?: number;
  opacity?: number;
  outlineCache?: GlyphOutlineCache;
  /** Overrides the run's own unitsPerEm (fonts default to 1000 only when no
   *  source records it — Calibri/Arial are 2048). */
  unitsPerEm?: number;
  /** Extra x-advance stretch (justification / CJK advance compression): both
   *  the glyph positions and the x scale multiply by it. */
  advanceScale?: number;
}

/**
 * Paints a laid-out OpenType glyph run as high-fidelity vector paths into the scene tree.
 */
export function paintGlyphRun(
  tree: IGroup,
  glyphRun: LaidOutGlyphRun,
  options: PaintGlyphRunOptions,
): boolean {
  if (!glyphRun.glyphs || glyphRun.glyphs.length === 0) {
    return false;
  }

  const fontId = glyphRun.fontId ?? 1;
  const cache = options.outlineCache ?? defaultGlyphOutlineCache;
  const unitsPerEm = options.unitsPerEm ?? glyphRun.unitsPerEm ?? 1000;
  const scale = glyphRun.fontSizePx / unitsPerEm;
  const advanceScale = options.advanceScale ?? 1;
  const scaleX = (options.scaleX ?? 1) * advanceScale * scale;
  const scaleY = scale;

  let paintedAny = false;

  for (const glyph of glyphRun.glyphs) {
    const pathStr = cache.getGlyphPath(fontId, glyph.glyphId, glyphRun.variations);
    if (!pathStr) continue;

    const glyphEl = new Path({
      x: options.x + glyph.xPx * advanceScale,
      y: options.y + glyph.yPx,
      scaleX,
      scaleY,
      path: pathStr,
      fill: options.fill ?? "#1b1b1b",
      stroke: options.stroke,
      strokeWidth: options.strokeWidth,
      shadow: options.shadow,
      opacity: options.opacity,
      // Marks glyph outlines for the PDF exporter: the vector export emits
      // them as paths (or drops them in embedded-font text measurement mode,
      // where the invisible text layer renders visibly instead).
      data: { docenGlyph: true },
      hittable: false,
    });

    tree.add(glyphEl);
    paintedAny = true;
  }

  return paintedAny;
}
