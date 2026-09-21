/**
 * Deterministic width measurement for painted "chrome" text — chart labels
 * and formatting-mark labels. These paint without an explicit font family, so
 * the PDF falls back to base-14 Helvetica; Liberation Sans is its
 * metric-compatible registered stand-in. Measuring with that face (the
 * shaping pipeline) gives the same advance in the browser canvas and in the
 * server-side Node render, instead of the platform font the canvas would
 * otherwise substitute — the scene serializers read the painted row geometry,
 * so a platform-dependent width here leaks into the PDF.
 */

import {
  browserFontMetrics,
  createMeasurer,
  getShapingFontManager,
  type TextMeasurer,
} from "@docen/layout";
import { isShapingWasmInitialized } from "@docen/shaping";

/** The registered face used to measure chrome text (metric-compatible with
 *  the base-14 Helvetica the PDF draws it with). */
const CHROME_FAMILY = "Arial";

let measurer: TextMeasurer | undefined;
let canvasCtx: CanvasRenderingContext2D | null | undefined;

/** The painted width of `text` at `px`, identical in the browser and Node
 *  when the shaping pipeline knows the stand-in face; falls back to the
 *  browser canvas (platform metrics) and finally to a per-character estimate
 *  for headless runs without registered fonts. */
export function measureChromeText(text: string, px: number): number {
  if (!text) return 0;
  if (
    isShapingWasmInitialized() &&
    getShapingFontManager().getActiveFont(CHROME_FAMILY.toLowerCase())
  ) {
    measurer ??= createMeasurer(browserFontMetrics);
    const run = measurer.glyphRunOf(text, { family: CHROME_FAMILY, sizePx: px });
    if (run) return run.totalAdvancePx;
  }
  canvasCtx ??=
    typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
  if (canvasCtx) {
    canvasCtx.font = `${px}px ${CHROME_FAMILY}, sans-serif`;
    return canvasCtx.measureText(text).width;
  }
  return text.length * px;
}
