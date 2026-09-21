import type { PaintNodeLike } from "@docen/core";

import { NodeImageCache, type PdfImageCache } from "./node-image";
import {
  composeMatrix,
  invertMatrix,
  type PdfMatrix,
  type PdfSceneNode,
  type PdfScenePage,
  type PdfSceneShapeNode,
  type PdfSceneTextNode,
  type PdfSceneTextRow,
} from "./pdf-scene";

const SHAPE_TAGS = new Set(["Rect", "Path", "Line", "Ellipse", "Polygon", "Star", "Pen"]);

interface NodeTextRow {
  x?: number;
  y?: number;
  width?: number;
  text?: string;
  data?: { char?: string }[];
  words?: { data?: { char?: string }[] }[];
}

function rowTexts(elementText: string, rows: readonly NodeTextRow[]): string[] {
  const out: string[] = [];
  let cursor = 0;
  for (const row of rows) {
    if (typeof row.text === "string" && row.text.length > 0) {
      out.push(row.text);
      cursor += row.text.length;
      continue;
    }
    const chars = row.data ?? [];
    let text = "";
    let i = 0;
    while (i < chars.length && cursor < elementText.length) {
      const source = elementText[cursor]!;
      if (source === " ") {
        text += " ";
        cursor++;
        continue;
      }
      if (source === "\n") {
        cursor++;
        break;
      }
      const char = chars[i]?.char;
      if (typeof char === "string") text += char;
      i++;
      cursor++;
    }
    while (cursor < elementText.length && elementText[cursor] === " ") {
      text += " ";
      cursor++;
    }
    out.push(text);
  }
  return out;
}

export function paintColor(value: unknown): string | undefined {
  if (typeof value === "string") return value;
  if (Array.isArray(value)) {
    for (const entry of value) {
      const color = paintColor(entry);
      if (color) return color;
    }
    return undefined;
  }
  if (value && typeof value === "object") {
    const paint = value as { type?: string; stops?: { color?: string }[]; color?: string };
    if (typeof paint.color === "string") return paint.color;
    if (paint.type === "linear" || paint.type === "radial") return paint.stops?.[0]?.color;
  }
  return undefined;
}

function numberOr(value: unknown, fallback?: number): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function letterSpacingOf(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value && typeof value === "object") {
    const obj = value as { type?: string; value?: number };
    if (typeof obj.value === "number" && Number.isFinite(obj.value)) {
      return obj.value;
    }
  }
  return undefined;
}

function strokeJoinOf(value: unknown): "miter" | "round" | "bevel" | undefined {
  return value === "round" || value === "bevel" || value === "miter" ? value : undefined;
}

function strokeCapOf(value: unknown): "butt" | "round" | "square" | undefined {
  if (value === "round" || value === "square") return value;
  if (value === "butt" || value === "none") return "butt";
  return undefined;
}

/**
 * Serializes a pure PaintNodeLike scene tree (e.g. built by nodeKit or Leafer)
 * into a pure PdfScenePage vector representation for PDF export without DOM.
 */
export async function serializeNodeScene(
  root: PaintNodeLike,
  pageWidthPx: number,
  pageHeightPx: number,
  imageCache: PdfImageCache = new NodeImageCache(),
): Promise<PdfScenePage> {
  const nodes: PdfSceneNode[] = [];
  const skipped: Record<string, number> = {};

  const normalize = invertMatrix(root.worldTransform as PdfMatrix);
  const matrixOf = (el: PaintNodeLike): PdfMatrix =>
    composeMatrix(normalize, el.worldTransform as unknown as PdfMatrix);
  const childrenOf = (el: PaintNodeLike): PaintNodeLike[] =>
    (el as unknown as { children?: PaintNodeLike[] }).children ?? [];

  const walk = async (parent: PaintNodeLike, sink: PdfSceneNode[]): Promise<void> => {
    for (const el of childrenOf(parent)) {
      if (el.visible === false) continue;
      const worldOpacity = el.worldOpacity ?? 1;
      if (worldOpacity <= 0) continue;
      const tag = el.tag;
      const matrix = matrixOf(el);
      const blendMode = el.blendMode;
      const base = {
        ...(worldOpacity < 1 ? { opacity: worldOpacity } : {}),
        ...(blendMode && blendMode !== "pass-through" ? { blendMode } : {}),
      };

      const overflow = (el as unknown as { overflow?: string }).overflow;
      if (tag === "Group" || (tag === "Box" && overflow !== "hide")) {
        await walk(el, sink);
        continue;
      }
      if (tag === "Box" || tag === "Frame" || tag === "Flow") {
        const children: PdfSceneNode[] = [];
        await walk(el, children);
        if (children.length === 0) continue;
        const clipPath = typeof el.getPathString === "function" ? el.getPathString(true, true) : "";
        sink.push({
          type: "group",
          ...base,
          ...(clipPath ? { clip: { path: clipPath, matrix } } : {}),
          children,
        });
        continue;
      }
      if (SHAPE_TAGS.has(tag)) {
        const path = typeof el.getPathString === "function" ? el.getPathString(true, true) : "";
        if (!path) continue;
        const withEl = el as unknown as {
          fill?: unknown;
          stroke?: unknown;
          strokeWidth?: unknown;
          dashPattern?: unknown;
          strokeCap?: unknown;
          strokeJoin?: unknown;
          windingRule?: string;
          data?: { docenGlyph?: boolean };
        };
        const fill = paintColor(withEl.fill);
        const stroke = paintColor(withEl.stroke);
        const strokeWidth = numberOr(withEl.strokeWidth);
        const cap = strokeCapOf(withEl.strokeCap);
        const join = strokeJoinOf(withEl.strokeJoin);
        const shape: PdfSceneShapeNode = {
          type: "shape",
          ...base,
          matrix,
          path,
          ...(fill ? { fill } : {}),
          ...(stroke ? { stroke } : {}),
          ...(stroke && strokeWidth ? { strokeWidth } : {}),
          ...(Array.isArray(withEl.dashPattern) ? { dash: withEl.dashPattern as number[] } : {}),
          ...(cap ? { cap } : {}),
          ...(join ? { join } : {}),
          ...(withEl.windingRule === "evenodd" ? { winding: "evenodd" as const } : {}),
          ...(withEl.data?.docenGlyph ? { glyph: true } : {}),
        };
        sink.push(shape);
        continue;
      }
      if (tag === "Text") {
        const textEl = el as unknown as {
          text?: string | number;
          textDrawData?: { rows?: NodeTextRow[] };
          fontSize?: number;
          fontFamily?: string;
          fontWeight?: number | string;
          italic?: boolean;
          letterSpacing?: unknown;
          fill?: unknown;
          stroke?: unknown;
          strokeWidth?: unknown;
        };
        const rows: PdfSceneTextRow[] = [];
        const sourceText =
          typeof textEl.text === "string" ? textEl.text : String(textEl.text ?? "");
        const drawRows = textEl.textDrawData?.rows ?? [];
        const rowStrings = rowTexts(sourceText, drawRows);
        for (const [rowIndex, row] of drawRows.entries()) {
          const text = rowStrings[rowIndex] ?? "";
          if (!text.trim()) continue;
          rows.push({
            x: numberOr(row.x, 0)!,
            y: numberOr(row.y, 0)!,
            width: numberOr(row.width, 0)!,
            text,
          });
        }
        if (rows.length === 0) continue;
        const fill = paintColor(textEl.fill);
        const stroke = paintColor(textEl.stroke);
        const strokeWidth = numberOr(textEl.strokeWidth);
        const weight = textEl.fontWeight;
        const bold =
          typeof weight === "number" ? weight >= 600 : weight === "bold" || weight === "600";
        const letterSpacing = letterSpacingOf(textEl.letterSpacing);
        const text: PdfSceneTextNode = {
          type: "text",
          ...base,
          matrix,
          rows,
          fontSize: numberOr(textEl.fontSize, 12)!,
          ...(textEl.fontFamily ? { fontFamily: textEl.fontFamily } : {}),
          ...(bold ? { bold: true } : {}),
          ...(textEl.italic ? { italic: true } : {}),
          ...(letterSpacing ? { letterSpacing } : {}),
          ...(fill ? { fill } : {}),
          ...(stroke ? { stroke } : {}),
          ...(stroke && strokeWidth ? { strokeWidth } : {}),
        };
        sink.push(text);
        continue;
      }
      if (tag === "Image") {
        const imageEl = el as unknown as {
          url?: string;
          ready?: boolean;
          width?: number;
          height?: number;
          image?: { width?: number; height?: number };
        };
        if (!imageEl.url || imageEl.ready === false) continue;
        const image = await imageCache.get(
          imageEl.url,
          imageEl.image?.width,
          imageEl.image?.height,
        );
        if (!image) continue;
        sink.push({
          type: "image",
          ...base,
          matrix,
          image,
          width: numberOr(imageEl.width, image.width)!,
          height: numberOr(imageEl.height, image.height)!,
        });
        continue;
      }

      if (childrenOf(el).length > 0) await walk(el, sink);
      else skipped[tag] = (skipped[tag] ?? 0) + 1;
    }
  };

  await walk(root, nodes);
  return {
    width: pageWidthPx,
    height: pageHeightPx,
    nodes,
    ...(Object.keys(skipped).length > 0 ? { skipped } : {}),
  };
}
