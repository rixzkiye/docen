// Live LeaferJS scene → serializable PDF scene (see pdf-scene.ts).
//
// The canvas is already a vector scene graph, so the PDF exporter reuses it
// instead of re-deriving layout: this module walks one page's painted tree
// (`app.tree`), normalizes every element's world transform into page space,
// and flattens the elements into the plain node contract the pure content
// writer consumes. Groups stay groups only when they clip (a Leafer `Box`
// with overflow hidden); everything else flattens to shapes/images/text with
// absolute matrices, which keeps the content writer free of transform-stack
// bookkeeping.
//
// Kept deliberately separate from stage.ts (a parallel lane edits it): the
// stage exposes one `sceneSnapshots()` method that calls in here, and nothing
// else in the stage touches this path.

import type { App, IUI } from "leafer-ui";

import {
  composeMatrix,
  invertMatrix,
  type PdfMatrix,
  type PdfSceneImageData,
  type PdfSceneNode,
  type PdfScenePage,
  type PdfSceneShapeNode,
  type PdfSceneTextNode,
  type PdfSceneTextRow,
} from "../pdf-scene";

/** A shape tag — the Leafer elements `getPathString(true, true)` describes as
 *  a filled/stroked path in local coordinates. */
const SHAPE_TAGS = new Set(["Rect", "Path", "Line", "Ellipse", "Polygon", "Star", "Pen"]);

/** One text row from Leafer's own layout. Char-mode rows (letter spacing /
 *  wrapping) carry `words`, not `text`; the letters are reassembled so the
 *  PDF row still reads as the laid string. */
interface LeaferTextRow {
  x?: number;
  y?: number;
  width?: number;
  text?: string;
  /** Char-mode rows (wrapping, letter spacing, explicit box) carry the run's
   *  glyphs here; plain rows carry `text`. */
  data?: { char?: string }[];
  words?: { data?: { char?: string }[] }[];
}

function rowText(row: LeaferTextRow): string {
  if (typeof row.text === "string" && row.text.length > 0) return row.text;
  let text = "";
  if (row.data) for (const char of row.data) if (typeof char.char === "string") text += char.char;
  if (text) return text;
  if (!row.words) return "";
  for (const word of row.words) {
    if (!word.data) continue;
    for (const char of word.data) if (typeof char.char === "string") text += char.char;
  }
  return text;
}

/** Rebuild one element's row strings. Plain rows carry their own `text`; in
 *  char mode Leafer strips spaces from `data`, so the element's source string
 *  is walked in parallel to re-insert them (spaces are gaps in char mode, and
 *  a PDF Tj must carry them explicitly — the row's glyph positions are the
 *  face's own advances, see the /W widths the exporter emits). */
function rowTexts(elementText: string, rows: readonly LeaferTextRow[]): string[] {
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
    // Trailing source spaces belong to this row (Leafer trims them from the
    // glyph run, but they are part of the row's text).
    while (cursor < elementText.length && elementText[cursor] === " ") {
      text += " ";
      cursor++;
    }
    out.push(text);
  }
  return out;
}

/** An FNV-1a hash of a byte run — image dedupe keys. */
function hashBytes(data: Uint8Array): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < data.length; i++) {
    hash ^= data[i]!;
    hash = Math.imul(hash, 0x01000193);
  }
  return `${(hash >>> 0).toString(16)}-${data.length}`;
}

function decodeBase64(base64: string): Uint8Array | undefined {
  try {
    const binary = atob(base64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
  } catch {
    return undefined;
  }
}

/** The data-URL mime + payload, or undefined for a non-data URL. */
function parseDataUrl(url: string): { mime: string; base64?: string; payload?: string } | undefined {
  const match = /^data:([^,]*),/.exec(url);
  if (!match) return undefined;
  const header = match[1] ?? "";
  const payload = url.slice(match[0].length);
  const isBase64 = /;base64$/.test(header);
  return {
    mime: header.replace(/;base64$/, ""),
    ...(isBase64 ? { base64: payload } : { payload }),
  };
}

const isJpeg = (bytes: Uint8Array): boolean =>
  bytes.length > 2 && bytes[0] === 0xff && bytes[1] === 0xd8;

/** The shared per-export image cache: url → decoded image data (or undefined
 *  for an undecodable source — the node is dropped, never rasterized). */
export class SceneImageCache {
  private readonly cache = new Map<string, Promise<PdfSceneImageData | undefined>>();

  get(
    url: string,
    naturalWidth?: number,
    naturalHeight?: number,
  ): Promise<PdfSceneImageData | undefined> {
    let entry = this.cache.get(url);
    if (!entry) {
      entry = this.#load(url, naturalWidth, naturalHeight);
      this.cache.set(url, entry);
    }
    return entry;
  }

  async #load(
    url: string,
    naturalWidth?: number,
    naturalHeight?: number,
  ): Promise<PdfSceneImageData | undefined> {
    const dataUrl = parseDataUrl(url);
    let bytes: Uint8Array | undefined;
    let mime = dataUrl?.mime;
    if (dataUrl) {
      bytes =
        dataUrl.base64 !== undefined
          ? decodeBase64(dataUrl.base64)
          : new TextEncoder().encode(decodeURIComponent(dataUrl.payload ?? ""));
    } else {
      try {
        const response = await fetch(url);
        if (!response.ok) return undefined;
        bytes = new Uint8Array(await response.arrayBuffer());
        mime = response.headers.get("content-type") ?? undefined;
      } catch {
        return undefined;
      }
    }
    if (!bytes || bytes.length === 0) return undefined;
    if (isJpeg(bytes) || mime?.includes("jpeg") || mime?.includes("jpg")) {
      const width = naturalWidth ?? 0;
      const height = naturalHeight ?? 0;
      if (!width || !height) return undefined;
      return { key: hashBytes(bytes), width, height, jpeg: bytes };
    }
    return this.#decodeRgba(bytes, mime, naturalWidth, naturalHeight);
  }

  /** PNG/WebP/GIF/SVG → straight RGBA (browser canvas decode), alpha stripped
   *  when every pixel is opaque. */
  async #decodeRgba(
    bytes: Uint8Array,
    mime: string | undefined,
    naturalWidth?: number,
    naturalHeight?: number,
  ): Promise<PdfSceneImageData | undefined> {
    if (typeof createImageBitmap !== "function") return undefined;
    let bitmap: ImageBitmap;
    try {
      bitmap = await createImageBitmap(
        new Blob([bytes as unknown as BlobPart], { type: mime ?? "image/png" }),
      );
    } catch {
      return undefined;
    }
    const width = naturalWidth ?? bitmap.width;
    const height = naturalHeight ?? bitmap.height;
    if (!width || !height || width * height > 64_000_000) {
      bitmap.close();
      return undefined;
    }
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    if (!ctx) {
      bitmap.close();
      return undefined;
    }
    ctx.drawImage(bitmap, 0, 0, width, height);
    bitmap.close();
    const source = ctx.getImageData(0, 0, width, height).data;
    let hasAlpha = false;
    for (let i = 3; i < source.length; i += 4) {
      if (source[i] !== 255) {
        hasAlpha = true;
        break;
      }
    }
    const channels = hasAlpha ? 4 : 3;
    const rgba = new Uint8Array(width * height * channels);
    for (let i = 0, j = 0; i < source.length; i += 4, j += channels) {
      rgba[j] = source[i]!;
      rgba[j + 1] = source[i + 1]!;
      rgba[j + 2] = source[i + 2]!;
      if (hasAlpha) rgba[j + 3] = source[i + 3]!;
    }
    return { key: hashBytes(rgba), width, height, rgba, hasAlpha };
  }
}

/** A Leafer element's paint property: strings pass through, paint arrays take
 *  their first color paint, gradients fall back to their first stop. */
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
    const paint = value as { type?: string; stops?: { color?: string }[] };
    if (paint.type === "linear" || paint.type === "radial") return paint.stops?.[0]?.color;
  }
  return undefined;
}

function numberOr(value: unknown, fallback?: number): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function strokeJoinOf(value: unknown): "miter" | "round" | "bevel" | undefined {
  return value === "round" || value === "bevel" || value === "miter" ? value : undefined;
}

function strokeCapOf(value: unknown): "butt" | "round" | "square" | undefined {
  if (value === "round" || value === "square") return value;
  if (value === "butt" || value === "none") return "butt";
  return undefined;
}

/** Serialize one page's painted tree. The page's CSS-px size comes from the
 *  caller (the section flow box — the App's canvas is zoomed, the page is not). */
export async function serializePageScene(
  app: App,
  pageWidthPx: number,
  pageHeightPx: number,
  imageCache: SceneImageCache,
): Promise<PdfScenePage> {
  const tree = app.tree as unknown as IUI | undefined;
  const nodes: PdfSceneNode[] = [];
  const skipped: Record<string, number> = {};
  if (!tree) return { width: pageWidthPx, height: pageHeightPx, nodes, skipped };
  const normalize = invertMatrix(tree.worldTransform as PdfMatrix);
  const matrixOf = (el: IUI): PdfMatrix =>
    composeMatrix(normalize, el.worldTransform as unknown as PdfMatrix);
  const childrenOf = (el: IUI): IUI[] =>
    (el as unknown as { children?: IUI[] }).children ?? [];

  const walk = async (parent: IUI, sink: PdfSceneNode[]): Promise<void> => {
    for (const el of childrenOf(parent)) {
      if (el.visible === false) continue;
      const worldOpacity = (el as unknown as { worldOpacity?: number }).worldOpacity ?? 1;
      if (worldOpacity <= 0) continue;
      const tag = el.tag;
      const matrix = matrixOf(el);
      const blendMode = (el as unknown as { blendMode?: string }).blendMode;
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
        const clipPath = el.getPathString(true, true);
        sink.push({
          type: "group",
          ...base,
          ...(clipPath ? { clip: { path: clipPath, matrix } } : {}),
          children,
        });
        continue;
      }
      if (SHAPE_TAGS.has(tag)) {
        const path = el.getPathString(true, true);
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
          textDrawData?: { rows?: LeaferTextRow[] };
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
        const bold = typeof weight === "number" ? weight >= 600 : weight === "bold" || weight === "600";
        const letterSpacing = numberOr(textEl.letterSpacing);
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
        const image = await imageCache.get(imageEl.url, imageEl.image?.width, imageEl.image?.height);
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
      // Nested roots (none today) recurse; non-visual tags are counted for
      // the export report — an unexpected tag means a silent visual gap.
      if (childrenOf(el).length > 0) await walk(el, sink);
      else skipped[tag] = (skipped[tag] ?? 0) + 1;
    }
  };

  await walk(tree, nodes);
  return {
    width: pageWidthPx,
    height: pageHeightPx,
    nodes,
    ...(Object.keys(skipped).length > 0 ? { skipped } : {}),
  };
}
