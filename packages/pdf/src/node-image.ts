import { inflateSync } from "node:zlib";

import type { PdfSceneImageData } from "./pdf-scene";

export interface PdfImageCache {
  get(
    url: string,
    naturalWidth?: number,
    naturalHeight?: number,
  ): Promise<PdfSceneImageData | undefined>;
}

export function hashBytes(data: Uint8Array): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < data.length; i++) {
    hash ^= data[i]!;
    hash = Math.imul(hash, 0x01000193);
  }
  return `${(hash >>> 0).toString(16)}-${data.length}`;
}

export function decodeBase64(base64: string): Uint8Array | undefined {
  try {
    if (typeof Buffer !== "undefined") {
      return new Uint8Array(Buffer.from(base64, "base64"));
    }
    const binary = atob(base64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
  } catch {
    return undefined;
  }
}

export function parseDataUrl(
  url: string,
): { mime: string; base64?: string; payload?: string } | undefined {
  const match = /^data:([^,]*),/.exec(url);
  if (!match) return undefined;
  const header = match[1] ?? "";
  const payload = url.slice(match[0].length);
  const isBase64 = header.endsWith(";base64");
  return {
    mime: header.replace(/;base64$/, ""),
    ...(isBase64 ? { base64: payload } : { payload }),
  };
}

export function isJpeg(bytes: Uint8Array): boolean {
  return bytes.length > 2 && bytes[0] === 0xff && bytes[1] === 0xd8;
}

export function parseJpegDimensions(
  bytes: Uint8Array,
): { width: number; height: number } | undefined {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined;
  let offset = 2;
  while (offset < bytes.length - 8) {
    if (bytes[offset] !== 0xff) {
      offset++;
      continue;
    }
    const marker = bytes[offset + 1]!;
    if (marker === 0xd9) break; // EOI
    if (marker === 0xda) break; // SOS
    const len = (bytes[offset + 2]! << 8) | bytes[offset + 3]!;
    if (
      (marker >= 0xc0 && marker <= 0xc3) ||
      (marker >= 0xc5 && marker <= 0xc7) ||
      (marker >= 0xc9 && marker <= 0xcb) ||
      (marker >= 0xcd && marker <= 0xcf)
    ) {
      const height = (bytes[offset + 5]! << 8) | bytes[offset + 6]!;
      const width = (bytes[offset + 7]! << 8) | bytes[offset + 8]!;
      return { width, height };
    }
    offset += 2 + len;
  }
  return undefined;
}

export function isPng(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a
  );
}

function paeth(a: number, b: number, c: number): number {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  if (pb <= pc) return b;
  return c;
}

export function decodePngToRgba(bytes: Uint8Array):
  | {
      width: number;
      height: number;
      rgba: Uint8Array;
      hasAlpha: boolean;
    }
  | undefined {
  if (!isPng(bytes)) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  const idatChunks: Uint8Array[] = [];
  let palette: Uint8Array | undefined;
  let trns: Uint8Array | undefined;

  let offset = 8;
  while (offset < bytes.length) {
    if (offset + 8 > bytes.length) break;
    const len = view.getUint32(offset);
    const type = String.fromCharCode(
      bytes[offset + 4]!,
      bytes[offset + 5]!,
      bytes[offset + 6]!,
      bytes[offset + 7]!,
    );
    const chunkDataOffset = offset + 8;
    if (type === "IHDR") {
      width = view.getUint32(chunkDataOffset);
      height = view.getUint32(chunkDataOffset + 4);
      bitDepth = bytes[chunkDataOffset + 8]!;
      colorType = bytes[chunkDataOffset + 9]!;
    } else if (type === "PLTE") {
      palette = bytes.subarray(chunkDataOffset, chunkDataOffset + len);
    } else if (type === "tRNS") {
      trns = bytes.subarray(chunkDataOffset, chunkDataOffset + len);
    } else if (type === "IDAT") {
      idatChunks.push(bytes.subarray(chunkDataOffset, chunkDataOffset + len));
    } else if (type === "IEND") {
      break;
    }
    offset += 12 + len; // 4 len + 4 type + data + 4 crc
  }

  if (width === 0 || height === 0 || idatChunks.length === 0) return undefined;
  if (bitDepth !== 8 && bitDepth !== 16 && !(bitDepth < 8 && colorType === 3)) {
    // 8-bit or indexed standard PNGs are supported
  }

  // Concatenate IDAT
  const totalIdat = idatChunks.reduce((acc, c) => acc + c.length, 0);
  const compressed = new Uint8Array(totalIdat);
  let idatPos = 0;
  for (const c of idatChunks) {
    compressed.set(c, idatPos);
    idatPos += c.length;
  }

  let uncompressed: Uint8Array;
  try {
    uncompressed = new Uint8Array(inflateSync(compressed));
  } catch {
    return undefined;
  }

  let bpp = 1;
  if (colorType === 2) bpp = 3; // RGB
  else if (colorType === 4) bpp = 2; // Gray + Alpha
  else if (colorType === 6) bpp = 4; // RGBA
  else if (colorType === 0) bpp = 1; // Gray
  else if (colorType === 3) bpp = 1; // Indexed

  const rowSize = 1 + width * bpp;
  if (uncompressed.length < height * rowSize) return undefined;

  const rawPixels = new Uint8Array(width * height * bpp);
  let priorRow: Uint8Array | null = null;

  for (let y = 0; y < height; y++) {
    const rowStart = y * rowSize;
    const filter = uncompressed[rowStart]!;
    const rowData = uncompressed.subarray(rowStart + 1, rowStart + rowSize);
    const destStart = y * width * bpp;
    const currentRow = rawPixels.subarray(destStart, destStart + width * bpp);

    for (let i = 0; i < width * bpp; i++) {
      const left = i >= bpp ? currentRow[i - bpp]! : 0;
      const up = priorRow ? priorRow[i]! : 0;
      const upLeft = priorRow && i >= bpp ? priorRow[i - bpp]! : 0;
      const val = rowData[i]!;

      let unfiltered = 0;
      switch (filter) {
        case 0: // None
          unfiltered = val;
          break;
        case 1: // Sub
          unfiltered = (val + left) & 0xff;
          break;
        case 2: // Up
          unfiltered = (val + up) & 0xff;
          break;
        case 3: // Average
          unfiltered = (val + Math.floor((left + up) / 2)) & 0xff;
          break;
        case 4: // Paeth
          unfiltered = (val + paeth(left, up, upLeft)) & 0xff;
          break;
        default:
          unfiltered = val;
      }
      currentRow[i] = unfiltered;
    }
    priorRow = currentRow;
  }

  // Convert rawPixels to RGBA
  const rgba = new Uint8Array(width * height * 4);
  let hasAlpha = false;

  if (colorType === 6) {
    // RGBA already
    rgba.set(rawPixels);
    for (let i = 3; i < rgba.length; i += 4) {
      if (rgba[i] !== 255) {
        hasAlpha = true;
        break;
      }
    }
  } else if (colorType === 2) {
    // RGB -> RGBA
    for (let i = 0, j = 0; i < rawPixels.length; i += 3, j += 4) {
      rgba[j] = rawPixels[i]!;
      rgba[j + 1] = rawPixels[i + 1]!;
      rgba[j + 2] = rawPixels[i + 2]!;
      rgba[j + 3] = 255;
    }
  } else if (colorType === 3 && palette) {
    // Indexed
    for (let i = 0, j = 0; i < rawPixels.length; i++, j += 4) {
      const idx = rawPixels[i]!;
      rgba[j] = palette[idx * 3] ?? 0;
      rgba[j + 1] = palette[idx * 3 + 1] ?? 0;
      rgba[j + 2] = palette[idx * 3 + 2] ?? 0;
      const alpha = trns && idx < trns.length ? trns[idx]! : 255;
      rgba[j + 3] = alpha;
      if (alpha !== 255) hasAlpha = true;
    }
  } else if (colorType === 0) {
    // Grayscale
    for (let i = 0, j = 0; i < rawPixels.length; i++, j += 4) {
      const g = rawPixels[i]!;
      rgba[j] = g;
      rgba[j + 1] = g;
      rgba[j + 2] = g;
      rgba[j + 3] = 255;
    }
  } else if (colorType === 4) {
    // Gray + Alpha
    for (let i = 0, j = 0; i < rawPixels.length; i += 2, j += 4) {
      const g = rawPixels[i]!;
      const a = rawPixels[i + 1]!;
      rgba[j] = g;
      rgba[j + 1] = g;
      rgba[j + 2] = g;
      rgba[j + 3] = a;
      if (a !== 255) hasAlpha = true;
    }
  }

  return { width, height, rgba, hasAlpha };
}

export class NodeImageCache implements PdfImageCache {
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
    } else if (url.startsWith("http://") || url.startsWith("https://")) {
      try {
        const response = await fetch(url);
        if (!response.ok) return undefined;
        bytes = new Uint8Array(await response.arrayBuffer());
        mime = response.headers.get("content-type") ?? undefined;
      } catch {
        return undefined;
      }
    } else {
      // Local file or relative path
      try {
        const fs = await import("node:fs");
        bytes = new Uint8Array(fs.readFileSync(url));
      } catch {
        return undefined;
      }
    }

    if (!bytes || bytes.length === 0) return undefined;

    if (isJpeg(bytes) || mime?.includes("jpeg") || mime?.includes("jpg")) {
      const dims = parseJpegDimensions(bytes);
      const width = dims?.width ?? naturalWidth ?? 0;
      const height = dims?.height ?? naturalHeight ?? 0;
      if (!width || !height) return undefined;
      return { key: hashBytes(bytes), width, height, jpeg: bytes };
    }

    if (isPng(bytes) || mime?.includes("png")) {
      const decoded = decodePngToRgba(bytes);
      if (decoded) {
        const width = naturalWidth ?? decoded.width;
        const height = naturalHeight ?? decoded.height;
        // PdfSceneImageData's channel contract (shared with the browser
        // scene-export path): 3-channel RGB when the bitmap is opaque,
        // 4-channel RGBA when it has alpha — the exporter strips alpha into
        // an /SMask, but embeds an opaque buffer as raw DeviceRGB as-is.
        // Returning 4 channels for an opaque bitmap would write RGBA bytes
        // into a DeviceRGB image and shear every row.
        let rgba = decoded.rgba;
        if (!decoded.hasAlpha) {
          const pixels = decoded.width * decoded.height;
          const rgb = new Uint8Array(pixels * 3);
          for (let p = 0; p < pixels; p++) {
            rgb[p * 3] = decoded.rgba[p * 4]!;
            rgb[p * 3 + 1] = decoded.rgba[p * 4 + 1]!;
            rgb[p * 3 + 2] = decoded.rgba[p * 4 + 2]!;
          }
          rgba = rgb;
        }
        return {
          key: hashBytes(rgba),
          width,
          height,
          rgba,
          hasAlpha: decoded.hasAlpha,
        };
      }
    }

    return undefined;
  }
}
