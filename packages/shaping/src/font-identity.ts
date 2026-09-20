import type { FontIdentity } from "./types.js";
import { isFontEmbeddingAllowed, isFontSubsettingAllowed } from "./types.js";

export interface FontDescriptor {
  readonly name: string;
  readonly postscriptName?: string;
  readonly familyName?: string;
  readonly weight?: number;
  readonly style?: "normal" | "italic";
  readonly script?: string;
  readonly hash?: string;
}

/** Read a table record (tag/offset/length) from an sfnt font's directory. */
function findTable(
  fontBytes: Uint8Array,
  tag: string,
): { offset: number; length: number } | undefined {
  if (fontBytes.byteLength < 12) return undefined;
  const view = new DataView(fontBytes.buffer, fontBytes.byteOffset, fontBytes.byteLength);
  const numTables = view.getUint16(4);
  for (let i = 0; i < numTables; i++) {
    const rec = 12 + i * 16;
    if (rec + 16 > fontBytes.byteLength) return undefined;
    if (
      fontBytes[rec] === tag.charCodeAt(0) &&
      fontBytes[rec + 1] === tag.charCodeAt(1) &&
      fontBytes[rec + 2] === tag.charCodeAt(2) &&
      fontBytes[rec + 3] === tag.charCodeAt(3)
    ) {
      return { offset: view.getUint32(rec + 8), length: view.getUint32(rec + 12) };
    }
  }
  return undefined;
}

/**
 * Read the OpenType `OS/2.fsType` embedding-permission bits directly from the
 * font bytes (no WASM needed). Returns 0 (installable) when OS/2 is absent.
 */
export function readFontFsType(fontBytes: Uint8Array): number {
  const os2 = findTable(fontBytes, "OS/2");
  if (!os2 || os2.offset + 10 > fontBytes.byteLength) return 0;
  return new DataView(fontBytes.buffer, fontBytes.byteOffset + os2.offset, 10).getUint16(8);
}

/** Read `head.unitsPerEm` directly from the font bytes (no WASM needed). */
export function readFontUnitsPerEm(fontBytes: Uint8Array): number {
  const head = findTable(fontBytes, "head");
  if (!head || head.offset + 20 > fontBytes.byteLength) return 1000;
  const length = Math.min(54, fontBytes.byteLength - head.offset);
  return new DataView(fontBytes.buffer, fontBytes.byteOffset + head.offset, length).getUint16(18);
}

/** Read every glyph's horizontal advance from `hmtx`, scaled to 1000/em —
 *  the widths a PDF CIDFont needs in its /W array (index = glyph ID). The
 *  last long metric repeats for glyphs past `numberOfHMetrics`. */
export function readGlyphAdvances(fontBytes: Uint8Array): number[] {
  const head = findTable(fontBytes, "head");
  const hhea = findTable(fontBytes, "hhea");
  const maxp = findTable(fontBytes, "maxp");
  const hmtx = findTable(fontBytes, "hmtx");
  if (!head || !hhea || !maxp || !hmtx) return [];
  const unitsPerEm = readFontUnitsPerEm(fontBytes) || 1000;
  const rawView = new DataView(fontBytes.buffer, fontBytes.byteOffset, fontBytes.byteLength);
  const numberOfHMetrics = rawView.getUint16(hhea.offset + 34);
  const numGlyphs = rawView.getUint16(maxp.offset + 4);
  if (numberOfHMetrics === 0 || numGlyphs === 0) return [];
  const lastAdvanceOffset = hmtx.offset + (numberOfHMetrics - 1) * 4;
  if (lastAdvanceOffset + 2 > fontBytes.byteLength) return [];
  const scale = 1000 / unitsPerEm;
  const advances: number[] = new Array(numGlyphs);
  for (let gid = 0; gid < numGlyphs; gid++) {
    const offset = gid < numberOfHMetrics ? hmtx.offset + gid * 4 : lastAdvanceOffset;
    const advance = offset + 2 <= fontBytes.byteLength ? rawView.getUint16(offset) : 0;
    advances[gid] = Math.round(advance * scale);
  }
  return advances;
}

/** Whether the font bytes may be embedded / subsetted per their OS/2.fsType. */
export function fontLicenseAllows(fontBytes: Uint8Array): {
  embed: boolean;
  subset: boolean;
  fsType: number;
} {
  const fsType = readFontFsType(fontBytes);
  return {
    embed: isFontEmbeddingAllowed(fsType),
    subset: isFontSubsettingAllowed(fsType),
    fsType,
  };
}

/**
 * Computes a deterministic SHA-256 hex digest for font byte content.
 */
export async function computeFontHash(fontBytes: Uint8Array): Promise<string> {
  const hashBuffer = await crypto.subtle.digest("SHA-256", fontBytes as BufferSource);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

export function createFontDescriptor(identity: FontIdentity, hash?: string): FontDescriptor {
  return {
    name: identity.familyName ?? identity.postscriptName ?? "Unknown Font",
    familyName: identity.familyName,
    postscriptName: identity.postscriptName,
    hash,
  };
}
