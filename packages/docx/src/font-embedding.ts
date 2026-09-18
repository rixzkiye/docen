import { fontLicenseAllows, isFontEmbeddingAllowed, isFontSubsettingAllowed } from "@docen/shaping";
import type { DocumentOptions } from "@office-open/docx";

/** The engine's embedded-font option shape (its own interface is internal). */
type EmbeddedFontOptions = NonNullable<DocumentOptions["fonts"]>[number];

/** A host-registered font the DOCX compiler may embed. */
export interface EmbeddableFontSource {
  /** CSS/Word family name the run styles reference. */
  readonly family: string;
  readonly fontData: Uint8Array;
  /** OS/2.fsType override; read from the bytes when omitted. */
  readonly fsType?: number;
}

/**
 * Prepare `DocumentOptions.fonts` for the DOCX compiler: every source whose
 * OS/2 fsType permits embedding becomes a `word/fonts/fontN.odttf` part
 * (obfuscation and fontTable/relationship wiring are the compiler's job).
 * Restricted fonts (fsType bit 0x0002) are skipped — embedding them would
 * violate the license. Full fonts are embedded, not subsets: DOCX consumers
 * shape with the font's own GSUB/GPOS, which a glyf-only subset would drop.
 */
export function prepareEmbeddedFonts(
  sources: readonly EmbeddableFontSource[],
): NonNullable<DocumentOptions["fonts"]> {
  const fonts: EmbeddedFontOptions[] = [];
  let index = 1;
  for (const source of sources) {
    const fsType = source.fsType ?? fontLicenseAllows(source.fontData).fsType;
    if (!isFontEmbeddingAllowed(fsType)) continue;
    fonts.push({
      name: source.family,
      odttfPath: `word/fonts/font${index}.odttf`,
      data: source.fontData,
      fontKey: generateFontKey().replace(/[{}]/g, ""),
    });
    index++;
  }
  return fonts;
}

export interface EmbeddedDocxFont {
  readonly name: string;
  readonly fontData: Uint8Array;
  readonly fontKey?: string;
  readonly fsType?: number;
}

export interface PreparedEmbeddedPart {
  readonly fontName: string;
  readonly partPath: string;
  readonly contentType: string;
  readonly data: Uint8Array;
  readonly fontKey: string;
  readonly relationshipId: string;
}

/**
 * Generates a standard uppercase GUID string wrapped in braces:
 * e.g. "{B09C156C-32AC-4BE8-B4C2-A62E8D2E49E0}"
 */
export function generateFontKey(): string {
  const bytes = new Uint8Array(16);
  if (typeof crypto !== "undefined" && crypto.getRandomValues) {
    crypto.getRandomValues(bytes);
  } else {
    for (let i = 0; i < 16; i++) {
      bytes[i] = Math.floor(Math.random() * 256);
    }
  }

  // Set UUID version 4 and variant
  bytes[6] = (bytes[6]! & 0x0f) | 0x40;
  bytes[8] = (bytes[8]! & 0x3f) | 0x80;

  const hex = Array.from(bytes)
    .map((b) => b.toString(16).toUpperCase().padStart(2, "0"))
    .join("");

  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
}

/**
 * Parses a GUID string into its 16 binary bytes.
 */
export function parseGuidToBytes(guid: string): Uint8Array {
  const cleaned = guid.replace(/[{}-]/g, "");
  if (cleaned.length !== 32) {
    throw new Error(`Invalid GUID format: "${guid}"`);
  }

  const bytes = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    bytes[i] = parseInt(cleaned.slice(i * 2, i * 2 + 2), 16);
  }
  return bytes;
}

/**
 * Obfuscates font data according to ECMA-376 Part 4 §14.2.14.
 * The first 32 bytes of the font file are XORed with the reversed 16-byte fontKey.
 */
export function obfuscateFont(fontData: Uint8Array, fontKey: string): Uint8Array {
  const keyBytes = parseGuidToBytes(fontKey);
  const reversedKey = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    reversedKey[i] = keyBytes[15 - i]!;
  }

  const out = new Uint8Array(fontData.length);
  out.set(fontData);

  const nObfuscated = Math.min(32, fontData.length);
  for (let i = 0; i < nObfuscated; i++) {
    out[i] = fontData[i]! ^ reversedKey[i % 16]!;
  }

  return out;
}

/**
 * Deobfuscates font data according to ECMA-376 Part 4 §14.2.14.
 * Symmetrical with obfuscateFont.
 */
export function deobfuscateFont(obfuscatedData: Uint8Array, fontKey: string): Uint8Array {
  return obfuscateFont(obfuscatedData, fontKey);
}

/**
 * Validates font embedding permissions based on OpenType fsType.
 */
export function checkFontEmbeddingPermitted(fontName: string, fsType?: number): void {
  if (fsType !== undefined && !isFontEmbeddingAllowed(fsType)) {
    throw new Error(
      `Font "${fontName}" has restricted license (fsType: 0x${fsType.toString(16)}). Embedding is forbidden by license.`,
    );
  }
}

/**
 * Validates font subsetting permissions based on OpenType fsType.
 */
export function checkFontSubsettingPermitted(fontName: string, fsType?: number): void {
  if (fsType !== undefined && !isFontSubsettingAllowed(fsType)) {
    throw new Error(
      `Font "${fontName}" has restricted license (fsType: 0x${fsType.toString(16)}). Subsetting is forbidden by license.`,
    );
  }
}

/**
 * Builds the word/fontTable.xml content referencing embedded fonts.
 */
export function buildFontTableXml(embeddedParts: readonly PreparedEmbeddedPart[]): string {
  const lines: string[] = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
  ];

  for (const part of embeddedParts) {
    lines.push(`  <w:font w:name="${escapeXml(part.fontName)}">`);
    lines.push(`    <w:embedRegular r:id="${part.relationshipId}" w:fontKey="${part.fontKey}"/>`);
    lines.push(`  </w:font>`);
  }

  lines.push(`</w:fonts>`);
  return lines.join("\n");
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
