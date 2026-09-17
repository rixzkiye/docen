import { describe, expect, it } from "vitest";

import {
  buildFontTableXml,
  checkFontEmbeddingPermitted,
  checkFontSubsettingPermitted,
  deobfuscateFont,
  generateFontKey,
  obfuscateFont,
  parseGuidToBytes,
} from "./font-embedding";

describe("DOCX Font Embedding (ECMA-376 §14.2.14)", () => {
  it("generates a valid uppercase GUID wrapped in braces", () => {
    const key = generateFontKey();
    expect(key).toMatch(
      /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/,
    );
  });

  it("parses GUID to exactly 16 bytes", () => {
    const key = "{B09C156C-32AC-4BE8-B4C2-A62E8D2E49E0}";
    const bytes = parseGuidToBytes(key);
    expect(bytes.length).toBe(16);
    expect(bytes[0]).toBe(0xb0);
    expect(bytes[1]).toBe(0x9c);
    expect(bytes[15]).toBe(0xe0);
  });

  it("obfuscates and symmetrically deobfuscates font data", () => {
    const key = "{12345678-ABCD-4EF0-9123-456789ABCDEF}";
    const original = new Uint8Array(64);
    for (let i = 0; i < 64; i++) {
      original[i] = i;
    }

    const obfuscated = obfuscateFont(original, key);
    expect(obfuscated.length).toBe(64);
    // First 32 bytes must be modified
    expect(obfuscated.slice(0, 32)).not.toEqual(original.slice(0, 32));
    // Bytes after 32 must be untouched
    expect(obfuscated.slice(32)).toEqual(original.slice(32));

    const deobfuscated = deobfuscateFont(obfuscated, key);
    expect(deobfuscated).toEqual(original);
  });

  it("validates embedding permission based on fsType", () => {
    // Installable (0) or preview/print (4) or editable (8) allowed
    expect(() => checkFontEmbeddingPermitted("FontA", 0)).not.toThrow();
    expect(() => checkFontEmbeddingPermitted("FontB", 0x0008)).not.toThrow();
    expect(() => checkFontEmbeddingPermitted("FontC", undefined)).not.toThrow();

    // Restricted license embedding (0x0002) forbidden
    expect(() => checkFontEmbeddingPermitted("RestrictedFont", 0x0002)).toThrow(
      /RestrictedFont.*forbidden/,
    );
  });

  it("validates subsetting permission based on fsType", () => {
    expect(() => checkFontSubsettingPermitted("FontA", 0)).not.toThrow();
    expect(() => checkFontSubsettingPermitted("FontB", 0x0008)).not.toThrow();
    expect(() => checkFontSubsettingPermitted("FontC", undefined)).not.toThrow();

    // No subsetting bit (0x0100) forbidden
    expect(() => checkFontSubsettingPermitted("NoSubsetFont", 0x0100)).toThrow(
      /NoSubsetFont.*forbidden/,
    );
  });

  it("builds correct word/fontTable.xml", () => {
    const xml = buildFontTableXml([
      {
        fontName: "CustomFont",
        fontKey: "{B09C156C-32AC-4BE8-B4C2-A62E8D2E49E0}",
        relationshipId: "rId10",
        partPath: "word/fonts/font1.odttf",
        contentType: "application/vnd.openxmlformats-officedocument.obfuscatedFont",
        data: new Uint8Array(10),
      },
    ]);

    expect(xml).toContain('<w:font w:name="CustomFont">');
    expect(xml).toContain(
      '<w:embedRegular r:id="rId10" w:fontKey="{B09C156C-32AC-4BE8-B4C2-A62E8D2E49E0}"/>',
    );
  });
});
