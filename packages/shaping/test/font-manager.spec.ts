import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { FontManager, computeFontHash, initShapingWasm } from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.join(__dirname, "fixtures/fonts/OpenSans-Regular.ttf");
const arabicPath = path.join(__dirname, "fixtures/fonts/NotoNaskhArabic-Regular.ttf");

describe("R6.2 FontManager (Sources, Cache, Fallback, Diagnostics)", () => {
  let openSansBytes: Uint8Array;
  let arabicBytes: Uint8Array;

  beforeAll(async () => {
    await initShapingWasm();
    openSansBytes = fs.readFileSync(openSansPath);
    arabicBytes = fs.readFileSync(arabicPath);
  });

  it("computes deterministic SHA-256 hash of font bytes", async () => {
    const hash1 = await computeFontHash(openSansBytes);
    const hash2 = await computeFontHash(openSansBytes);

    expect(hash1).toBe(hash2);
    expect(hash1.length).toBe(64);
  });

  it("resolves fonts from memory sources and caches them", async () => {
    const manager = new FontManager({ maxMemoryEntries: 10, enableOpfs: false });
    manager.registerFontBytes("Open Sans", openSansBytes);

    const font1 = await manager.resolveFont({ name: "Open Sans" });
    expect(font1.familyName).toBe("Open Sans");

    // Resolving again returns the cached instance
    const font2 = await manager.resolveFont({ name: "Open Sans" });
    expect(font1).toBe(font2);

    manager.dispose();
  });

  it("evicts oldest entries when cache exceeds max capacity (LRU)", async () => {
    const manager = new FontManager({ maxMemoryEntries: 2, enableOpfs: false });
    await manager.cache.set("font-a", new Uint8Array([1]));
    await manager.cache.set("font-b", new Uint8Array([2]));
    await manager.cache.set("font-c", new Uint8Array([3])); // evicts font-a

    expect(await manager.cache.get("font-a")).toBeUndefined();
    expect(await manager.cache.get("font-b")).toBeDefined();
    expect(await manager.cache.get("font-c")).toBeDefined();
  });

  it("invalidates cached font refs and data properly", async () => {
    const manager = new FontManager({ enableOpfs: false });
    manager.registerFontBytes("Open Sans", openSansBytes);

    const font = await manager.resolveFont({ name: "Open Sans" });
    expect(font).toBeDefined();

    await manager.invalidate("Open Sans");
    expect(await manager.cache.get("open sans")).toBeUndefined();

    // Resolving again reconstructs a new FontRef
    const newFont = await manager.resolveFont({ name: "Open Sans" });
    expect(newFont).not.toBe(font);

    manager.dispose();
  });

  it("resolves script-specific fallback and records missing glyph diagnostics", async () => {
    const manager = new FontManager({ enableOpfs: false });
    manager.registerFontBytes("Open Sans", openSansBytes);
    manager.registerFontBytes("Noto Naskh Arabic", arabicBytes);

    // Latin character 'A' resolves via Open Sans
    const latinFallback = await manager.resolveFallback("A", "Latn");
    expect(latinFallback).toBeDefined();
    expect(latinFallback?.familyName).toBe("Open Sans");

    // Arabic character 'م' resolves via Noto Naskh Arabic
    const arabicFallback = await manager.resolveFallback("م", "Arab");
    expect(arabicFallback).toBeDefined();

    // Query character not in any registered font (e.g. unassigned Private Use Area)
    manager.clearMissingGlyphs();
    const missing = await manager.resolveFallback("\uE000", "Latn");
    expect(missing).toBeUndefined();

    const reports = manager.getMissingGlyphs();
    expect(reports.length).toBe(1);
    expect(reports[0].codePoint).toBe(0xe000);

    manager.dispose();
  });
});
