import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { RustybuzzBackend, initShapingWasm } from "../src/index.js";
import { readCmap, subsetFontWithPlan } from "../src/subsetter.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fontsDir = path.join(__dirname, "fixtures/fonts");
const gateScript = path.join(__dirname, "subset-gate.py");

interface FontCase {
  name: string;
  file: string;
  text: string;
  options?: { direction?: "rtl"; script?: string };
}

const CASES: FontCase[] = [
  { name: "OpenSans Latin", file: "OpenSans-Regular.ttf", text: "Hello café World!" },
  {
    name: "Noto Naskh Arabic",
    file: "NotoNaskhArabic-Regular.ttf",
    text: "مرحبا",
    options: { direction: "rtl", script: "Arab" },
  },
  { name: "Noto Sans Devanagari", file: "NotoSansDevanagari-Regular.ttf", text: "नमस्ते" },
];

describe("R6.7 subset gate (fontTools + wasm re-shape)", () => {
  let backend: RustybuzzBackend;

  beforeAll(async () => {
    await initShapingWasm();
    backend = new RustybuzzBackend();
  });

  for (const fontCase of CASES) {
    it(`emits a valid, cmap-consistent subset for ${fontCase.name}`, () => {
      const originalPath = path.join(fontsDir, fontCase.file);
      const originalBytes = fs.readFileSync(originalPath);
      const origId = backend.registerFont(originalBytes);
      const shaped = backend.shape(origId, fontCase.text, fontCase.options);

      // Glyph closure: shaped glyphs (ligatures/marks) plus the cmap lookup of
      // every code point the text uses.
      const cmap = readCmap(originalBytes);
      const used = new Set<number>([0, ...shaped.glyphs.map((g) => g.glyphId)]);
      const codePoints: number[] = [];
      for (const ch of fontCase.text) {
        const cp = ch.codePointAt(0)!;
        codePoints.push(cp);
        const gid = cmap.get(cp);
        if (gid !== undefined) used.add(gid);
      }

      const plan = subsetFontWithPlan(originalBytes, [...used]);
      expect(plan.data.length).toBeGreaterThan(0);
      expect(plan.data.length).toBeLessThanOrEqual(originalBytes.length);

      const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "docen-subset-gate-"));
      const subsetPath = path.join(tmp, "subset.ttf");
      fs.writeFileSync(subsetPath, plan.data);

      // Per-code-point oracle: the original font's hmtx advance (read by the
      // fontTools validator) plus the subset's cmap mapping. The subset drops
      // GSUB/GPOS, so cross-glyph substitutions are intentionally not an
      // oracle; per-codepoint addresses and advances are.
      const checks: { cp: number; subsetGid: number }[] = [];
      const expectedAdvances = new Map<number, number>();
      for (const cp of new Set(codePoints)) {
        const gid = cmap.get(cp);
        if (gid === undefined) continue;
        const newGid = plan.glyphMap.get(gid);
        expect(newGid, `U+${cp.toString(16)} present in plan`).toBeDefined();
        checks.push({ cp, subsetGid: newGid! });
      }
      expect(checks.length).toBeGreaterThan(0);

      const configPath = path.join(tmp, "gate.json");
      fs.writeFileSync(
        configPath,
        JSON.stringify({ original: originalPath, subset: subsetPath, checks }),
      );
      const stdout = (() => {
        try {
          return execFileSync("python3", [gateScript, configPath], { encoding: "utf-8" });
        } catch (err) {
          const out = (err as { stdout?: string | Buffer }).stdout?.toString() ?? "";
          throw new Error(`fontTools gate failed: ${out || String(err)}`);
        }
      })();
      const result = JSON.parse(stdout) as {
        errors: string[];
        glyphCount: number;
        advances: Record<string, number>;
      };
      expect(result.errors).toEqual([]);
      expect(result.glyphCount).toBe(plan.glyphMap.size);
      for (const check of checks) {
        expectedAdvances.set(check.cp, result.advances[String(check.cp)] ?? 0);
      }

      // Register the subset in the SAME wasm and re-shape each code point:
      // glyph IDs must resolve through the rebuilt cmap and advances must be
      // the original hmtx advances.
      const subsetId = backend.registerFont(plan.data);
      expect(subsetId).toBeGreaterThan(0);
      for (const check of checks) {
        const subsetShape = backend.shape(
          subsetId,
          String.fromCodePoint(check.cp),
          fontCase.options,
        );
        expect(subsetShape.glyphs).toHaveLength(1);
        expect(subsetShape.glyphs[0]!.glyphId).toBe(check.subsetGid);
        expect(subsetShape.glyphs[0]!.xAdvance).toBe(expectedAdvances.get(check.cp));
      }

      fs.rmSync(tmp, { recursive: true, force: true });
    });
  }

  it("keeps ligature/context glyphs alive in the subset outline data", () => {
    const originalBytes = fs.readFileSync(path.join(fontsDir, "OpenSans-Regular.ttf"));
    const origId = backend.registerFont(originalBytes);
    const shaped = backend.shape(origId, "fi ffi");
    const ligatureGids = shaped.glyphs.map((g) => g.glyphId);
    expect(shaped.glyphs.length).toBeLessThan("fi ffi".length); // ligated

    const plan = subsetFontWithPlan(originalBytes, [0, ...ligatureGids]);
    const subsetId = backend.registerFont(plan.data);
    const subsetShape = backend.shape(subsetId, "fi ffi");
    // Without GSUB tables the subset maps through cmap to f+i; both glyphs and
    // their advances must exist (valid outlines), not .notdef.
    for (const glyph of subsetShape.glyphs) {
      expect(glyph.glyphId).not.toBe(0);
      expect(glyph.xAdvance).toBeGreaterThan(0);
    }
  });
});
