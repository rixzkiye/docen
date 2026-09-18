// WASM artifact reproducibility. The pinned toolchain
// (packages/shaping/rust-toolchain.toml) plus `scripts/build-wasm.sh` must
// produce the exact bytes recorded in `wasm/docen_shaping.wasm.sha256`; this
// spec re-computes that hash independently and proves the artifact is a
// loadable module with the ABI entry points the engine imports.

import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const wasm = readFileSync(join(root, "wasm", "docen_shaping.wasm"));
const manifest = JSON.parse(
  readFileSync(join(root, "wasm", "docen_shaping.wasm.sha256"), "utf8"),
) as { file: string; sha256: string; bytes: number; toolchain: string };

describe("wasm artifact reproducibility", () => {
  it("matches the committed build hash", () => {
    const sha256 = createHash("sha256").update(wasm).digest("hex");
    expect(manifest.file).toBe("docen_shaping.wasm");
    expect(manifest.sha256).toMatch(/^[a-f0-9]{64}$/);
    expect(sha256).toBe(manifest.sha256);
    expect(wasm.byteLength).toBe(manifest.bytes);
  });

  it("is a loadable module exposing the shaping ABI", async () => {
    const module = await WebAssembly.compile(wasm);
    const instance = await WebAssembly.instantiate(module, {});
    for (const symbol of [
      "alloc",
      "dealloc",
      "register_font",
      "register_font_index",
      "drop_font",
      "shape_text",
      "get_font_metrics",
      "get_glyph_outline",
      "get_font_name",
      "get_glyph_count",
      "get_font_fs_type",
    ]) {
      expect(typeof (instance.exports as Record<string, unknown>)[symbol]).toBe("function");
    }
  });

  it("pins the toolchain that produced the hash", () => {
    const pin = readFileSync(join(root, "rust-toolchain.toml"), "utf8");
    expect(manifest.toolchain).toBe("rust-toolchain.toml");
    expect(pin).toMatch(/channel\s*=\s*"\d+\.\d+\.\d+"/);
    expect(pin).toContain('"wasm32-unknown-unknown"');
  });
});
