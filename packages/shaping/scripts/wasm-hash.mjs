#!/usr/bin/env node
// Reproducible-build hash check for the @docen/shaping WASM artifact.
//
//   node scripts/wasm-hash.mjs --check   # verify artifact vs manifest (CI)
//   node scripts/wasm-hash.mjs --write   # record the current artifact's hash
//
// The manifest (`wasm/docen_shaping.wasm.sha256`) is a JSON record of the
// pinned-toolchain build. `--check` is also a test oracle: the same hash is
// asserted by test/reproducibility.spec.ts.

import { createHash } from "node:crypto";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const wasmPath = join(root, "wasm", "docen_shaping.wasm");
const manifestPath = join(root, "wasm", "docen_shaping.wasm.sha256");

function digest(path) {
  const bytes = readFileSync(path);
  return {
    sha256: createHash("sha256").update(bytes).digest("hex"),
    bytes: bytes.byteLength,
  };
}

const mode = process.argv.includes("--write") ? "write" : "check";

if (!existsSync(wasmPath)) {
  console.error(`wasm-hash: missing artifact ${wasmPath} — run build:wasm first`);
  process.exit(1);
}

const artifact = digest(wasmPath);

if (mode === "write") {
  const manifest = {
    file: "docen_shaping.wasm",
    ...artifact,
    toolchain: "rust-toolchain.toml",
  };
  writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
  console.log(`wasm-hash: recorded ${artifact.sha256} (${artifact.bytes} bytes)`);
  process.exit(0);
}

if (!existsSync(manifestPath)) {
  console.error(`wasm-hash: missing manifest ${manifestPath} — run with --write`);
  process.exit(1);
}

const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
if (manifest.sha256 !== artifact.sha256 || manifest.bytes !== artifact.bytes) {
  console.error(
    [
      "wasm-hash: MISMATCH — the committed artifact does not match the pinned build.",
      `  manifest: ${manifest.sha256} (${manifest.bytes} bytes)`,
      `  artifact: ${artifact.sha256} (${artifact.bytes} bytes)`,
      "  Rebuild with the pinned toolchain (pnpm --filter @docen/shaping build:wasm).",
      "  If the change is intentional, update the crate and record the new hash with",
      "  DOCEN_WASM_UPDATE=1 pnpm --filter @docen/shaping build:wasm.",
    ].join("\n"),
  );
  process.exit(1);
}

console.log(`wasm-hash: ok ${artifact.sha256} (${artifact.bytes} bytes)`);
