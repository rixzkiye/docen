import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

/**
 * Packaging contract gate (item 9) — the real thing, not a restatement:
 * `scripts/pack-smoke.mjs` packs the shipping packages with pnpm, inspects the
 * tarballs (wasm asset present, `linkedom` runtime dep, no bundled
 * `@docen/markdown` dependency), installs a minimal consumer from the
 * tarballs, and runs Markdown/DOCX/HTML conversion plus the default wasm
 * loader from the packed artifacts. Requires `pnpm build` (the CI gate order
 * is build → test).
 */
const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "../../..");

describe("published surface packaging", () => {
  it("packs, inspects and consumes the tarballs", { timeout: 240_000 }, () => {
    if (!existsSync(join(ROOT, "packages/docx/dist/index.mjs"))) {
      throw new Error("packages/docx/dist missing — run `pnpm build` before the packaging test");
    }
    const out = execFileSync("node", ["scripts/pack-smoke.mjs"], {
      cwd: ROOT,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
    });
    expect(out).toContain("all packaging checks passed");
  });
});
