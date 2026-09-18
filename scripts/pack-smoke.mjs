#!/usr/bin/env node
/**
 * Packaging smoke test — the published/vendored surface contract.
 *
 * Packs the shipping packages and proves, from the tarballs alone:
 *
 *  1. `@docen/shaping` ships `wasm/docen_shaping.wasm` and its default loader
 *     (`initShapingWasm()`, no explicit bytes) resolves it from the packed
 *     artifact — the asset Next.js/webpack consumers must configure.
 *  2. `@docen/docx` is self-contained: `@docen/markdown` (a never-published
 *     workspace package) is bundled into `dist`, not declared as a runtime
 *     dependency; `linkedom` (server HTML parse) IS a runtime dependency.
 *  3. A minimal consumer installs `@docen/docx` + `@docen/shaping` from the
 *     tarballs and exercises Markdown/DOCX/HTML conversion with no
 *     `@docen/markdown` in the dependency graph.
 *  4. `docen` packs with resolved (non-`workspace:`) dependency ranges.
 *
 * Run: `pnpm test:pack` (expects `pnpm build` to have produced dist first).
 */
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const FONT_FIXTURE = join(ROOT, "packages/shaping/test/fixtures/fonts/OpenSans-Regular.ttf");

const results = [];
function check(name, fn) {
  try {
    fn();
    results.push({ name, ok: true });
  } catch (err) {
    results.push({ name, ok: false, error: err instanceof Error ? err.message : String(err) });
    console.error(
      `FAIL ${name}: ${err instanceof Error ? (err.stack ?? err.message) : String(err)}`,
    );
  }
}

function run(cmd, args, options = {}) {
  return execFileSync(cmd, args, {
    cwd: ROOT,
    encoding: "utf8",
    stdio: ["ignore", "pipe", "pipe"],
    ...options,
  });
}

function pack(packageDir, destination) {
  const out = run(
    "pnpm",
    ["--config.ignore-scripts=true", "pack", "--pack-destination", destination],
    {
      cwd: join(ROOT, "packages", packageDir),
    },
  );
  const filename = out.trim().split("\n").pop();
  if (!filename || !existsSync(filename)) throw new Error(`pack produced no tarball: ${out}`);
  return filename;
}

/** `tar tzf` without assuming GNU tar flags beyond the common subset. */
function tarEntries(tarball) {
  return run("tar", ["tzf", tarball]).split("\n").filter(Boolean);
}

function tarFile(tarball, path) {
  return run("tar", ["xzf", tarball, "-O", path]);
}

const work = mkdtempSync(join(tmpdir(), "docen-pack-smoke-"));
try {
  const packs = join(work, "packs");
  mkdirSync(packs);

  for (const pkg of ["shaping", "docx", "core", "layout", "docen"]) {
    const distEntry = join("packages", pkg, "dist", "index.mjs");
    if (!existsSync(join(ROOT, distEntry))) {
      throw new Error(`${distEntry} missing — run \`pnpm build\` before the packaging smoke test`);
    }
  }

  const shapingTgz = pack("shaping", packs);
  const docxTgz = pack("docx", packs);
  const coreTgz = pack("core", packs);
  const layoutTgz = pack("layout", packs);
  const docenTgz = pack("docen", packs);

  check("shaping tarball ships the wasm asset", () => {
    const entries = tarEntries(shapingTgz);
    if (!entries.includes("package/wasm/docen_shaping.wasm")) {
      throw new Error(`wasm asset missing from tarball (${entries.length} entries)`);
    }
  });

  check("docx tarball declares linkedom and does not depend on @docen/markdown", () => {
    const manifest = JSON.parse(tarFile(docxTgz, "package/package.json"));
    if (manifest.dependencies?.linkedom !== "0.18.13") {
      throw new Error(
        `linkedom must be a runtime dependency, got ${manifest.dependencies?.linkedom}`,
      );
    }
    if (manifest.dependencies?.["@docen/markdown"]) {
      throw new Error("@docen/markdown must not be a runtime dependency (bundled instead)");
    }
    if (JSON.stringify(manifest).includes("workspace:")) {
      throw new Error("unresolved workspace: protocol in packed manifest");
    }
  });

  check("docx dist contains no @docen/markdown import (bundled)", () => {
    const entries = tarEntries(docxTgz).filter((e) => e.endsWith(".mjs") || e.endsWith(".d.mts"));
    const offenders = entries.filter((entry) =>
      tarFile(docxTgz, entry).includes("@docen/markdown"),
    );
    if (offenders.length > 0) throw new Error(`unbundled imports in ${offenders.join(", ")}`);
  });

  check("docen tarball has resolved dependency ranges", () => {
    const manifest = JSON.parse(tarFile(docenTgz, "package/package.json"));
    if (
      manifest.dependencies["@docen/docx"] !== "0.7.0" ||
      manifest.dependencies["@docen/editor"] !== "0.7.0"
    ) {
      throw new Error(`unexpected docen deps: ${JSON.stringify(manifest.dependencies)}`);
    }
    if (JSON.stringify(manifest).includes("workspace:")) {
      throw new Error("unresolved workspace: protocol in packed manifest");
    }
  });

  // ── Consumer install + runtime smoke ──
  const consumer = join(work, "consumer");
  mkdirSync(consumer);
  const fileSpec = (tgz) => `file:${tgz}`;
  writeFileSync(
    join(consumer, "package.json"),
    JSON.stringify(
      {
        private: true,
        type: "module",
        dependencies: {
          "@docen/docx": fileSpec(docxTgz),
          "@docen/shaping": fileSpec(shapingTgz),
          "@docen/core": fileSpec(coreTgz),
          "@docen/layout": fileSpec(layoutTgz),
        },
        pnpm: {
          overrides: {
            "@docen/shaping": fileSpec(shapingTgz),
            "@docen/core": fileSpec(coreTgz),
            "@docen/layout": fileSpec(layoutTgz),
          },
        },
      },
      null,
      2,
    ),
  );
  writeFileSync(
    join(consumer, "smoke.mjs"),
    `import { generateDOCXSync, parseDOCXSync, parseHTML, parseMarkdown, generateMarkdown } from "@docen/docx";
import { initShapingWasm, createFontRefSync } from "@docen/shaping";
import { readFileSync } from "node:fs";

const md = "# Title\\n\\nHello **world**.";
const json = parseMarkdown(md);
if (generateMarkdown(json).trimEnd() !== md) throw new Error("markdown round-trip mismatch");
console.log("markdown: OK (vendored, no @docen/markdown installed)");

const docx = generateDOCXSync(json);
if (parseDOCXSync(new Uint8Array(docx)).type !== "doc") throw new Error("docx round-trip failed");
console.log("docx: OK");

const html = await parseHTML("<p>Hello <b>bold</b></p>");
if (html.type !== "doc") throw new Error("html parse failed");
console.log("html server parse (linkedom runtime dep): OK");

await initShapingWasm();
const font = createFontRefSync(new Uint8Array(readFileSync(process.env.DOCEN_SMOKE_FONT)));
const shaped = font.shape("office fi", { direction: "ltr" });
if (shaped.glyphs.length === 0) throw new Error("shaping produced no glyphs");
console.log("wasm loader from packed tarball: OK glyphs=" + shaped.glyphs.length);
`,
  );
  run("pnpm", ["install", "--prefer-offline", "--no-frozen-lockfile"], { cwd: consumer });
  const smokeOut = run("node", ["smoke.mjs"], {
    cwd: consumer,
    env: { ...process.env, DOCEN_SMOKE_FONT: FONT_FIXTURE },
  });
  check("consumer installs and runs from tarballs", () => {
    for (const marker of [
      "markdown: OK (vendored, no @docen/markdown installed)",
      "docx: OK",
      "html server parse (linkedom runtime dep): OK",
      "wasm loader from packed tarball: OK",
    ]) {
      if (!smokeOut.includes(marker))
        throw new Error(`missing smoke marker: ${marker}\n${smokeOut}`);
    }
  });

  const failed = results.filter((r) => !r.ok);
  console.log("\npackaging smoke summary:");
  for (const r of results) console.log(`  ${r.ok ? "OK  " : "FAIL"} ${r.name}`);
  if (failed.length > 0) {
    process.exitCode = 1;
  } else {
    console.log("all packaging checks passed");
  }
} finally {
  rmSync(work, { recursive: true, force: true });
}
