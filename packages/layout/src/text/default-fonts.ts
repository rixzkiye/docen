// Production font set for the document pipeline.
//
// Word documents reference metric-specific family names (Calibri, Cambria,
// Arial, Times New Roman). Without registered faces the shaping measurer
// silently falls back to canvas `measureText`, whose advances differ per
// browser/OS. This module bundles metric-compatible open faces (OFL) for the
// Word defaults under `packages/layout/assets/fonts` and registers them with
// the shaping font manager:
//
//   Calibri          → Carlito
//   Calibri Light    → Carlito (no Light face exists; regular metrics)
//   Cambria          → Caladea
//   Arial            → Liberation Sans
//   Times New Roman  → Liberation Serif
//
// Each family registers all four weight/slant slots (regular/bold/italic/
// bold-italic); `ShapedMeasurer` selects the slot and falls back to regular.
// The bytes are pinned by the font-metrics golden
// (`test/font-metrics-golden.json`), which hashes them and the shaped runs, so
// another machine can reproduce and verify the metrics.

import { initShapingWasm } from "@docen/shaping";

import { registerShapingFont, type ShapingFontStyle } from "./shaped-measurer";

/** One bundled face: the Word family it stands in for + its file. */
export interface DefaultFontFace extends ShapingFontStyle {
  /** The Word family name documents reference (registered as the family). */
  readonly family: string;
  /** Path under the package's `assets/fonts` directory. */
  readonly file: string;
}

function family(name: string, dir: string, stem: string): readonly DefaultFontFace[] {
  return [
    { family: name, file: `${dir}/${stem}-Regular.ttf` },
    { family: name, file: `${dir}/${stem}-Bold.ttf`, bold: true },
    { family: name, file: `${dir}/${stem}-Italic.ttf`, italic: true },
    { family: name, file: `${dir}/${stem}-BoldItalic.ttf`, bold: true, italic: true },
  ];
}

/** The bundled production set: every Word default family × four styles. */
export const DEFAULT_FONT_FACES: readonly DefaultFontFace[] = [
  ...family("Calibri", "carlito", "Carlito"),
  // Word's heading theme font has no metric-compatible Light face in the
  // Carlito family; registering Carlito's regular slot keeps Calibri Light
  // deterministic instead of canvas-probed (documented approximation).
  ...family("Calibri Light", "carlito", "Carlito"),
  ...family("Cambria", "caladea", "Caladea"),
  ...family("Arial", "liberation-sans", "LiberationSans"),
  ...family("Times New Roman", "liberation-serif", "LiberationSerif"),
];

/** Candidate asset roots relative to this module: `src/text/` in the source
 *  tree, `dist/` in the bundled package. The first one holding the assets (or
 *  answering a fetch) wins. */
const ASSET_ROOT_CANDIDATES = ["../../assets/fonts/", "../assets/fonts/"] as const;

function relativeAssetUrl(candidate: string, file: string): URL {
  return new URL(`${candidate}${file}`, import.meta.url);
}

/** Resolve a bundled face's URL. Node probes the filesystem; the browser
 *  fetches the candidates in order. Consumers whose bundler does not emit
 *  `new URL` assets can pass an explicit `baseUrl` (see `@docen/shaping`'s
 *  bundler notes — the same asset-handling caveat). */
async function resolveAssetUrl(file: string, baseUrl?: URL | string): Promise<URL> {
  if (baseUrl) return new URL(file, baseUrl);
  if (typeof process !== "undefined" && process.versions?.node) {
    const { existsSync } = await import("node:fs");
    const { fileURLToPath } = await import("node:url");
    for (const candidate of ASSET_ROOT_CANDIDATES) {
      const url = relativeAssetUrl(candidate, file);
      if (existsSync(fileURLToPath(url))) return url;
    }
    return relativeAssetUrl(ASSET_ROOT_CANDIDATES[0], file);
  }
  let lastError: unknown;
  for (const candidate of ASSET_ROOT_CANDIDATES) {
    const url = relativeAssetUrl(candidate, file);
    try {
      const response = await fetch(url);
      if (response.ok) {
        // The caller re-fetches; this probe only settles the root. Browsers
        // cache the hit, so the cost is one extra 404 on miss.
        return url;
      }
    } catch (err) {
      lastError = err;
    }
  }
  throw new Error(
    `[@docen/layout] bundled fonts not found next to the package (${ASSET_ROOT_CANDIDATES.join(
      ", ",
    )}); pass baseUrl to registerDefaultFonts()`,
    { cause: lastError },
  );
}

async function readAsset(url: URL): Promise<Uint8Array> {
  if (typeof process !== "undefined" && process.versions?.node && url.protocol === "file:") {
    const { readFileSync } = await import("node:fs");
    const { fileURLToPath } = await import("node:url");
    return new Uint8Array(readFileSync(fileURLToPath(url)));
  }
  const response = await fetch(url);
  if (!response.ok) throw new Error(`[@docen/layout] failed to load bundled font ${url.href}`);
  return new Uint8Array(await response.arrayBuffer());
}

export interface RegisterDefaultFontsOptions {
  /** Base URL of the `assets/fonts` directory (browser bundles whose
   *  bundler does not emit `new URL` assets should set this). */
  readonly baseUrl?: URL | string;
  /** Override the bundled face list (tests / partial registration). */
  readonly faces?: readonly DefaultFontFace[];
}

/** One bundled face with its loaded bytes. */
export interface LoadedDefaultFontFace extends DefaultFontFace {
  readonly bytes: Uint8Array;
}

/** Load every bundled production face (bytes included). Consumers that
 *  register through their own font pipeline (the editor's export-embedding
 *  map) use this; `registerDefaultFonts` wraps it for the shaping manager. */
export async function loadDefaultFonts(
  options?: RegisterDefaultFontsOptions,
): Promise<LoadedDefaultFontFace[]> {
  const faces = options?.faces ?? DEFAULT_FONT_FACES;
  const bytesByFile = new Map<string, Uint8Array>();
  const loaded: LoadedDefaultFontFace[] = [];
  for (const face of faces) {
    let bytes = bytesByFile.get(face.file);
    if (!bytes) {
      bytes = await readAsset(await resolveAssetUrl(face.file, options?.baseUrl));
      bytesByFile.set(face.file, bytes);
    }
    loaded.push({ ...face, bytes });
  }
  return loaded;
}

/** Initialize the shaper and register the bundled production fonts. Returns
 *  the `family (style)` labels registered, in registration order.
 *
 *  Registering leaks no state across calls: a family registered twice simply
 *  replaces its slots (the latest bytes win), matching `registerFont`. */
export async function registerDefaultFonts(
  options?: RegisterDefaultFontsOptions,
): Promise<string[]> {
  await initShapingWasm();
  const registered: string[] = [];
  for (const face of await loadDefaultFonts(options)) {
    registerShapingFont(face.family, face.bytes, 0, {
      bold: face.bold,
      italic: face.italic,
    });
    registered.push(`${face.family}${face.bold ? " bold" : ""}${face.italic ? " italic" : ""}`);
  }
  return registered;
}
