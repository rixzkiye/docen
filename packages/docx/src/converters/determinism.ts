/**
 * Deterministic generation scope — the seam that makes two DOCX generations of
 * the same input byte-identical, in the same process and across processes.
 *
 * Office-open's generator has four nondeterminism sources: `wp:docPr` /
 * VML-shape / SmartArt-UUID counters, `crypto.randomUUID()` (font keys,
 * altChunk/subDoc part names), the core-properties/comment clock
 * (`new Date().toISOString()`), and the ZIP writer's per-entry mtime. The
 * upstream packages are patched (see `patches/@office-open__*.patch`) to read
 * this scope from `globalThis[Symbol.for("docen.ooxml.determinism")]` whenever
 * it is installed; outside a scope their original behavior is unchanged.
 *
 * {@link withGenerationScope} installs a fresh scope around the synchronous
 * compile phase of one generation. Upstream compiles all package parts before
 * its first `await`, so no two generations can interleave while a scope is
 * active even under `Promise.all` — each call sees its own counters.
 */

/** Symbol key of the installed generation scope (shared with the patches). */
export const GENERATION_SCOPE_KEY = Symbol.for("docen.ooxml.determinism");

/**
 * Fixed clock used when neither the source document nor the caller supplies a
 * date. The ZIP epoch (`1980-01-01T00:00:00.000Z`) is the conventional
 * reproducible-build timestamp; a fixed date also keeps `docProps/core.xml`
 * byte-stable across runs.
 */
export const DOCX_EPOCH = "1980-01-01T00:00:00.000Z";

/**
 * The scope object contract read by the patched upstream packages:
 * - `now`: ISO-8601 timestamp for every date default, or `undefined` to omit.
 * - `nextDocPr()`: `wp:docPr/@id` (Word starts at 1).
 * - `nextVml()`: VML `_x0000_s` shape id (Word hands out from 1025).
 * - `nextUuid()`: RFC 4122 v4-shaped UUID.
 * - `nextToken()`: 21-char lowercase alphanumeric id (altChunk/subDoc names).
 */
export interface GenerationScope {
  readonly now: string | undefined;
  nextDocPr(): number;
  nextVml(): number;
  nextUuid(): string;
  nextToken(): string;
}

/** Normalize a caller-supplied clock value to an ISO-8601 string. */
export function toIsoDate(value: string | Date): string {
  return value instanceof Date ? value.toISOString() : value;
}

function createScope(now: string | undefined): GenerationScope {
  let docPr = 0;
  let vml = 1024;
  let uuid = 0;
  let token = 0;
  return {
    now,
    nextDocPr: () => ++docPr,
    nextVml: () => ++vml,
    // v4-shaped, all-constant but the trailing 12-hex-digit sequence: valid
    // for consumers that only need an opaque unique GUID per node.
    nextUuid: () => `00000000-0000-4000-8000-${(++uuid).toString(16).padStart(12, "0")}`,
    nextToken: () => String(++token).padStart(21, "0"),
  };
}

/**
 * Run `fn` inside a fresh deterministic scope and restore the previous scope
 * afterwards. `fn` must contain the full synchronous compile phase of one
 * generation (see module docs); nested scopes are restored correctly.
 */
export function withGenerationScope<T>(now: string | undefined, fn: () => T): T {
  const global = globalThis as Record<symbol, unknown>;
  const previous = global[GENERATION_SCOPE_KEY];
  global[GENERATION_SCOPE_KEY] = createScope(now);
  try {
    return fn();
  } finally {
    if (previous === undefined) delete global[GENERATION_SCOPE_KEY];
    else global[GENERATION_SCOPE_KEY] = previous;
  }
}
