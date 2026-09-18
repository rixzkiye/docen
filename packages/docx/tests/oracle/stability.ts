// Round-trip stability helpers for the oracle/fuzz harnesses.
//
// A generated DOCX is stable when a second parse→generate cycle is
// byte-identical to the first (the audit's `roundTrip(roundTrip(x)) ==
// roundTrip(x)` invariant). These helpers provide the deterministic
// serialization and part hashing both harness phases compare.

import { createHash } from "node:crypto";

import { unzipSync } from "fflate";

/** Keys whose values are generation-time bookkeeping rather than content.
 *  They legitimately change between generations and are excluded from the
 *  model-equality oracle (byte equality of gen2/gen3 is the strict check). */
const VOLATILE_KEYS = new Set([
  "paraId",
  "textId",
  "rsids",
  "rawParts",
  "documentAttributes",
  "defaultNamespace",
  "contentTypes",
  "passthroughRelationships",
]);

/** JSON-safe, key-sorted serialization of a parsed model with volatile
 *  bookkeeping removed and bytes base64-encoded. */
export function stableSerialize(value: unknown): string {
  return JSON.stringify(stableValue(value));
}

function stableValue(value: unknown): unknown {
  if (value instanceof Uint8Array) {
    return { $bytes: Buffer.from(value).toString("base64") };
  }
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(stableValue);
  if (value && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(value).sort()) {
      if (VOLATILE_KEYS.has(key) || key.startsWith("w14") || /^rsid/i.test(key)) continue;
      const entry = stableValue((value as Record<string, unknown>)[key]);
      if (entry !== undefined) out[key] = entry;
    }
    return out;
  }
  return value;
}

/** SHA-256 per ZIP part — the byte-level identity of a generated package. */
export function partHashes(bytes: Uint8Array): Record<string, string> {
  const parts = unzipParts(bytes);
  const out: Record<string, string> = {};
  for (const name of Object.keys(parts).sort()) {
    out[name] = createHash("sha256").update(parts[name]!).digest("hex");
  }
  return out;
}

/** First differing part between two hash maps (undefined when identical). */
export function firstHashDiff(
  a: Record<string, string>,
  b: Record<string, string>,
): { part: string; left?: string; right?: string } | undefined {
  for (const part of [...new Set([...Object.keys(a), ...Object.keys(b)])].sort()) {
    if (a[part] !== b[part]) return { part, left: a[part], right: b[part] };
  }
  return undefined;
}

function unzipParts(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipSync(bytes);
}

/** All XML/rels text parts of a generated package, decoded. */
export function xmlParts(bytes: Uint8Array): Record<string, string> {
  const decoder = new TextDecoder();
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipParts(bytes))) {
    if (name.endsWith(".xml") || name.endsWith(".rels")) out[name] = decoder.decode(data);
  }
  return out;
}

/** Normalized PDF text comparison: whitespace collapsed, form feeds dropped. */
export function normalizePdfText(text: string): string {
  return text
    .replace(/\f/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\s*\n\s*/g, "\n")
    .trim();
}
