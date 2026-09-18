// Systematic round-trip fuzz campaign (item 15).
//
// The generator (./generator.ts) builds documents across the audited feature
// matrix from checked-in seeds; every case runs the public pipeline
//   doc → generate(gen1) → parse → generate(gen2) → parse → generate(gen3)
// and must satisfy:
//   - no exception/hang anywhere (a case that throws fails the campaign)
//   - gen2 and gen3 are byte-identical (the stable point)
//   - parse(gen2) and parse(gen3) agree after volatile-field normalization
//   - generate(json) is deterministic (same input twice → same bytes)
// Failures are shrunk with a bounded delta-debugger whose predicate re-runs
// the synchronous pipeline on the reduced document; the minimal document,
// message and seed are written to tests/.temp/fuzz/failures/ and the spec
// fails. The external oracles (python-docx, LibreOffice PDF/resave, pdftotext)
// run over the same corpus via ../oracle/run.mjs.
//
// Budget: seeds.json holds the corpus; DOCEN_FUZZ_SAMPLES / DOCEN_FUZZ_BUDGET_MS
// override the per-run bounds (defaults keep the suite fast, the lane campaign
// runs the full list).

import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, resolveDocument } from "../../src";
import { firstHashDiff, partHashes, stableSerialize } from "../oracle/stability";
import { generateFuzzCase, shrinkCase } from "./generator";

const HERE = dirname(fileURLToPath(import.meta.url));
const OUT = join(HERE, "..", ".temp", "fuzz");
const SEEDS = JSON.parse(readFileSync(join(HERE, "seeds.json"), "utf8")) as {
  samples: number[];
  budgetMs: number;
  maxSamples: number;
};
const budgetMs = Number(process.env.DOCEN_FUZZ_BUDGET_MS ?? SEEDS.budgetMs);
const sampleCount = Number(process.env.DOCEN_FUZZ_SAMPLES ?? SEEDS.maxSamples);
const timeoutMs = Math.max(180_000, budgetMs * 6);

interface ChainResult {
  error: string | null;
  gens: [Uint8Array, Uint8Array, Uint8Array] | null;
}

/** The audited round-trip chain, synchronous so the shrinker can re-run it. */
function chainSync(doc: Record<string, unknown>): ChainResult {
  try {
    const source = resolveDocument(doc as never);
    const gen1 = generateDOCXSync(source, { prepare: false });
    const j1 = parseDOCXSync(gen1);
    const gen2 = generateDOCXSync(j1, { prepare: false });
    const j2 = parseDOCXSync(gen2);
    const gen3 = generateDOCXSync(j2, { prepare: false });
    const j3 = parseDOCXSync(gen3);

    const again = generateDOCXSync(j1, { prepare: false });
    if (!Buffer.from(again).equals(Buffer.from(gen2))) {
      return { error: "same JSON generated two different packages (determinism)", gens: null };
    }
    const diff = firstHashDiff(partHashes(gen2), partHashes(gen3));
    if (diff) {
      return {
        error: `gen2/gen3 differ in ${diff.part} (${diff.left} → ${diff.right})`,
        gens: null,
      };
    }
    if (!Buffer.from(gen2).equals(Buffer.from(gen3))) {
      return { error: "gen2/gen3 package bytes differ", gens: null };
    }
    if (stableSerialize(j2) !== stableSerialize(j3)) {
      return { error: "parse(gen2) and parse(gen3) models differ", gens: null };
    }
    return { error: null, gens: [gen1, gen2, gen3] };
  } catch (error) {
    return {
      error: error instanceof Error ? `${error.name}: ${error.message}` : String(error),
      gens: null,
    };
  }
}

describe("systematic editability fuzz", () => {
  it(
    "generated corpus round-trips stably within budget",
    () => {
      rmSync(OUT, { recursive: true, force: true });
      mkdirSync(join(OUT, "cases"), { recursive: true });
      const started = Date.now();
      const results: {
        seed: number;
        features: string[];
        stable: boolean;
        documentHash: string;
        ms: number;
      }[] = [];
      const failures: { seed: number; message: string; shrinkCalls: number }[] = [];

      for (const seed of SEEDS.samples.slice(0, sampleCount)) {
        if (results.length > 0 && Date.now() - started > budgetMs) break;
        const caseStarted = Date.now();
        const fuzzCase = generateFuzzCase(seed);
        const outcome = chainSync(fuzzCase.doc);

        const caseDir = join(OUT, "cases", String(seed));
        mkdirSync(caseDir, { recursive: true });
        writeFileSync(join(caseDir, "case.json"), `${JSON.stringify(fuzzCase, null, 2)}\n`);
        if (outcome.gens) {
          outcome.gens.forEach((bytes, index) => {
            writeFileSync(join(caseDir, `gen${index + 1}.docx`), bytes);
          });
        }

        if (outcome.error) {
          // Bounded delta-debugging on the exact failing pipeline.
          const shrunk = shrinkCase(fuzzCase.doc, (candidate) => chainSync(candidate).error);
          const failureDir = join(OUT, "failures");
          mkdirSync(failureDir, { recursive: true });
          writeFileSync(
            join(failureDir, `seed-${seed}.json`),
            `${JSON.stringify(
              {
                seed,
                features: fuzzCase.features,
                message: outcome.error,
                shrink: { calls: shrunk.calls, doc: shrunk.doc },
              },
              null,
              2,
            )}\n`,
          );
          const minimal = chainSync(shrunk.doc);
          writeFileSync(
            join(failureDir, `seed-${seed}-minimal.json`),
            `${JSON.stringify({ doc: shrunk.doc, error: minimal.error }, null, 2)}\n`,
          );
          failures.push({ seed, message: outcome.error, shrinkCalls: shrunk.calls });
        } else {
          results.push({
            seed,
            features: fuzzCase.features,
            stable: true,
            documentHash: partHashes(outcome.gens![1])["word/document.xml"] ?? "",
            ms: Date.now() - caseStarted,
          });
        }
      }

      writeFileSync(
        join(OUT, "manifest.json"),
        `${JSON.stringify(
          {
            generatedAt: new Date().toISOString(),
            budgetMs,
            elapsedMs: Date.now() - started,
            samples: results,
            failures,
          },
          null,
          2,
        )}\n`,
      );

      expect(
        failures,
        failures.map((entry) => `seed ${entry.seed}: ${entry.message}`).join("\n"),
      ).toEqual([]);
      expect(results.length).toBeGreaterThan(0);
    },
    timeoutMs,
  );
});
