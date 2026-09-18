// DOCX oracle fixtures — items 6 (footnote/endnote) and 7 (headers/footers/
// page numbers) of the R8 audit, plus the shared 3-generation stability
// invariants every fixture must satisfy.
//
// Each fixture is a directory:
//   model.json  DocumentOptions input (resolved to runtime JSON first), or
//   input.docx  a foreign-writer package parsed with parseDOCX, and
//   checks.json declarative expectations consumed here (xml) and by run.mjs
//               (python structure signals + LibreOffice/pdftotext pages).
//
// The harness generates gen1 (from the fixture), then two parse→generate
// cycles (gen2, gen3) and asserts:
//   - byte identity of gen2 and gen3 (roundTrip(roundTrip(x)) == roundTrip(x))
//   - generation determinism (same JSON twice → same bytes)
//   - model equality of parse(gen2) / parse(gen3) after volatile-field
//     normalization
//   - every `xml` expectation on every generation
//   - no exception anywhere (a throw fails the fixture)
//
// run.mjs (`pnpm --filter @docen/docx test:oracles`) consumes the gen*.docx
// files written to tests/.temp/oracle and applies the independent
// python-docx, LibreOffice (PDF + resave) and pdftotext oracles.

import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { generateDOCX, parseDOCX, resolveDocument } from "../../src";
import { firstHashDiff, partHashes, stableSerialize, xmlParts } from "./stability";

const HERE = dirname(fileURLToPath(import.meta.url));
const FIXTURES = join(HERE, "fixtures");
const OUT = join(HERE, "..", ".temp", "oracle");
const TIMEOUT_MS = 60_000;

interface XmlCheck {
  part: string;
  contains?: string[];
  notContains?: string[];
  order?: [string, string][];
  regex?: string[];
}

interface SignalCheck {
  path: string;
  equals?: unknown;
  contains?: unknown;
  length?: number;
  truthy?: boolean;
}

interface PdfCheck {
  page: number;
  contains?: string[];
  notContains?: string[];
  regex?: string[];
}

interface ChecksFile {
  description?: string;
  /** Compare gen1/gen2/gen3 page text; a known-normalization fixture may
   *  disable it, but every fixture currently keeps it on. */
  pdfStable?: boolean;
  /** Parse the LibreOffice resave with docen and compare its PDF text. */
  resaveStable?: boolean;
  xml?: XmlCheck[];
  python?: SignalCheck[];
  pdf?: PdfCheck[];
}

const CHECK_KEYS = new Set(["description", "pdfStable", "resaveStable", "xml", "python", "pdf"]);

function loadChecks(path: string): ChecksFile {
  const checks = JSON.parse(readFileSync(path, "utf8")) as ChecksFile;
  for (const key of Object.keys(checks)) {
    if (!CHECK_KEYS.has(key)) throw new Error(`${path}: unknown checks key "${key}"`);
  }
  for (const check of checks.xml ?? []) {
    if (!check.part || typeof check.part !== "string") throw new Error(`${path}: xml.part missing`);
  }
  for (const check of checks.python ?? []) {
    if (!check.path || typeof check.path !== "string") {
      throw new Error(`${path}: python.path missing`);
    }
    const operators = ["equals", "contains", "length", "truthy"].filter(
      (key) => (check as unknown as Record<string, unknown>)[key] !== undefined,
    );
    if (operators.length !== 1) {
      throw new Error(`${path}: python check needs exactly one operator: ${JSON.stringify(check)}`);
    }
  }
  for (const check of checks.pdf ?? []) {
    if (!Number.isInteger(check.page) || check.page < 1) {
      throw new Error(`${path}: pdf.page must be a positive integer`);
    }
  }
  return checks;
}

function parseXmlExpectation(check: XmlCheck, text: string): string | undefined {
  for (const needle of check.contains ?? []) {
    if (!text.includes(needle)) return `missing ${JSON.stringify(needle)}`;
  }
  for (const needle of check.notContains ?? []) {
    if (text.includes(needle)) return `unexpected ${JSON.stringify(needle)}`;
  }
  for (const [before, after] of check.order ?? []) {
    const beforeAt = text.indexOf(before);
    const afterAt = text.indexOf(after);
    if (beforeAt < 0) return `missing ${JSON.stringify(before)} (order)`;
    if (afterAt < 0) return `missing ${JSON.stringify(after)} (order)`;
    if (beforeAt > afterAt)
      return `${JSON.stringify(before)} must precede ${JSON.stringify(after)}`;
  }
  for (const pattern of check.regex ?? []) {
    if (!new RegExp(pattern).test(text)) return `no match for /${pattern}/`;
  }
  return undefined;
}

const fixtures = readdirSync(FIXTURES, { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .map((entry) => entry.name)
  .sort();

describe("docx oracle fixtures", () => {
  it("ships a fixture matrix for note and header/footer edge coverage", () => {
    expect(fixtures.length).toBeGreaterThanOrEqual(12);
    for (const name of fixtures) {
      const dir = join(FIXTURES, name);
      const model = existsSync(join(dir, "model.json"));
      const input = existsSync(join(dir, "input.docx"));
      expect(model !== input, `${name}: exactly one of model.json / input.docx`).toBe(true);
      expect(existsSync(join(dir, "checks.json")), `${name}: checks.json`).toBe(true);
    }
  });

  for (const name of fixtures) {
    it(
      `${name}: 3-generation stability + XML checks`,
      async () => {
        const dir = join(FIXTURES, name);
        const checks = loadChecks(join(dir, "checks.json"));
        const modelPath = join(dir, "model.json");
        const inputPath = join(dir, "input.docx");

        // gen1 derives from the fixture (resolved model or parsed foreign DOCX).
        const source = existsSync(modelPath)
          ? resolveDocument(JSON.parse(readFileSync(modelPath, "utf8")))
          : await parseDOCX(readFileSync(inputPath));
        const gen1 = await generateDOCX(source, { prepare: false });

        // The two audited round-trips.
        const j1 = await parseDOCX(gen1);
        const gen2 = await generateDOCX(j1, { prepare: false });
        const j2 = await parseDOCX(gen2);
        const gen3 = await generateDOCX(j2, { prepare: false });
        const j3 = await parseDOCX(gen3);

        // Same JSON in → same bytes out (determinism under the fixed clock).
        const again = await generateDOCX(j1, { prepare: false });
        expect(Buffer.from(again).equals(Buffer.from(gen2)), `${name}: gen2 is deterministic`).toBe(
          true,
        );

        // Stable point: gen2 and gen3 must be byte-identical, part for part.
        const h2 = partHashes(gen2);
        const h3 = partHashes(gen3);
        const diff = firstHashDiff(h2, h3);
        expect(
          diff,
          `${name}: gen3 differs from gen2 in ${diff?.part ?? "?"} (${diff?.left ?? "?"} → ${diff?.right ?? "?"})`,
        ).toBeUndefined();
        expect(
          Buffer.from(gen2).equals(Buffer.from(gen3)),
          `${name}: gen2/gen3 package bytes are identical`,
        ).toBe(true);

        // Model equality after volatile-field normalization.
        expect(stableSerialize(j2), `${name}: parsed gen2 model == gen3 model`).toBe(
          stableSerialize(j3),
        );

        // XML expectations hold on every generation (round-trip preserved them).
        const gens = [gen1, gen2, gen3];
        for (let i = 0; i < gens.length; i++) {
          const parts = xmlParts(gens[i]!);
          for (const check of checks.xml ?? []) {
            const text = parts[check.part!];
            expect(text, `${name}: gen${i + 1} lacks ${check.part}`).toBeDefined();
            const failure = parseXmlExpectation(check, text!);
            expect(failure, `${name}: gen${i + 1} ${check.part}: ${failure ?? ""}`).toBeUndefined();
          }
        }

        // Persist the generated packages for the external oracle phase (run.mjs).
        rmSync(join(OUT, name), { recursive: true, force: true });
        mkdirSync(join(OUT, name), { recursive: true });
        for (const [index, bytes] of gens.entries()) {
          writeFileSync(join(OUT, name, `gen${index + 1}.docx`), bytes);
        }
        writeFileSync(
          join(OUT, name, "meta.json"),
          `${JSON.stringify(
            {
              name,
              mode: existsSync(modelPath) ? "model" : "input",
              description: checks.description ?? "",
              pdfStable: checks.pdfStable ?? true,
              resaveStable: checks.resaveStable ?? true,
              sha256: gens.map((bytes) => partHashes(bytes)["word/document.xml"] ?? ""),
            },
            null,
            2,
          )}\n`,
        );
      },
      TIMEOUT_MS,
    );
  }
});
