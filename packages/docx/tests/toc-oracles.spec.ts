import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import {
  discoverTocFixtures,
  documentXmlOf,
  generateFixture,
  runPython,
  tocSlicesOf,
} from "./toc-driver";

/**
 * Item 8 fixture matrix, generation half: every fixture is generated with its
 * field/page context and checked against three independent oracles —
 * 1. the generated `word/document.xml` (field structure + switches),
 * 2. python-docx/lxml reading the same package (cached values, TOC entries),
 * 3. (the LO half, `toc-oracles-lo.spec.ts`) PDF rendering + resave + update.
 *
 * Fixtures are data: `tests/toc-fixtures/<case>/model.json` +
 * `expected.json` (description, field context, expectation sets). The python
 * extractor is `tests/toc-oracle.py`; it reads real OOXML, not our model.
 */
describe("toc-oracles generation matrix", { timeout: 120000 }, () => {
  const fixtures = discoverTocFixtures();

  it("discovers the full matrix", () => {
    expect(fixtures.map((fixture) => fixture.name)).toEqual([
      "header-footer-fields",
      "pageref-sections",
      "ref-crossrefs",
      "seq-chapters",
      "toc-bookmark-scope",
      "toc-custom-styles",
      "toc-empty",
      "toc-levels",
      "toc-multilevel",
      "toc-multiple",
      "toc-nested-in-table",
      "toc-preserved",
    ]);
  });

  it("matches every fixture's structural and python-docx expectations", () => {
    const dir = mkdtempSync(join(tmpdir(), "docen-toc-gen-"));
    try {
      const paths: string[] = [];
      const xmls = new Map<string, string>();
      for (const fixture of fixtures) {
        const name = fixture.name;
        const bytes = generateFixture(fixture);
        const path = join(dir, `${name}.docx`);
        writeFileSync(path, Buffer.from(bytes));
        paths.push(path);
        xmls.set(name, documentXmlOf(bytes));
      }
      const reports = runPython(paths);

      for (const fixture of fixtures) {
        const { name, expected } = fixture;
        const xml = xmls.get(name)!;
        const report = reports[join(dir, `${name}.docx`)]!;

        // 1. XML structure: required instructions/style markers present, and
        //    absent outside the TOC where the fixture says so.
        for (const needle of expected.xmlPresent ?? []) {
          expect(xml, `${name}: xml contains ${needle}`).toContain(needle);
        }
        for (const needle of expected.xmlAbsent ?? []) {
          expect(xml, `${name}: xml excludes ${needle}`).not.toContain(needle);
        }
        for (const needle of expected.tocSliceAbsent ?? []) {
          expect(tocSlicesOf(xml), `${name}: TOC excludes ${needle}`).not.toContain(needle);
        }

        // 2. python-docx/lxml structural read of the generated package.
        const python = expected.python;
        if (python?.simple)
          expect(report.simple, `${name}: fldSimple caches`).toEqual(python.simple);
        if (python?.instrTexts)
          expect(report.instrTexts, `${name}: instrTexts`).toEqual(python.instrTexts);
        if (python?.tocEntries)
          expect(report.tocEntries, `${name}: TOC entries`).toEqual(python.tocEntries);
        for (const [part, partExpected] of Object.entries(python?.parts ?? {})) {
          const partReport = report.parts[part];
          expect(partReport, `${name}: ${part} exists`).toBeDefined();
          if (partExpected.simple)
            expect(partReport!.simple, `${name}: ${part} fldSimple`).toEqual(partExpected.simple);
          if (partExpected.instrTexts)
            expect(partReport!.instrTexts, `${name}: ${part} instrTexts`).toEqual(
              partExpected.instrTexts,
            );
        }
      }
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
