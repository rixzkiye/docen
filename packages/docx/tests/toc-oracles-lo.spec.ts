import { execFileSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import type { JSONContent } from "../src/core";
import { parseDOCXSync } from "../src/index";
import {
  discoverTocFixtures,
  expectationMissing,
  generateFixture,
  haveSoffice,
  runPython,
  tocEntriesMissing,
  type ExpectedTocEntry,
  type PythonExpectation,
  type PythonReport,
} from "./toc-driver";

/**
 * Item 8 fixture matrix, LibreOffice half — the reader's view plus field-update
 * behavior:
 * 1. plain headless PDF (`--convert-to pdf`): pdftotext must show the cached
 *    TOC entries/page numbers, SEQ/REF values and page fields the reader sees
 *    WITHOUT a field update;
 * 2. resave (`--convert-to docx`, general-purpose bit-3 ZIP entries): the
 *    package must re-import (`parseDOCXSync`) with the fields still updatable
 *    and the cached text stable (python-docx check on the resave);
 * 3. explicit update over a UNO socket (update all text fields + document
 *    indexes, then PDF/DOCX): assert the supported subset does not drift and
 *    record the LO limitations per fixture in `expected.json`
 *    (`loLimitations`).
 */

const PORT = 3217;
const fixtures = discoverTocFixtures();

function findTocFields(node: JSONContent, out: JSONContent[] = []): JSONContent[] {
  if (node.type === "tocField") out.push(node);
  for (const child of node.content ?? []) findTocFields(child, out);
  return out;
}

describe("toc-oracles LibreOffice matrix", { timeout: 900000 }, () => {
  it("renders, resaves and explicitly updates every fixture", () => {
    if (!haveSoffice()) {
      console.log("⏭ soffice missing — LibreOffice oracle skipped");
      return;
    }
    const dir = mkdtempSync(join(tmpdir(), "docen-toc-lo-"));
    const pdfDir = join(dir, "pdf");
    const resaveDir = join(dir, "resave");
    const updateDir = join(dir, "update");
    mkdirSync(pdfDir, { recursive: true });
    mkdirSync(resaveDir, { recursive: true });
    mkdirSync(updateDir, { recursive: true });

    const stopLo = (): void => {
      try {
        execFileSync("pkill", ["-f", `${dir}/uno-profile`]);
      } catch {
        /* already gone */
      }
    };
    try {
      const names: string[] = [];
      const docs: string[] = [];
      for (const fixture of fixtures) {
        const path = join(dir, `${fixture.name}.docx`);
        writeFileSync(path, Buffer.from(generateFixture(fixture)));
        names.push(fixture.name);
        docs.push(path);
      }

      // 1. Plain headless PDF.
      execFileSync(
        "/usr/bin/soffice",
        [
          "--headless",
          `-env:UserInstallation=file://${join(dir, "pdf-profile")}`,
          "--convert-to",
          "pdf",
          "--outdir",
          pdfDir,
          ...docs,
        ],
        { stdio: "ignore", timeout: 300_000, env: { ...process.env, TMPDIR: dir } },
      );
      for (const fixture of fixtures) {
        const text = execFileSync("pdftotext", [join(pdfDir, `${fixture.name}.pdf`), "-"], {
          encoding: "utf8",
        });
        const pdf = fixture.expected.pdf;
        if (!pdf) continue;
        for (const needle of pdf.contains ?? []) {
          expect(text, `${fixture.name}: pdf contains ${needle}`).toContain(needle);
        }
        for (const pattern of pdf.patterns ?? []) {
          expect(text, `${fixture.name}: pdf matches ${pattern}`).toMatch(new RegExp(pattern));
        }
        for (const needle of pdf.absent ?? []) {
          expect(text, `${fixture.name}: pdf excludes ${needle}`).not.toContain(needle);
        }
      }

      // 2. Resave (LO writes data-descriptor ZIP entries — the archive guard
      //    must admit them) + python-docx stability + docen re-import.
      execFileSync(
        "/usr/bin/soffice",
        [
          "--headless",
          `-env:UserInstallation=file://${join(dir, "pdf-profile")}`,
          "--convert-to",
          "docx",
          "--outdir",
          resaveDir,
          ...docs,
        ],
        { stdio: "ignore", timeout: 300_000, env: { ...process.env, TMPDIR: dir } },
      );
      const resavePaths = names.map((name) => join(resaveDir, `${name}.docx`));
      const resaveReports = runPython(resavePaths);
      for (const fixture of fixtures) {
        const path = join(resaveDir, `${fixture.name}.docx`);
        const parsed = parseDOCXSync(new Uint8Array(readFileSync(path)));
        const tocs = findTocFields(parsed);
        if ((fixture.expected.xmlPresent ?? []).some((needle) => needle.includes("TOC"))) {
          expect(tocs.length, `${fixture.name}: resave keeps a TOC field`).toBeGreaterThan(0);
        }
        const report = resaveReports[path]!;
        checkResaveReport(fixture.name, fixture.expected.resavePython, report);
      }

      // 3. Explicit update (fresh socket; restart LO when a fixture crashes it).
      const startLo = (): void => {
        execFileSync("bash", [
          "-c",
          `setsid nohup /usr/bin/soffice --headless --norestore --nologo ` +
            `-env:UserInstallation=file://${dir}/uno-profile ` +
            `--accept="socket,host=127.0.0.1,port=${PORT};urp;" >${dir}/uno.log 2>&1 < /dev/null &`,
        ]);
        execFileSync("sleep", ["7"]);
      };
      const updateOnce = (name: string): string =>
        execFileSync(
          "python3",
          [
            join(import.meta.dirname, "toc-oracle-update.py"),
            join(dir, `${name}.docx`),
            join(updateDir, `${name}.pdf`),
            join(updateDir, `${name}.docx`),
            "--port",
            String(PORT),
          ],
          { encoding: "utf8", timeout: 120_000 },
        ).trim();
      startLo();
      try {
        const updatedPaths: string[] = [];
        for (const fixture of fixtures) {
          const update = fixture.expected.update;
          if (update?.skip) {
            console.log(`⏭ ${fixture.name}: explicit update skipped — ${update.skip}`);
            continue;
          }
          try {
            updateOnce(fixture.name);
          } catch {
            // A fixture that crashes LO (table-nested TOC) gets one restart,
            // then is reported rather than failing the whole matrix.
            stopLo();
            startLo();
            try {
              updateOnce(fixture.name);
            } catch (retry) {
              throw new Error(
                `${fixture.name}: LO explicit update failed: ${String(retry).slice(0, 200)}`,
              );
            }
          }
          const text = execFileSync("pdftotext", [join(updateDir, `${fixture.name}.pdf`), "-"], {
            encoding: "utf8",
          });
          for (const needle of update?.contains ?? []) {
            expect(text, `${fixture.name}: updated pdf contains ${needle}`).toContain(needle);
          }
          for (const pattern of update?.patterns ?? []) {
            expect(text, `${fixture.name}: updated pdf matches ${pattern}`).toMatch(
              new RegExp(pattern),
            );
          }
          updatedPaths.push(join(updateDir, `${fixture.name}.docx`));
        }
        // The updated package must stay parseable/updatable too.
        const updateReports = runPython(updatedPaths);
        for (const fixture of fixtures) {
          const path = join(updateDir, `${fixture.name}.docx`);
          if (fixture.expected.update?.skip) continue;
          const parsed = parseDOCXSync(new Uint8Array(readFileSync(path)));
          void parsed;
          checkResaveReport(
            fixture.name,
            fixture.expected.updatePython ?? fixture.expected.resavePython,
            updateReports[path]!,
          );
        }
      } finally {
        stopLo();
      }
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

/** Apply the fixture's post-resave python expectations (substring/subset — LO
 *  rewrites some instructions and styles, the cached values must survive). */
function checkResaveReport(
  name: string,
  expected: PythonExpectation | undefined,
  report: PythonReport,
): void {
  if (!expected) return;
  if (expected.instrTexts) {
    expect(
      expectationMissing(expected.instrTexts, report.instrTexts),
      `${name}: resave fields`,
    ).toEqual([]);
  }
  if (expected.tocEntries) {
    expect(
      tocEntriesMissing(expected.tocEntries as ExpectedTocEntry[], report.tocEntries),
      `${name}: resave TOC caches`,
    ).toEqual([]);
  }
  const unionParts = (pattern: RegExp): string[] =>
    Object.entries(report.parts)
      .filter(([part]) => pattern.test(part))
      .flatMap(([, part]) => part.instrTexts);
  for (const [key, value] of Object.entries({
    header: expected.headers,
    footer: expected.footers,
  })) {
    if (!value?.instrTexts) continue;
    const union = unionParts(new RegExp(`word/${key}`));
    expect(expectationMissing(value.instrTexts, union), `${name}: resave ${key}s`).toEqual([]);
  }
}
