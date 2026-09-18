// Lane D2 generation harness — items 4/5/13/17 end-to-end: a document covering
// preserved run elements, math, generated field caches and a from-scratch TOC
// runs gen1 → gen2 → gen3 through python-docx and LibreOffice headless
// (PDF + pdftotext). The per-construct assertions live in the focused specs
// (run-marker.spec.ts, math-tex.spec.ts, field-eval*.spec.ts); this harness
// proves the combined document stays editable across generations.
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { unzipSync } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, resolveDocument } from "../src";

const docx = (bytes: Uint8Array): Buffer => Buffer.from(bytes);
const xmlOf = (bytes: Uint8Array): string =>
  new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]);

function buildDoc(): DocumentOptions {
  return {
    sections: [
      {
        children: [
          { paragraph: { heading: "Heading1", children: ["Alpha Heading"] } },
          { paragraph: { heading: "Heading2", children: ["Sub Heading"] } },
          {
            paragraph: {
              children: [
                { bookmarkStart: { id: 1, name: "_Ref1" } },
                "Reference target",
                { bookmarkEnd: { id: 1 } },
              ],
            },
          },
          {
            paragraph: {
              children: [
                "See ",
                { simpleField: { instruction: " REF _Ref1 \\h ", cachedValue: "stale" } },
                " now",
              ],
            },
          },
          {
            paragraph: {
              style: "Caption",
              children: [
                "Figure ",
                { simpleField: { instruction: " SEQ Figure \\* ARABIC ", cachedValue: "9" } },
                ": One",
              ],
            },
          },
          {
            paragraph: {
              children: [
                "Page ",
                { complexField: { instruction: " PAGE ", result: "99" } },
                " / ",
                { complexField: { instruction: " NUMPAGES ", result: "99" } },
              ],
            },
          },
          {
            paragraph: {
              children: ["marks ", { children: [{ pgNum: true }, { dayShort: true }] }],
            },
          },
          {
            paragraph: {
              children: [
                {
                  math: {
                    children: [
                      { text: "\u03b1 + " },
                      { fraction: { numerator: ["a"], denominator: ["b"] } },
                    ],
                  },
                },
              ],
            },
          },
          {
            toc: {
              hyperlink: true,
              headingStyleRange: "1-3",
              useAppliedParagraphOutlineLevel: true,
              entries: [],
            },
          },
          { paragraph: { children: ["trailer"] } },
        ],
      },
    ],
  };
}

const PY_CHECK = `
import json, sys
import docx
from docx.oxml.ns import qn

path = sys.argv[1]
d = docx.Document(path)
paras = [p.text for p in d.paragraphs]
body = d.element.body
flds = {}
for fld in body.iter(qn("w:fldSimple")):
    instr = (fld.get(qn("w:instr")) or "").strip()
    text = "".join(t.text or "" for t in fld.iter(qn("w:t")))
    flds[instr] = text
print(json.dumps({
    "paragraphs": len(paras),
    "has_alpha": any("Alpha Heading" in p for p in paras),
    "seq": flds.get("SEQ Figure \\\\* ARABIC"),
    "ref": flds.get("REF _Ref1 \\\\h"),
}))
`;

const EXPECTED_TEXT = [
  "Alpha Heading",
  "Sub Heading",
  "Reference target",
  "Figure 1: One",
  "trailer",
  "\u03b1",
];

describe("D2 generation harness (gen1-3)", { timeout: 300000 }, () => {
  it("round-trips three generations, opens in python-docx and LibreOffice", () => {
    let python = false;
    try {
      python =
        execFileSync("python3", ["-c", "import docx; print('ok')"], {
          encoding: "utf8",
        }).trim() === "ok";
    } catch {
      console.log("⏭ python-docx missing — structural check skipped");
    }
    const dir = mkdtempSync(join(tmpdir(), "docen-d2-gen-"));
    try {
      const gen1 = generateDOCXSync(resolveDocument(buildDoc()), {
        prepare: false,
      }) as Uint8Array;
      const gen2 = generateDOCXSync(parseDOCXSync(gen1), { prepare: false }) as Uint8Array;
      const gen3 = generateDOCXSync(parseDOCXSync(gen2), { prepare: false }) as Uint8Array;
      writeFileSync(join(dir, "d2-gen1.docx"), docx(gen1));
      writeFileSync(join(dir, "d2-gen2.docx"), docx(gen2));
      writeFileSync(join(dir, "d2-gen3.docx"), docx(gen3));

      // Structure per generation: cached field results, TOC entries, math,
      // preserved run elements.
      for (const [name, bytes] of [
        ["gen1", gen1],
        ["gen2", gen2],
        ["gen3", gen3],
      ] as const) {
        const xml = xmlOf(bytes);
        expect(xml, `${name} REF`).toContain(">Reference target<");
        expect(xml, `${name} SEQ`).toContain(">1<");
        expect(xml, `${name} TOC entry`).toContain("Sub Heading");
        expect(xml, `${name} math`).toContain("<m:oMath");
        expect(xml, `${name} pgNum`).toContain("<w:pgNum/>");
        expect(xml, `${name} dayShort`).toContain("<w:dayShort/>");
        expect(xml, `${name} stale`).not.toContain("stale");
      }
      // gen2 → gen3 is byte-stable on the document part.
      expect(xmlOf(gen3)).toBe(xmlOf(gen2));

      if (python) {
        const py = join(dir, "check.py");
        writeFileSync(py, PY_CHECK);
        for (const name of ["gen1", "gen2", "gen3"]) {
          const out = execFileSync("python3", [py, join(dir, `d2-${name}.docx`)], {
            encoding: "utf8",
          });
          const report = JSON.parse(out) as Record<string, unknown>;
          expect(report.paragraphs, `${name} paragraphs`).toBeGreaterThan(5);
          expect(report.has_alpha, `${name} alpha`).toBe(true);
          expect(report.seq, `${name} seq`).toBe("1");
          expect(report.ref, `${name} ref`).toBe("Reference target");
        }
      }

      // LibreOffice headless PDF + pdftotext: every expected text in every gen.
      if (!existsSync("/usr/bin/soffice")) {
        console.log("⏭ soffice missing — LibreOffice oracle skipped");
        return;
      }
      const loDir = join(dir, "lo");
      mkdirSync(loDir, { recursive: true });
      for (const name of ["gen1", "gen2", "gen3"]) {
        execFileSync(
          "/usr/bin/soffice",
          [
            "--headless",
            `-env:UserInstallation=file://${join(loDir, "profile")}`,
            "--convert-to",
            "pdf",
            "--outdir",
            loDir,
            join(dir, `d2-${name}.docx`),
          ],
          { stdio: "ignore", timeout: 120000, env: { ...process.env, TMPDIR: loDir } },
        );
        const text = execFileSync("pdftotext", [join(loDir, `d2-${name}.pdf`), "-"], {
          encoding: "utf8",
        });
        const missing = EXPECTED_TEXT.filter((needle) => !text.includes(needle));
        console.log(
          `[lo] ${name}: ${EXPECTED_TEXT.length - missing.length}/${EXPECTED_TEXT.length} texts`,
        );
        expect(missing, `${name} missing`).toEqual([]);
      }
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
