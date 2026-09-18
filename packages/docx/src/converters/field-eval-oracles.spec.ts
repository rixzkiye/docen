import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import type { JSONContent } from "../core";
import { generateDOCXSync } from "../index";

/**
 * External oracles for items 13/17 (kept in their own file: they shell out to
 * python-docx and LibreOffice, so the pure field-cache spec stays fast).
 */

const para = (children: JSONContent[], attrs?: Record<string, unknown>): JSONContent => ({
  type: "paragraph",
  ...(attrs ? { attrs } : {}),
  content: children,
});

const text = (value: string): JSONContent => ({ type: "text", text: value });

const field = (branch: Record<string, unknown>): JSONContent => ({
  type: "inlinePassthrough",
  attrs: { data: JSON.stringify(branch) },
});

const simple = (instruction: string, cachedValue?: string): JSONContent =>
  field({ simpleField: { instruction, ...(cachedValue !== undefined ? { cachedValue } : {}) } });

const bookmarkStart = (id: number, name: string): JSONContent =>
  field({ bookmarkStart: { id, name } });
const bookmarkEnd = (id: number): JSONContent => field({ bookmarkEnd: { id } });

const docOf = (content: JSONContent[]): JSONContent => ({ type: "doc", content });

const PY_CHECK = `
import json, sys
import docx
from docx.oxml.ns import qn

path = sys.argv[1]
d = docx.Document(path)
body = d.element.body
seq, ref = [], []
for fld in body.iter(qn("w:fldSimple")):
    instr = fld.get(qn("w:instr")) or ""
    text = "".join(t.text or "" for t in fld.iter(qn("w:t")))
    if "SEQ" in instr:
        seq.append(text)
    if "REF" in instr:
        ref.append(text)
print(json.dumps({"seq": seq, "ref": ref}))
`;

describe("python-docx structural check", { timeout: 120000 }, () => {
  it("reads the generated SEQ/REF caches from the package", () => {
    let python = "";
    try {
      python = execFileSync("python3", ["-c", "import docx; print('ok')"], {
        encoding: "utf8",
      });
    } catch {
      console.log("⏭ python-docx missing — structural check skipped");
      return;
    }
    expect(python.trim()).toBe("ok");
    const dir = mkdtempSync(join(tmpdir(), "docen-fields-py-"));
    try {
      const json = docOf([
        para([text("A")], { heading: "Heading1" }),
        para([simple(" SEQ Figure \\* ARABIC ", "9")]),
        para([bookmarkStart(1, "_Ref1"), text("Alpha"), bookmarkEnd(1)]),
        para([simple(" REF _Ref1 \\h ", "stale")]),
      ]);
      const docxPath = join(dir, "fields.docx");
      writeFileSync(
        docxPath,
        Buffer.from(generateDOCXSync(json, { prepare: false }) as Uint8Array),
      );
      const script = join(dir, "check.py");
      writeFileSync(script, PY_CHECK);
      const out = execFileSync("python3", [script, docxPath], { encoding: "utf8" });
      expect(JSON.parse(out)).toEqual({ seq: ["1"], ref: ["Alpha"] });
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe("LibreOffice field text oracle", { timeout: 180000 }, () => {
  it("renders the TOC entries, reference text and page numbers", () => {
    if (!existsSync("/usr/bin/soffice")) {
      console.log("⏭ soffice missing — LibreOffice oracle skipped");
      return;
    }
    const dir = mkdtempSync(join(tmpdir(), "docen-fields-lo-"));
    try {
      const json = docOf([
        para([text("Alpha heading")], { heading: "Heading1" }),
        para([bookmarkStart(1, "_Ref1"), text("Reference target"), bookmarkEnd(1)]),
        para([simple(" REF _Ref1 \\h ", "stale")]),
        para([simple(" SEQ Figure \\* ARABIC ", "9")]),
        para([text("Page "), simple(" PAGE ", "9"), text(" / "), simple(" NUMPAGES ", "9")]),
        {
          type: "tocField",
          attrs: { options: { headingStyleRange: "1-3" } },
          content: [{ type: "paragraph" }],
        },
      ]);
      const docxPath = join(dir, "fields.docx");
      writeFileSync(
        docxPath,
        Buffer.from(generateDOCXSync(json, { prepare: false }) as Uint8Array),
      );
      const env = { ...process.env, TMPDIR: dir };
      execFileSync(
        "/usr/bin/soffice",
        [
          "--headless",
          `-env:UserInstallation=file://${join(dir, "lo-profile")}`,
          "--convert-to",
          "pdf",
          "--outdir",
          dir,
          docxPath,
        ],
        { stdio: "ignore", timeout: 90000, env },
      );
      const pdfText = execFileSync("pdftotext", [join(dir, "fields.pdf"), "-"], {
        encoding: "utf8",
      });
      // The TOC cached result renders (its heading is the entry text).
      expect(pdfText).toContain("Alpha heading");
      // The REF renders its target text (stale cache absent).
      expect(pdfText).toContain("Reference target");
      expect(pdfText).not.toContain("stale");
      // The SEQ number renders and the page fields resolve.
      expect(pdfText).toMatch(/Page\s+1\s*\/\s*1/);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
