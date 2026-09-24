// @vitest-environment node
// Regression (W23): headless PDF PAGE fields in section furniture must paint
// the section-local number (w:pgNumType format + start), not the global
// physical page index — the academic front-matter (lowerRoman) / body (decimal)
// shape. The independent oracle is pdftotext per page: before the fix the scene
// painted the physical index while the text layer painted the section number,
// so extraction returned both ("i1" on the first front-matter page) and the
// body pages showed 3/4 instead of 1/2.

import { execFileSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { describe, expect, it } from "vitest";

import { renderPdf } from "./render";

/** A PAGE field atom (the furniture band the presentation layer emits). */
const pageField = (): Record<string, unknown> => ({
  type: "inlinePassthrough",
  attrs: { data: JSON.stringify({ simpleField: { instruction: "PAGE" } }) },
});

/** A centered footer band carrying only the PAGE field. */
const pageFooter = (): Record<string, unknown> => ({
  type: "paragraph",
  attrs: { alignment: "center" },
  content: [pageField()],
});

/** Front matter (lowerRoman, restart at 1) over two pages, then body (decimal,
 *  restart at 1) over two pages — the section break closes the front section on
 *  "Beta" and the final section is the document-level one. */
const numberedSections = (): Record<string, unknown> => ({
  type: "doc",
  attrs: {
    sectionProperties: { pageNumberType: { format: "decimal", start: 1 } },
    sectionFooters: { default: [pageFooter()] },
  },
  content: [
    { type: "paragraph", content: [{ type: "text", text: "Alpha" }] },
    {
      type: "paragraph",
      attrs: {
        pageBreakBefore: true,
        sectionProperties: { pageNumberType: { format: "lowerRoman", start: 1 } },
        sectionFooters: { default: [pageFooter()] },
      },
      content: [{ type: "text", text: "Beta" }],
    },
    { type: "paragraph", content: [{ type: "text", text: "Gamma" }] },
    {
      type: "paragraph",
      attrs: { pageBreakBefore: true },
      content: [{ type: "text", text: "Delta" }],
    },
  ],
});

/** Per-page pdftotext extraction, whitespace-collapsed for exact comparison. */
function pdfPages(bytes: Uint8Array): string[] {
  const tmp = path.join(os.tmpdir(), `docen-page-numbers-${process.pid}-${Date.now()}.pdf`);
  fs.writeFileSync(tmp, bytes);
  try {
    return execFileSync("pdftotext", [tmp, "-"], { encoding: "utf-8" })
      .split("\f")
      .map((page) => page.replace(/\s+/g, " ").trim())
      .filter((page) => page.length > 0);
  } finally {
    fs.rmSync(tmp, { force: true });
  }
}

describe("headless PDF PAGE fields use the section page-number context", () => {
  it("paints the section-local number per page in the PDF/A lane", async () => {
    const bytes = await renderPdf(numberedSections(), {
      pdfa: "2b",
      pdfUa: true,
      title: "Page Numbers",
    });
    expect(pdfPages(bytes)).toEqual(["Alpha i", "Beta ii", "Gamma 1", "Delta 2"]);
  }, 60000);

  it("keeps the scene and text layers agreeing in the default outlines lane", async () => {
    const bytes = await renderPdf(numberedSections(), { title: "Page Numbers" });
    expect(pdfPages(bytes)).toEqual(["Alpha i", "Beta ii", "Gamma 1", "Delta 2"]);
  }, 60000);
});
