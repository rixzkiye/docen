import { unzipSync } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { ORDERED_REFERENCE_PREFIX, buildListLevels } from "../extensions/list-numbering";
import {
  compileDocument,
  docxExtensions,
  generateDOCXSync,
  parseDOCX,
  resolveDocument,
} from "../index";

/**
 * Compile-side persistence guarantees for the document-level attrs:
 *  - `numbering` round-trips whole (pic bullets, cleanup id, per-definition
 *    instances/aliases) — only abstractNumberings is rebuilt, for regenerated
 *    editor lists that must merge with the source definitions.
 *  - `documentExtras.rawParts` entries whose bytes a JSON round-trip already
 *    corrupted drop instead of reaching office-open's media reader, and the
 *    compile input stays untouched.
 */

function docWithAttrs(attrs: Record<string, unknown>): JSONContent {
  return {
    type: "doc",
    attrs,
    content: [{ type: "paragraph", content: [{ type: "text", text: "x" }] }],
  };
}

describe("compileDocument numbering persistence", () => {
  it("carries the full NumberingOptions through resolve→compile", () => {
    const doc: DocumentOptions = {
      sections: [{ children: [{ paragraph: { text: "x" } }] }],
      numbering: {
        abstractNumberings: [
          {
            reference: "list_1",
            levels: buildListLevels(`${ORDERED_REFERENCE_PREFIX}-1`)!,
            instanceCount: 2,
            aliases: ["2", "3"],
          },
        ],
        numIdMacAtCleanup: 9,
        numPicBullets: [{ numPicBulletId: 1, drawing: "<w:drawing/>" }],
      },
    };
    const json = resolveDocument(doc, docxExtensions);
    const compiled = compileDocument(json, docxExtensions);
    expect(compiled.numbering).toEqual(doc.numbering);
  });

  it("merges regenerated list definitions without dropping source numbering keys", () => {
    const sourceDefinition = { reference: "list_1", levels: buildListLevels("docen-bullet")! };
    const numPicBullets = [{ numPicBulletId: 7, drawing: "<w:drawing/>" }];
    const json = docWithAttrs({
      numbering: {
        abstractNumberings: [sourceDefinition],
        numIdMacAtCleanup: 4,
        numPicBullets,
      },
    });
    (json.content![0] as JSONContent).attrs = {
      numbering: { reference: `${ORDERED_REFERENCE_PREFIX}-1` },
    };

    const compiled = compileDocument(json, docxExtensions);
    const numbering = compiled.numbering!;
    expect(numbering.numPicBullets).toEqual(numPicBullets);
    expect(numbering.numIdMacAtCleanup).toBe(4);
    const references = numbering.abstractNumberings.map((c) => c.reference);
    expect(references).toContain("list_1");
    expect(references.some((r) => r.startsWith(ORDERED_REFERENCE_PREFIX))).toBe(true);
  });

  it("omits numbering when the JSON carries none", () => {
    const compiled = compileDocument({ type: "doc", content: [{ type: "paragraph" }] });
    expect("numbering" in compiled).toBe(false);
  });
});

describe("compileDocument rawParts sanitation", () => {
  const good = { path: "word/theme/theme1.xml", data: new Uint8Array([1, 2, 3]) };
  const text = { path: "customXml/item1.xml", data: "<x/>" };
  const corrupted = { path: "word/embeddings/ole.bin", data: { 0: 80, 1: 75 } };

  it("drops JSON-corrupted entries and never mutates the input attrs", () => {
    const json = docWithAttrs({ documentExtras: { rawParts: [good, corrupted, text] } });
    const compiled = compileDocument(json, docxExtensions);
    expect(compiled.rawParts).toEqual([good, text]);
    expect(json.attrs!.documentExtras).toEqual({ rawParts: [good, corrupted, text] });
  });

  it("drops the rawParts key once every entry is corrupted", () => {
    const json = docWithAttrs({ documentExtras: { rawParts: [corrupted] } });
    const compiled = compileDocument(json, docxExtensions);
    expect("rawParts" in compiled).toBe(false);
  });

  it("keeps a legal rawParts list verbatim", () => {
    const rawParts = [good, text];
    const compiled = compileDocument(docWithAttrs({ documentExtras: { rawParts } }));
    expect(compiled.rawParts).toBe(rawParts);
  });
});

/** [Content_Types].xml Override the packer must stamp per save variant. */
const MAIN_DOCUMENT_CONTENT_TYPES = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  docm: "application/vnd.ms-word.document.macroEnabled.main+xml",
  dotx: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
  dotm: "application/vnd.ms-word.template.macroEnabled.main+xml",
} as const;

/** The generated package's [Content_Types].xml, decoded for substring probes. */
function contentTypesXml(bytes: Uint8Array): string {
  const entry = unzipSync(bytes)["[Content_Types].xml"];
  expect(entry).toBeDefined();
  return new TextDecoder().decode(entry);
}

describe("generateDOCX variants", () => {
  const emptyDoc: JSONContent = { type: "doc", content: [{ type: "paragraph" }] };

  for (const variant of ["docx", "docm", "dotx", "dotm"] as const) {
    it(`stamps the ${variant} main document part content type`, () => {
      const xml = contentTypesXml(generateDOCXSync(emptyDoc, { variant }));
      expect(xml).toContain(
        `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPES[variant]}"/>`,
      );
    });
  }

  it("replaces a source document-type override while keeping other declarations", () => {
    const json = docWithAttrs({
      documentExtras: {
        contentTypes: {
          defaults: [{ extension: "bin", contentType: "application/vnd.ms-office.vbaProject" }],
          overrides: [
            { partName: "/word/document.xml", contentType: MAIN_DOCUMENT_CONTENT_TYPES.docm },
            { partName: "/word/styles.xml", contentType: "application/vnd.custom.styles+xml" },
          ],
        },
      },
    });
    const xml = contentTypesXml(generateDOCXSync(json, { variant: "dotx" }));
    expect(xml).toContain(
      `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPES.dotx}"/>`,
    );
    expect(xml).not.toContain(MAIN_DOCUMENT_CONTENT_TYPES.docm);
    expect(xml).toContain('ContentType="application/vnd.custom.styles+xml"');
    expect(xml).toContain('Extension="bin" ContentType="application/vnd.ms-office.vbaProject"');
  });
});

describe("docm macro-part round-trip", () => {
  const VBA = new Uint8Array(Array.from({ length: 96 }, (_, i) => (i * 13 + 7) % 256));
  const UNKNOWN = new Uint8Array(Array.from({ length: 24 }, (_, i) => 255 - i));

  /** A source .docm whose package carries a dummy vbaProject plus an unknown
   *  part, exactly the shape office-open passes through as rawParts. */
  function sourceDocm(): Uint8Array {
    return generateDOCXSync(
      {
        type: "doc",
        content: [{ type: "paragraph", content: [{ type: "text", text: "macro" }] }],
        attrs: {
          documentExtras: {
            rawParts: [
              {
                path: "word/vbaProject.bin",
                data: VBA,
                contentType: "application/vnd.ms-office.vbaProject",
              },
              {
                path: "customXml/unknown.dat",
                data: UNKNOWN,
                contentType: "application/octet-stream",
              },
            ],
          },
        },
      },
      {
        variant: "docm",
        document: {
          contentTypes: {
            defaults: [{ extension: "bin", contentType: "application/vnd.ms-office.vbaProject" }],
            overrides: [],
          },
        },
      },
    );
  }

  it("keeps vbaProject.bin and unknown parts byte-identical through open→save", async () => {
    const opened = await parseDOCX(sourceDocm());
    // The parse captured the macro part in the passthrough set.
    const extras = opened.attrs!.documentExtras as {
      rawParts?: Array<{ path: string; data: Uint8Array }>;
    };
    expect(extras.rawParts?.some((part) => part.path === "word/vbaProject.bin")).toBe(true);
    const saved = generateDOCXSync(opened, { variant: "docm" });
    const zip = unzipSync(saved);
    expect(zip["word/vbaProject.bin"]).toEqual(VBA);
    expect(zip["customXml/unknown.dat"]).toEqual(UNKNOWN);
    // The macro-enabled main content type survives the round-trip.
    const xml = new TextDecoder().decode(zip["[Content_Types].xml"]);
    expect(xml).toContain(MAIN_DOCUMENT_CONTENT_TYPES.docm);
  });

  it("keeps the parts through a plain save (no variant requested)", async () => {
    const opened = await parseDOCX(sourceDocm());
    const zip = unzipSync(generateDOCXSync(opened));
    expect(zip["word/vbaProject.bin"]).toEqual(VBA);
    const xml = new TextDecoder().decode(zip["[Content_Types].xml"]);
    expect(xml).toContain(MAIN_DOCUMENT_CONTENT_TYPES.docm);
  });
});
