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
  parseDOCXSync,
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

const PAGE = {
  pageSize: { width: 11906, height: 16838 },
  pageMargin: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
};

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

  it("flips docx→dotx→docx back to the standard document main type", async () => {
    const asDotx = generateDOCXSync(
      await parseDOCX(generateDOCXSync(emptyDoc, { variant: "docx" })),
      { variant: "dotx" },
    );
    expect(contentTypesXml(asDotx)).toContain(MAIN_DOCUMENT_CONTENT_TYPES.dotx);
    const back = generateDOCXSync(await parseDOCX(asDotx), { variant: "docx" });
    const xml = contentTypesXml(back);
    expect(xml).toContain(
      `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPES.docx}"/>`,
    );
    expect(xml).not.toContain("template.main+xml");
  });

  it("flips an opened docm package when docx is requested explicitly", async () => {
    const xml = contentTypesXml(
      generateDOCXSync(await parseDOCX(sourceDocm()), { variant: "docx" }),
    );
    expect(xml).toContain(
      `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPES.docx}"/>`,
    );
    expect(xml).not.toContain("macroEnabled.main+xml");
  });

  it("flips a dotx input package when docx is requested explicitly", async () => {
    const dotx = generateDOCXSync(emptyDoc, { variant: "dotx" });
    const xml = contentTypesXml(generateDOCXSync(await parseDOCX(dotx), { variant: "docx" }));
    expect(xml).toContain(
      `<Override PartName="/word/document.xml" ContentType="${MAIN_DOCUMENT_CONTENT_TYPES.docx}"/>`,
    );
    expect(xml).not.toContain("template.main+xml");
  });
});

describe("generateDOCX variants keep fresh-compile parts", () => {
  /** A fresh two-item numbered list — the model's generated numbering
   *  definitions must reach `word/numbering.xml` in every variant. */
  const numberedDoc: JSONContent = {
    type: "doc",
    content: [
      {
        type: "paragraph",
        attrs: { numbering: { reference: "docen-ordered-1", level: 0 } },
        content: [{ type: "text", text: "one" }],
      },
      {
        type: "paragraph",
        attrs: { numbering: { reference: "docen-ordered-1", level: 0 } },
        content: [{ type: "text", text: "two" }],
      },
    ],
  };

  /** A fresh note document with one referenced note, the shape the editor's
   *  note dialog commits (inline passthrough reference + documentExtras). */
  const noteDoc = (kind: "footnotes" | "endnotes"): JSONContent => ({
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          { type: "text", text: "body" },
          {
            type: "inlinePassthrough",
            attrs: {
              data: JSON.stringify({
                [`${kind === "footnotes" ? "footnote" : "endnote"}Reference`]: 1,
              }),
            },
          },
        ],
      },
    ],
    attrs: {
      documentExtras: {
        [kind]: [
          {
            id: 1,
            children: [
              {
                paragraph: {
                  children: [
                    {
                      text: "a note",
                      style: kind === "footnotes" ? "FootnoteText" : "EndnoteText",
                    },
                  ],
                },
              },
            ],
          },
        ],
      },
    },
  });

  for (const variant of ["docx", "docm", "dotx", "dotm"] as const) {
    it(`${variant}: a fresh numbered list keeps word/numbering.xml with no dangling numId`, () => {
      const bytes = generateDOCXSync(numberedDoc, { variant });
      const zip = unzipSync(bytes);
      const numberingXml = new TextDecoder().decode(zip["word/numbering.xml"]);
      expect(numberingXml).toBeTruthy();
      expect(contentTypesXml(bytes)).toContain('PartName="/word/numbering.xml"');
      const documentXml = new TextDecoder().decode(zip["word/document.xml"]);
      const numIds = [...documentXml.matchAll(/<w:numId w:val="(\d+)"/g)].map((match) => match[1]!);
      expect(numIds.length).toBeGreaterThan(0);
      for (const id of numIds) {
        expect(numberingXml).toContain(`<w:num w:numId="${id}"`);
      }
    });

    for (const kind of ["footnotes", "endnotes"] as const) {
      it(`${variant}: a fresh ${kind.slice(0, -1)} document keeps word/${kind}.xml + declaration + relationship`, () => {
        const bytes = generateDOCXSync(noteDoc(kind), { variant });
        const zip = unzipSync(bytes);
        const part = new TextDecoder().decode(zip[`word/${kind}.xml`]);
        expect(part).toContain("a note");
        expect(contentTypesXml(bytes)).toContain(`PartName="/word/${kind}.xml"`);
        const rels = new TextDecoder().decode(zip["word/_rels/document.xml.rels"]);
        expect(rels).toContain(`Target="${kind}.xml"`);
        // The note reference in document.xml is backed by the emitted part.
        const documentXml = new TextDecoder().decode(zip["word/document.xml"]);
        expect(documentXml).toContain(
          kind === "footnotes" ? "footnoteReference" : "endnoteReference",
        );
      });
    }
  }

  it("matches the no-variant part set for a fresh document", () => {
    const base = Object.keys(unzipSync(generateDOCXSync(numberedDoc))).sort();
    for (const variant of ["docx", "docm", "dotx", "dotm"] as const) {
      const parts = Object.keys(unzipSync(generateDOCXSync(numberedDoc, { variant }))).sort();
      expect(parts).toEqual(base);
    }
  });
});

describe("docm macro-part round-trip", () => {
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

describe("drop cap framePr round-trip", () => {
  it("writes w:hSpace/w:vSpace and parses the distances back", () => {
    const json = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: {
            dropCap: { val: "drop", lines: 3, distance: 200, vDistance: 40 },
          },
          content: [{ type: "text", text: "Once upon a time" }],
        },
      ],
    };
    const bytes = generateDOCXSync(json as never);
    const xml = new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]);
    expect(xml).toContain('w:dropCap="drop"');
    expect(xml).toContain('w:lines="3"');
    expect(xml).toContain('w:hSpace="200"');
    expect(xml).toContain('w:vSpace="40"');

    const reparsed = parseDOCXSync(bytes) as never as {
      content: Array<{ attrs?: Record<string, unknown> }>;
    };
    expect(reparsed.content[0]?.attrs?.dropCap).toEqual({
      val: "drop",
      lines: 3,
      distance: 200,
      vDistance: 40,
    });
  });
});

describe("notes schema order (w:pPr before w:r)", () => {
  /**
   * CT_P requires w:pPr first; office-open's endnote writer used to inject its
   * reference run before it (`<w:p><w:r>…<w:endnoteRef/>…</w:r><w:pPr>`), a
   * Word-repair risk on every endnote-bearing export. Both a model-carried
   * reference and the engine-injected one must land after the properties.
   */
  it("emits the endnote reference runs after the paragraph properties", () => {
    const doc: DocumentOptions = {
      sections: [{ children: [{ paragraph: { text: "body" } }] }],
      endnotes: [
        {
          id: 1,
          children: [
            {
              paragraph: {
                style: "EndnoteText",
                children: [{ children: [{ endnoteRef: true }] }, { text: " with mark" }],
              },
            },
          ],
        },
        {
          id: 2,
          children: [{ paragraph: { style: "EndnoteText", children: [{ text: " no mark" }] } }],
        },
      ],
    };
    const zip = unzipSync(generateDOCXSync(resolveDocument(doc), { prepare: false }));
    const xml = new TextDecoder().decode(zip["word/endnotes.xml"]);
    const notes = [...xml.matchAll(/<w:endnote w:id="[12]">(.*?)<\/w:endnote>/gs)];
    expect(notes).toHaveLength(2);
    for (const note of notes) {
      const body = note[1]!;
      const pPr = body.indexOf("<w:pPr>");
      const run = body.indexOf("<w:r");
      expect(pPr, note[0]).toBeGreaterThanOrEqual(0);
      expect(pPr, note[0]).toBeLessThan(run);
      expect(body).toContain("<w:endnoteRef/>");
    }
  });
});

describe("multi-section DocumentOptions resolve", () => {
  it("keeps a propertyless section's break instead of merging it", () => {
    const json = resolveDocument({
      sections: [
        { children: [{ paragraph: { children: ["one"] } }] },
        { children: [{ paragraph: { children: ["two"] } }] },
      ],
    });
    const first = json.content?.[0];
    expect(first?.type).toBe("paragraph");
    expect(first?.attrs?.sectionProperties).toEqual({});
    const xml = new TextDecoder().decode(
      unzipSync(generateDOCXSync(json, { prepare: false }) as Uint8Array)["word/document.xml"],
    );
    // One sectPr closing section one, one final sectPr.
    expect([...xml.matchAll(/<w:sectPr/g)]).toHaveLength(2);
  });

  it("carries a section's pageNumberType through to w:pgNumType", () => {
    const json = resolveDocument({
      sections: [
        { children: [{ paragraph: { children: ["one"] } }] },
        {
          properties: { pageNumberType: { start: 5 } },
          children: [{ paragraph: { children: ["two"] } }],
        },
      ],
    });
    const xml = new TextDecoder().decode(
      unzipSync(generateDOCXSync(json, { prepare: false }) as Uint8Array)["word/document.xml"],
    );
    expect(xml).toContain('<w:pgNumType w:start="5"/>');
  });
});

describe("embedded-font relationship escaping", () => {
  /**
   * A relationship Target is a URI: a font family name with spaces, `&`, `"`,
   * angle brackets or CJK must not leak raw into fontTable.xml.rels. The
   * engine escapes each path segment; the parser decodes it again so the font
   * bytes and family name survive the round-trip.
   */
  it("escapes a hostile family name and round-trips the font bytes", () => {
    const family = 'Hostile & <Font> "Q" (v2)';
    const data = new Uint8Array(Array.from({ length: 128 }, (_, i) => (i * 11 + 5) % 256));
    const json = resolveDocument({ sections: [{ children: [{ paragraph: { text: "x" } }] }] });
    const gen1 = generateDOCXSync(json, {
      prepare: false,
      document: { fonts: [{ name: family, data }] },
    }) as Uint8Array;
    const decoder = new TextDecoder();
    const zip1 = unzipSync(gen1);
    const rels1 = decoder.decode(zip1["word/_rels/fontTable.xml.rels"]);
    const target1 = /Target="([^"]+)"/.exec(rels1)?.[1] ?? "";
    expect(target1).toContain("%20");
    expect(target1).toContain("%26");
    expect(target1).not.toMatch(/[ "&<>]/);
    // The ZIP item name equals the escaped Target: readers resolve the Target
    // as a URI and look up the part by that exact name (python-docx and
    // LibreOffice both reject an unescaped part name behind an escaped
    // Target).
    const part1 = `word/${target1}`;
    expect(zip1[part1]).toBeDefined();

    const reparsed = parseDOCXSync(gen1) as JSONContent;
    const fonts = (
      reparsed.attrs?.documentExtras as { fonts?: { name?: string; data?: Uint8Array }[] }
    )?.fonts;
    expect(fonts?.[0]?.name).toBe(family);
    expect(fonts?.[0]?.data?.byteLength).toBeGreaterThan(0);

    const gen2 = generateDOCXSync(reparsed, { prepare: false }) as Uint8Array;
    const zip2 = unzipSync(gen2);
    const rels2 = decoder.decode(zip2["word/_rels/fontTable.xml.rels"]);
    expect(/Target="([^"]+)"/.exec(rels2)?.[1]).toBe(target1);
    expect(zip2[part1]).toEqual(zip1[part1]);
  });
});

describe("header/footer relationship stability", () => {
  /**
   * An opened package keeps its header/footer parts and the document
   * relationships that point at them. Re-generating must reuse the source
   * relationship ids (kind+target) instead of allocating fresh ones above the
   * reserved source range: allocating grew rId21…rId26 → rId27…rId32 on every
   * save and made byte-stable round-trips impossible.
   */
  it("reuses the source rIds and part names across save cycles", () => {
    const doc: DocumentOptions = {
      settings: { evenAndOddHeaders: true },
      sections: [
        {
          properties: {
            ...PAGE,
            titlePage: true,
            pageNumberType: { format: "lowerRoman", start: 1 },
          },
          headers: {
            default: [{ paragraph: { text: "A-DEF" } }],
            first: [{ paragraph: { text: "A-FIRST" } }],
            even: [{ paragraph: { text: "A-EVEN" } }],
          },
          footers: { default: [{ paragraph: { text: "A-FOOT" } }] },
          children: [{ paragraph: { text: "section one" } }],
        },
        {
          properties: PAGE,
          headers: { default: [{ paragraph: { text: "B-DEF" } }] },
          children: [{ paragraph: { text: "section two" } }],
        },
      ],
    };

    const gen1 = generateDOCXSync(resolveDocument(doc), { prepare: false });
    const parsed1 = parseDOCXSync(gen1) as JSONContent;
    const slots = parsed1.content?.[0]?.attrs?.sectionHeaders as
      | { partNames?: Record<string, string> }
      | undefined;
    expect(slots?.partNames?.default).toBe("header1.xml");

    const gen2 = generateDOCXSync(parsed1, { prepare: false });
    const gen3 = generateDOCXSync(parseDOCXSync(gen2), { prepare: false });
    expect(Buffer.from(gen2).equals(Buffer.from(gen3))).toBe(true);

    const refs = (bytes: Uint8Array): string[] => {
      const rels = new TextDecoder().decode(unzipSync(bytes)["word/_rels/document.xml.rels"]);
      return [...rels.matchAll(/Id="(rId\d+)" Type="[^"]*\/(?:header|footer)" Target="([^"]+)"/g)]
        .map((match) => `${match[1]}:${match[2]}`)
        .sort();
    };
    expect(refs(gen3)).toEqual(refs(gen2));
    expect(refs(gen2).some((entry) => entry.endsWith(":header1.xml"))).toBe(true);
  });
});
