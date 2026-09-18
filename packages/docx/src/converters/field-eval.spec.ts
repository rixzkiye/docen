import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vitest";

import type { JSONContent } from "../core";
import { generateDOCXSync, parseDOCXSync } from "../index";
import { fillGeneratedFields } from "./field-eval";

/**
 * Items 13 + 17: the builder must emit correct cached field results so
 * Word/LibreOffice show real values before any update — SEQ ordinals (with
 * `\s` heading restarts and `\*` formats), REF bookmark text, PAGE/NUMPAGES/
 * PAGEREF via the page context, and a TOC whose cached entries were missing.
 *
 * Oracles: the generated `word/document.xml` bytes (office-open's writer),
 * a python-docx structural read of the same package, LibreOffice PDF text,
 * and the DOCX → JSON → DOCX round-trip for the TOC field structure.
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

const complex = (instruction: string, result?: string, extra?: object): JSONContent =>
  field({ complexField: { instruction, ...(result !== undefined ? { result } : {}), ...extra } });

const bookmarkStart = (id: number, name: string): JSONContent =>
  field({ bookmarkStart: { id, name } });
const bookmarkEnd = (id: number): JSONContent => field({ bookmarkEnd: { id } });

const docOf = (content: JSONContent[], attrs?: Record<string, unknown>): JSONContent => ({
  type: "doc",
  ...(attrs ? { attrs } : {}),
  content,
});

const documentXml = (json: JSONContent, options?: Parameters<typeof generateDOCXSync>[1]): string =>
  new TextDecoder().decode(
    unzipSync(generateDOCXSync(json, { prepare: false, ...options }) as Uint8Array)[
      "word/document.xml"
    ],
  );

/** The TOC sdt slice of `document.xml` — entry assertions must not match the
 *  body headings the entries quote. */
const tocXml = (json: JSONContent, options?: Parameters<typeof generateDOCXSync>[1]): string => {
  const xml = documentXml(json, options);
  const start = xml.indexOf("<w:sdt>");
  const end = xml.indexOf("</w:sdt>");
  return start >= 0 && end > start ? xml.slice(start, end + 8) : "";
};

describe("SEQ generation-time numbering", () => {
  it("numbers per label, restarts at the `\\s` heading level, formats `\\*`", () => {
    const json = docOf([
      para([text("A")], { heading: "Heading1" }),
      para([simple(" SEQ Figure \\* ARABIC \\s 1 ", "99")]),
      para([simple(" SEQ Figure \\* ROMAN \\s 1 ", "99")]),
      para([text("B")], { heading: "Heading1" }),
      para([simple(" SEQ Figure \\* ARABIC \\s 1 ", "99")]),
      para([simple(" SEQ Table \\* ARABIC ", "99")]),
    ]);
    const xml = documentXml(json);
    // First heading restarts both labels: Figure 1 (Arabic) + Figure I (Roman).
    // Second heading restarts Figure again; Table keeps counting (no \s).
    const texts = [
      ...xml.matchAll(/<w:fldSimple[^>]*w:instr="([^"]*)"[^>]*>(.*?)<\/w:fldSimple>/g),
    ];
    const results = texts.map((match) => ({
      instr: match[1]!,
      value: [...match[2]!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]!).join(""),
    }));
    expect(results.map((r) => r.value)).toEqual(["1-1", "1-II", "2-1", "1"]);
  });

  it("patches a complex field's flat and structured caches", () => {
    const json = docOf([
      para([
        complex(" SEQ Figure \\* ARABIC ", "9", {
          resultRunsXml: "<w:r><w:t>9</w:t></w:r>",
        }),
      ]),
    ]);
    const xml = documentXml(json);
    expect(xml).not.toContain("<w:fldSimple");
    expect(xml).toContain('<w:t xml:space="preserve">1</w:t>');
    expect(xml).not.toContain("<w:t>9</w:t>");
  });

  it("prefixes the chapter number from the settings separator", () => {
    const json = docOf(
      [
        para([text("Chapter")], { heading: "Heading1" }),
        para([simple(" SEQ Figure \\* ARABIC \\s 1 ", "0")]),
      ],
      {
        documentExtras: {
          settings: { captions: { captions: [{ name: "Figure", sep: "period" }] } },
        },
      },
    );
    const xml = documentXml(json);
    expect(xml).toContain("1.1");
  });
});

describe("REF generation-time cross-references", () => {
  it("resolves backward and forward references to the bookmark text", () => {
    const json = docOf([
      para([simple(" REF _Ref1 \\h ", "stale")]),
      para([bookmarkStart(1, "_Ref1"), text("Alpha"), bookmarkEnd(1)]),
      para([simple(" REF _Ref1 \\h ", "stale")]),
    ]);
    const xml = documentXml(json);
    const values = [
      ...xml.matchAll(/<w:fldSimple[^>]*w:instr="[^"]*REF[^"]*"[^>]*>(.*?)<\/w:fldSimple>/g),
    ].map((m) => [...m[1]!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]!).join(""));
    expect(values).toEqual(["Alpha", "Alpha"]);
    expect(xml).not.toContain("stale");
  });

  it("keeps the cache for switches the builder cannot resolve (\\p, \\n)", () => {
    const json = docOf([
      para([bookmarkStart(1, "_Ref1"), text("Alpha"), bookmarkEnd(1)]),
      para([simple(" REF _Ref1 \\p ", "cached")]),
      para([simple(" REF _Ref1 \\n ", "cached")]),
    ]);
    const xml = documentXml(json);
    expect(xml).toContain("cached");
  });
});

describe("page fields through the caller's page context", () => {
  const pageDoc = docOf([
    para([simple(" PAGE ", "9")]),
    para([complex(" NUMPAGES ", "9")]),
    para([simple(" PAGEREF _Ref1 \\h ", "9")]),
    para([bookmarkStart(1, "_Ref1"), text("target"), bookmarkEnd(1)]),
    para([simple(" PAGE ", "9")]),
  ]);

  it("fills PAGE/NUMPAGES/PAGEREF from pageOf and clears stale caches", () => {
    const pages = new Map<number, number>([
      [0, 1],
      [1, 3],
      [2, 2],
      [3, 3],
    ]);
    const xml = documentXml(pageDoc, {
      fields: {
        pageCount: 3,
        pageOf: ({ index, bookmark }) => (bookmark ? pages.get(2) : pages.get(index)),
      },
    });
    expect(xml).toContain(">1<");
    expect(xml).toContain(">3<");
    expect(xml).toContain(">2<");
    expect(xml).not.toContain(">9<");
  });

  it("never invents a page number without a context", () => {
    const xml = documentXml(pageDoc);
    expect(xml).toContain(">9<");
  });
});

describe("TOC cached entries (item 13)", () => {
  const headingToc = docOf([
    para([text("Alpha")], { heading: "Heading1" }),
    para([text("Beta")], { heading: "Heading2" }),
    para([text("Body")]),
    {
      type: "tocField",
      attrs: { options: { headingStyleRange: "1-3", hyperlink: false } },
      content: [{ type: "paragraph" }],
    },
  ]);

  it("fills a placeholder TOC from the document headings and keeps the field updatable", () => {
    const bytes = generateDOCXSync(headingToc, { prepare: false }) as Uint8Array;
    const xml = new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]);
    // A real TOC field instruction + the rendered entries.
    expect(xml).toMatch(/TOC/);
    expect(xml).toContain(">Alpha<");
    expect(xml).toContain(">Beta<");
    expect(xml).toContain("TOC1");
    expect(xml).toContain("TOC2");
    // Round-trip: the generated TOC parses back as a tocField with entries
    // AND its field switches, so the editor's Update Table can rebuild it.
    const parsed = parseDOCXSync(bytes);
    const toc = parsed.content?.find((n) => n.type === "tocField");
    expect(toc?.content?.length).toBe(2);
    expect(
      (toc?.attrs?.options as { headingStyleRange?: string } | undefined)?.headingStyleRange,
    ).toBe("1-3");
    expect(
      parsed.content?.some((n) => n.type === "paragraph" && n.content?.[0]?.text === "Body"),
    ).toBe(true);
  });

  it("never recomputes a TOC that already carries entries", () => {
    const withEntries: JSONContent = {
      ...headingToc,
      content: [
        para([text("Alpha")], { heading: "Heading1" }),
        {
          type: "tocField",
          attrs: { options: { headingStyleRange: "1-3" } },
          content: [para([text("Existing entry")])],
        },
      ],
    };
    const xml = documentXml(withEntries);
    expect(xml).toContain("Existing entry");
  });

  it("fills a `\\c` table of figures from the matching captions", () => {
    const figure = docOf([
      para([text("Figure "), simple(" SEQ Figure \\* ARABIC ", "1"), text(": One")], {
        style: "Caption",
      }),
      para([text("Figure "), simple(" SEQ Figure \\* ARABIC ", "2"), text(": Two")], {
        style: "Caption",
      }),
      {
        type: "tocField",
        attrs: { options: { captionLabel: "Figure" } },
        content: [{ type: "paragraph" }],
      },
    ]);
    const xml = documentXml(figure);
    // The caption's SEQ result is part of the cached figure entry.
    expect(xml).toContain("Figure 1: One");
    expect(xml).toContain("Figure 2: Two");
    expect(xml).toMatch(/SEQ Figure/);
    const parsed = parseDOCXSync(generateDOCXSync(figure, { prepare: false }) as Uint8Array);
    const toc = parsed.content?.find((n) => n.type === "tocField");
    expect(toc?.content?.length).toBe(2);
  });

  it("honors a parsed DOCX `\\t` stylesWithLevels switch", () => {
    const json = docOf([
      para([text("Heading One")], { heading: "Heading1" }),
      para([text("Custom Style Heading")], { style: "MyHeading" }),
      {
        type: "tocField",
        attrs: {
          options: {
            headingStyleRange: "1-3",
            stylesWithLevels: [{ styleName: "MyHeading", level: 2 }],
          },
        },
        content: [{ type: "paragraph" }],
      },
    ]);
    const toc = tocXml(json);
    expect(toc).toContain(">Custom Style Heading<");
    expect(toc).toContain('\\t "MyHeading,2"');
  });

  it("matches `\\t` styles by their resolved style name, not just the id", () => {
    const json = docOf(
      [
        para([text("Named Heading")], { style: "MyStyle" }),
        {
          type: "tocField",
          attrs: {
            options: {
              headingStyleRange: "1-3",
              stylesWithLevels: [{ styleName: "My Heading", level: 2 }],
            },
          },
          content: [{ type: "paragraph" }],
        },
      ],
      { styles: { paragraphStyles: [{ id: "MyStyle", name: "My Heading", basedOn: "Normal" }] } },
    );
    const toc = tocXml(json);
    expect(toc).toContain(">Named Heading<");
    expect(toc).toContain("TOC2");
  });

  it("scopes entries to a `\\b` entriesFromBookmark", () => {
    const json = docOf([
      para([text("Outside Before")], { heading: "Heading1" }),
      para([bookmarkStart(1, "Scope"), text("Inside One")], { heading: "Heading1" }),
      para([field({ bookmarkEnd: { id: 1 } })]),
      para([text("Outside After")], { heading: "Heading1" }),
      {
        type: "tocField",
        attrs: { options: { headingStyleRange: "1-3", entriesFromBookmark: "Scope" } },
        content: [{ type: "paragraph" }],
      },
    ]);
    const toc = tocXml(json);
    expect(toc).toContain("Inside One");
    expect(toc).not.toContain("Outside Before");
    expect(toc).not.toContain("Outside After");
  });

  it("leaves the scope empty (placeholder field) for an unknown bookmark", () => {
    const json = docOf([
      para([text("One")], { heading: "Heading1" }),
      {
        type: "tocField",
        attrs: { options: { headingStyleRange: "1-3", entriesFromBookmark: "Missing" } },
        content: [{ type: "paragraph" }],
      },
    ]);
    const xml = documentXml(json);
    expect(xml).toContain("TOC");
    expect(xml).toContain('w:dirty="1"');
    expect(tocXml(json)).not.toContain(">One<");
  });

  it("fills a `\\c` table of figures from the parsed captionLabelIncludingNumbers", () => {
    const json = docOf([
      para([text("Figure "), simple(" SEQ Figure \\* ARABIC ", "1"), text(": One")], {
        style: "Caption",
      }),
      {
        type: "tocField",
        attrs: { options: { captionLabelIncludingNumbers: "Figure" } },
        content: [{ type: "paragraph" }],
      },
    ]);
    const toc = tocXml(json);
    expect(toc).toContain("Figure 1: One");
    expect(toc).toContain('\\c "Figure"');
  });

  it("does not include outline-only paragraphs when `\\u` is off", () => {
    const json = docOf([
      para([text("Styled Heading")], { heading: "Heading1" }),
      para([text("Outline Only")], { outlineLevel: 0 }),
      {
        type: "tocField",
        attrs: { options: { headingStyleRange: "1-3", useAppliedParagraphOutlineLevel: false } },
        content: [{ type: "paragraph" }],
      },
    ]);
    const toc = tocXml(json);
    expect(toc).toContain("Styled Heading");
    expect(toc).not.toContain(">Outline Only<");
  });

  it("keeps an unfillable TOC's field instruction (dirty placeholder)", () => {
    const json = docOf([
      para([text("body only")]),
      {
        type: "tocField",
        attrs: { options: { headingStyleRange: "1-3" } },
        content: [{ type: "paragraph" }],
      },
    ]);
    const bytes = generateDOCXSync(json, { prepare: false }) as Uint8Array;
    const xml = new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]);
    expect(xml).toContain("w:fldChar");
    expect(xml).toContain('w:dirty="1"');
    expect(xml).toContain("TOC");
    // Round-trip: the placeholder field parses back as a tocField.
    const parsed = parseDOCXSync(bytes);
    expect(parsed.content?.some((n) => n.type === "tocField")).toBe(true);
  });

  it("fills the generated entries' page numbers from tocPageOf", () => {
    const json = docOf([
      para([text("One")], { heading: "Heading1" }),
      para([text("Two")], { heading: "Heading2" }),
      {
        type: "tocField",
        attrs: { options: { headingStyleRange: "1-3" } },
        content: [{ type: "paragraph" }],
      },
    ]);
    const toc = tocXml(json, {
      fields: { tocPageOf: ({ index }) => [3, 5][index] },
    });
    expect(toc).toContain(">One<");
    expect(toc).toContain(">3<");
    expect(toc).toContain(">Two<");
    expect(toc).toContain(">5<");
  });

  it("re-resolves REF caches after the referenced text is edited", () => {
    const refCache = (json: JSONContent): string => {
      const match = /<w:fldSimple[^>]*w:instr="[^"]*REF[^"]*"[^>]*>(.*?)<\/w:fldSimple>/.exec(
        documentXml(json),
      );
      return [...match![1]!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
        .map((entry) => entry[1]!)
        .join("");
    };
    const before = docOf([
      para([bookmarkStart(1, "_Ref1"), text("Alpha"), bookmarkEnd(1)]),
      para([simple(" REF _Ref1 \\h ", "")]),
    ]);
    expect(refCache(before)).toBe("Alpha");
    // Edit the referenced text, then regenerate: the stale cache follows.
    const edited = docOf([
      para([bookmarkStart(1, "_Ref1"), text("Beta"), bookmarkEnd(1)]),
      para([simple(" REF _Ref1 \\h ", "Alpha")]),
    ]);
    expect(refCache(edited)).toBe("Beta");
  });

  it("leaves a document without any patchable field untouched (non-mutating)", () => {
    const plain = docOf([para([text("hello")])]);
    const snapshot = JSON.stringify(plain);
    const filled = fillGeneratedFields(plain);
    expect(filled).toBe(plain);
    expect(JSON.stringify(plain)).toBe(snapshot);
  });

  it("does not mutate the input when it does patch", () => {
    const json = docOf([para([simple(" SEQ Figure ", "9")])]);
    const snapshot = JSON.stringify(json);
    const filled = fillGeneratedFields(json);
    expect(filled).not.toBe(json);
    expect(JSON.stringify(json)).toBe(snapshot);
  });
});
