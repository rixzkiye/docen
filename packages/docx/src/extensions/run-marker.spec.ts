import { unzipSync } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import { generateDocumentSync } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import type { JSONContent } from "../index";
import { generateDOCXSync, parseDOCXSync } from "../index";
import { PRESERVED_RUN_ELEMENTS } from "./coverage";

/**
 * The item-4 fixture matrix: every run-level element the resolve walk used to
 * drop gets one fixture per disposition dimension —
 *
 *  1. body run child → resolve produces a `runMarker`, compile re-emits it,
 *     and the generated `word/document.xml` carries the exact `<w:…/>`
 *     element (Word/LO meaning preserved, not literal text);
 *  2. second generation is byte-stable (parse → generate idempotence);
 *  3. the PAGE-field result shape (`w:fldSimple` whose cached run holds
 *     `w:pgNum`) keeps the element through the field's round-trip.
 *
 * The XML assertions use office-open's own writer output (unzipped part
 * bytes) — an independent oracle from the resolve/compile JSON under test.
 */

const docXml = (bytes: Uint8Array): string =>
  new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]);

const markerElements = (json: JSONContent): string[] => {
  const found: string[] = [];
  const walk = (node: JSONContent): void => {
    if (node.type === "runMarker" && typeof node.attrs?.element === "string") {
      found.push(node.attrs.element);
    }
    for (const child of node.content ?? []) walk(child);
  };
  walk(json);
  return found;
};

const TAGS = Object.entries(PRESERVED_RUN_ELEMENTS);

describe("run element fixture matrix (body runs)", { timeout: 20000 }, () => {
  for (const [tag, entry] of TAGS) {
    it(`preserves ${entry.element} (${entry.semantics})`, () => {
      // gen1 authored as DocumentOptions: a run holding text + the element.
      const model: DocumentOptions = {
        sections: [
          {
            children: [
              {
                paragraph: {
                  children: [{ text: "x", children: [{ [tag]: true }] as never }],
                },
              },
            ],
          },
        ],
      };
      const bytes1 = generateDocumentSync(model) as Uint8Array;
      expect(docXml(bytes1)).toContain(`<w:${tag}/>`);

      // DOCX → JSON: the element is claimed by the runMarker atom.
      const parsed = parseDOCXSync(bytes1);
      expect(markerElements(parsed)).toContain(tag);

      // JSON → DOCX: the element re-emits; second generation is stable.
      const bytes2 = generateDOCXSync(parsed, { prepare: false }) as Uint8Array;
      expect(docXml(bytes2)).toContain(`<w:${tag}/>`);
      const bytes3 = generateDOCXSync(parseDOCXSync(bytes2), { prepare: false }) as Uint8Array;
      expect(docXml(bytes3)).toBe(docXml(bytes2));
    });
  }
});

describe("run elements inside field results", { timeout: 20000 }, () => {
  it("keeps w:pgNum inside a PAGE simple field's cached result", () => {
    const model: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  { text: "Page " },
                  {
                    simpleField: {
                      instruction: " PAGE ",
                      // The field's structured cached result: Word writes a
                      // bare pgNum run here, not a literal number.
                      cachedRunsXml: "<w:r><w:pgNum/></w:r>",
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const bytes1 = generateDocumentSync(model) as Uint8Array;
    const xml1 = docXml(bytes1);
    expect(xml1).toContain('w:instr=" PAGE "');
    expect(xml1).toContain("<w:pgNum/>");

    // The field rides the inlinePassthrough atom; the cached runs XML is
    // carried verbatim (pgNum included), never split into a dropped child.
    const parsed = parseDOCXSync(bytes1);
    const bytes2 = generateDOCXSync(parsed, { prepare: false }) as Uint8Array;
    expect(docXml(bytes2)).toContain("<w:pgNum/>");
  });

  it("keeps a date field's day/month/year elements in the result runs", () => {
    const model: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    complexField: {
                      instruction: ' DATE \\@ "yyyy/M/d" ',
                      resultRunsXml:
                        "<w:r><w:t>2026/9/18</w:t></w:r>" +
                        "<w:r><w:dayShort/><w:t>, </w:t><w:monthLong/><w:t> </w:t><w:dayLong/><w:t>, </w:t><w:yearLong/></w:r>",
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const bytes1 = generateDocumentSync(model) as Uint8Array;
    const xml1 = docXml(bytes1);
    for (const tag of ["dayShort", "monthLong", "dayLong", "yearLong"]) {
      expect(xml1).toContain(`<w:${tag}/>`);
    }
    const bytes2 = generateDOCXSync(parseDOCXSync(bytes1), { prepare: false }) as Uint8Array;
    for (const tag of ["dayShort", "monthLong", "dayLong", "yearLong"]) {
      expect(docXml(bytes2)).toContain(`<w:${tag}/>`);
    }
  });
});
