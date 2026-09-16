import {
  DocPartGallery,
  generateDocumentSync,
  parseDocumentSync,
  type DocumentOptions,
  type GlossaryDocumentOptions,
} from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, generateDOCXSync, parseDOCXSync, resolveDocument } from "../index";

/**
 * Evidence for the D2 glossary-part decision: office-open parses
 * word/glossary/document.xml into DocumentOptions.glossary and writes it back
 * on generate, and docen's documentExtras passthrough (glossary is not a
 * compile-owned key) carries it through resolve→compile untouched. These specs
 * pin that contract so the building-blocks feature can persist real Word
 * Quick Parts instead of a private store.
 */

/** One doc-part carrying a paragraph with a bold run. */
function glossaryDoc(): DocumentOptions {
  return {
    sections: [{ children: [{ paragraph: { text: "body" } }] }],
    glossary: {
      parts: [
        {
          name: "Signature",
          gallery: DocPartGallery.CUSTOM_QUICK_PARTS,
          category: "General",
          description: "Sign-off block",
          guid: "{11111111-2222-3333-4444-555555555555}",
          sections: [
            {
              children: [
                {
                  paragraph: {
                    children: [{ text: "Best " }, { text: "regards", bold: true }, { text: "," }],
                  },
                },
              ],
            },
          ],
        },
      ],
    },
  };
}

describe("glossary part round-trip (office-open)", () => {
  it("parses word/glossary/document.xml from bytes", () => {
    const bytes = generateDocumentSync(glossaryDoc());
    const parsed = parseDocumentSync(bytes);
    expect(parsed.glossary?.parts).toHaveLength(1);
    expect(parsed.glossary?.parts[0]?.name).toBe("Signature");
    expect(parsed.glossary?.parts[0]?.gallery).toBe(DocPartGallery.CUSTOM_QUICK_PARTS);
    expect(parsed.glossary?.parts[0]?.description).toBe("Sign-off block");
  });

  it("survives docen resolve → documentExtras → compile → generate → parse", () => {
    const first = parseDocumentSync(generateDocumentSync(glossaryDoc()));
    const json = resolveDocument(first);
    const extras = (json.attrs?.documentExtras ?? {}) as Record<string, unknown>;
    expect(extras.glossary).toBeDefined();

    const compiled = compileDocument(json);
    expect(compiled.glossary).toEqual(first.glossary);

    // office-open directly on the docen-compiled options (isolates the loss).
    const direct = parseDocumentSync(generateDocumentSync(compiled));
    expect(direct.glossary?.parts[0]?.name).toBe("Signature");

    const reparsed = parseDOCXSync(generateDOCXSync(json));
    const reparsedExtras = (reparsed.attrs?.documentExtras ?? {}) as {
      glossary?: GlossaryDocumentOptions;
    };
    expect(reparsedExtras.glossary?.parts[0]?.name).toBe("Signature");
    const child = reparsedExtras.glossary?.parts[0]?.sections?.[0]?.children?.[0];
    expect(child).toMatchObject({ paragraph: expect.anything() });
  });

  it("omits glossary when the JSON carries none", () => {
    const compiled = compileDocument({ type: "doc", content: [{ type: "paragraph" }] });
    expect("glossary" in compiled).toBe(false);
  });
});
