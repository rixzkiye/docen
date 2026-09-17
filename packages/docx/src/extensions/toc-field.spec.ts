import type { DocumentOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, resolveDocument, type JSONContent } from "../index";

/**
 * TOC field-result boundary. office-open's writer parks the field-end run in
 * the last rendered entry's paragraph; the parser used to treat that closing
 * paragraph as body content and drop it, so every round-trip exported a TOC
 * one entry shorter. These pin the field boundary: N rendered entries (with
 * repeats) survive parse→export cycles untouched, and nothing leaks into the
 * body.
 */

function tocEntry(style: string, anchor: string, text: string) {
  return {
    paragraph: {
      style,
      children: [
        { hyperlink: { anchor, children: [{ text }] } },
        { text: "\t" },
        { simpleField: { instruction: `PAGEREF ${anchor} \\h`, result: "1" } },
      ],
    },
  };
}

function tocDoc(entries: ReturnType<typeof tocEntry>[]): DocumentOptions {
  return {
    sections: [
      {
        children: [
          {
            toc: {
              hyperlink: true,
              headingStyleRange: "1-3",
              useAppliedParagraphOutlineLevel: true,
              entries,
            },
          },
          { paragraph: { text: "after toc" } },
        ],
      },
    ],
  };
}

function tocOf(json: JSONContent): JSONContent | undefined {
  return json.content?.find((n) => n.type === "tocField");
}

function entryTexts(toc: JSONContent | undefined): string[] {
  return (toc?.content ?? []).map(
    (entry) =>
      entry.content
        ?.filter((c) => c.type === "text")
        .map((c) => c.text ?? "")
        .join("")
        .replace(/\t/g, "") ?? "",
  );
}

describe("tocField entry round-trip", () => {
  it("keeps N entries (with repeats) across two export cycles", () => {
    const doc = tocDoc([
      tocEntry("TOC1", "_Toc1", "Alpha"),
      tocEntry("TOC2", "_Toc2", "Alpha sub"),
      tocEntry("TOC1", "_Toc1", "Alpha"), // repeated entry (same anchor + text)
    ]);
    const gen1 = generateDOCXSync(resolveDocument(doc), { prepare: false }) as Uint8Array;
    const parsed1 = parseDOCXSync(gen1);
    const toc1 = tocOf(parsed1);
    expect(toc1?.content).toHaveLength(3);
    expect(entryTexts(toc1)).toEqual(["Alpha", "Alpha sub", "Alpha"]);
    // The paragraph after the field stayed body content, no entry leaked out.
    expect(
      parsed1.content?.some((n) => n.type === "paragraph" && n.content?.[0]?.text === "after toc"),
    ).toBe(true);

    const gen2 = generateDOCXSync(parsed1, { prepare: false }) as Uint8Array;
    const toc2 = tocOf(parseDOCXSync(gen2));
    expect(toc2?.content).toHaveLength(3);
    expect(entryTexts(toc2)).toEqual(["Alpha", "Alpha sub", "Alpha"]);
  });

  it("keeps a single-entry TOC (the field end in the only entry)", () => {
    const doc = tocDoc([tocEntry("TOC1", "_Toc1", "Only")]);
    const parsed = parseDOCXSync(
      generateDOCXSync(resolveDocument(doc), { prepare: false }) as Uint8Array,
    );
    const toc = tocOf(parsed);
    expect(toc?.content).toHaveLength(1);
    expect(entryTexts(toc)).toEqual(["Only"]);
  });
});
