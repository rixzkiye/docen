import type { DocumentOptions, ParagraphChild, RunOptions, SectionChild } from "@office-open/docx";
import { generateDocumentSync, parseDocumentSync } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import {
  compileDocument,
  docxExtensions,
  formatMarkNames,
  resolveDocument,
  runPropsFromMarks,
  runPropsToMarks,
} from "../index";

/**
 * Track-change format revisions (w:rPrChange / w:pPrChange): the runtime model
 * carries the OLD properties (formatChange mark / paragraph `revision` attr)
 * and both directions round-trip through DocumentOptions and through real
 * generated XML.
 */

const REV_DATE = "2026-01-02T03:04:05Z";

/** Resolve a single-section document holding these children, then compile it
 *  back. Returns both legs for route/field assertions. */
function roundTrip(children: SectionChild[]) {
  const doc: DocumentOptions = { sections: [{ children }] };
  const json = resolveDocument(doc, docxExtensions);
  const out = compileDocument(json, docxExtensions);
  return { json, compiled: out.sections[out.sections.length - 1].children };
}

/** The same resolve → compile legs but through real generated XML, so the
 *  office-open shapes are canonical (generateDocument → parseDocument). */
function throughXml(children: SectionChild[]) {
  const binary = generateDocumentSync({ sections: [{ children }] });
  const parsed = parseDocumentSync(new Uint8Array(binary as Buffer));
  const json = resolveDocument(parsed, docxExtensions);
  const out = compileDocument(json, docxExtensions);
  return { json, compiled: out.sections[out.sections.length - 1].children };
}

/** Every text node's formatChange mark in the resolved JSON. */
function formatChangeMarks(json: JSONContent): { mark: JSONContent; text: string }[] {
  const out: { mark: JSONContent; text: string }[] = [];
  const walk = (node: JSONContent): void => {
    if (node.type === "text") {
      for (const mark of node.marks ?? []) {
        if (mark.type === "formatChange") out.push({ mark, text: node.text ?? "" });
      }
    }
    for (const child of node.content ?? []) walk(child);
  };
  walk(json);
  return out;
}

/** The first compiled paragraph's children. */
function paragraphChildren(compiled: SectionChild[]): ParagraphChild[] {
  const para = compiled[0] as { paragraph: { children?: ParagraphChild[] } };
  return para.paragraph.children ?? [];
}

describe("w:rPrChange (run format revisions)", () => {
  it("resolves the old props into a formatChange mark and compiles them back", () => {
    const { json, compiled } = roundTrip([
      {
        paragraph: {
          children: [
            { text: "styled", revision: { id: 7, author: "Ada", date: REV_DATE, bold: true } },
          ],
        },
      },
    ]);
    const marks = formatChangeMarks(json);
    expect(marks).toHaveLength(1);
    expect(marks[0]!.mark.attrs).toEqual({
      id: 7,
      author: "Ada",
      date: REV_DATE,
      props: JSON.stringify({ bold: true }),
    });
    // TextStyle never carries the revision itself — one carrier, one w:rPrChange.
    const textStyle = json.content?.[0]?.content?.[0]?.marks?.find((m) => m.type === "textStyle");
    expect(textStyle?.attrs?.revision ?? null).toBeNull();
    const run = paragraphChildren(compiled)[0] as RunOptions;
    expect(run.text).toBe("styled");
    expect(run.revision).toEqual({ id: 7, author: "Ada", date: REV_DATE, bold: true });
  });

  it("survives real XML round-trip", () => {
    const { json, compiled } = throughXml([
      {
        paragraph: {
          children: [
            {
              text: "styled",
              revision: { id: 9, author: "Bob", date: REV_DATE, italic: true, color: "FF0000" },
            },
          ],
        },
      },
    ]);
    const marks = formatChangeMarks(json);
    expect(marks).toHaveLength(1);
    expect(marks[0]!.mark.attrs?.author).toBe("Bob");
    expect(JSON.parse(marks[0]!.mark.attrs?.props as string)).toEqual({
      italic: true,
      color: "FF0000",
    });
    const run = paragraphChildren(compiled)[0] as RunOptions;
    expect(run.revision).toMatchObject({ id: 9, author: "Bob", italic: true, color: "FF0000" });
  });

  it("keeps the run's current formatting beside the revision", () => {
    const { json, compiled } = roundTrip([
      {
        paragraph: {
          children: [
            {
              text: "both",
              bold: true,
              revision: { id: 1, author: "A", date: REV_DATE },
            },
          ],
        },
      },
    ]);
    const marks = formatChangeMarks(json);
    expect(marks).toHaveLength(1);
    expect(marks[0]!.mark.attrs?.props).toBe("{}");
    const markTypes = (json.content?.[0]?.content?.[0]?.marks ?? []).map((m) => m.type);
    expect(markTypes).toContain("bold");
    const run = paragraphChildren(compiled)[0] as RunOptions;
    expect(run.bold).toBe(true);
    expect(run.revision).toMatchObject({ id: 1, author: "A" });
  });
});

describe("w:pPrChange (paragraph format revisions)", () => {
  it("round-trips the paragraph revision attr through DocumentOptions", () => {
    const { json, compiled } = roundTrip([
      {
        paragraph: {
          alignment: "center",
          children: [{ text: "aligned" }],
          revision: { id: 4, author: "Ada", date: REV_DATE, alignment: "left" },
        },
      },
    ]);
    const attrs = json.content?.[0]?.attrs as { alignment?: string; revision?: unknown };
    expect(attrs.alignment).toBe("center");
    expect(attrs.revision).toEqual({
      id: 4,
      author: "Ada",
      date: REV_DATE,
      alignment: "left",
    });
    const para = compiled[0] as {
      paragraph: { alignment?: string; revision?: unknown; children?: unknown[] };
    };
    expect(para.paragraph.alignment).toBe("center");
    expect(para.paragraph.revision).toEqual({
      id: 4,
      author: "Ada",
      date: REV_DATE,
      alignment: "left",
    });
  });

  it("survives real XML round-trip", () => {
    const { json, compiled } = throughXml([
      {
        paragraph: {
          alignment: "right",
          children: [{ text: "moved" }],
          revision: { id: 11, author: "Bob", date: REV_DATE, alignment: "left" },
        },
      },
    ]);
    const attrs = json.content?.[0]?.attrs as { revision?: { alignment?: string } };
    expect(attrs.revision?.alignment).toBe("left");
    const para = compiled[0] as {
      paragraph: { alignment?: string; revision?: { author?: string; alignment?: string } };
    };
    expect(para.paragraph.revision?.author).toBe("Bob");
    expect(para.paragraph.revision?.alignment).toBe("left");
  });
});

describe("run-prop mark helpers", () => {
  it("lists the rPr marks and excludes the revision carrier", () => {
    const names = formatMarkNames();
    expect(names).toContain("bold");
    expect(names).toContain("textStyle");
    expect(names).not.toContain("formatChange");
    expect(names).not.toContain("insertion");
  });

  it("reconstructs the old marks from the stored props (reject restore)", () => {
    const marks = runPropsToMarks({ bold: true, color: "FF0000", size: 14 });
    const types = marks.map((m) => m.type);
    expect(types).toContain("bold");
    expect(types).toContain("textStyle");
    const textStyle = marks.find((m) => m.type === "textStyle");
    expect(textStyle?.attrs).toMatchObject({ color: "FF0000", size: 14 });
    // The inverse direction merges them back into one rPr object.
    expect(runPropsFromMarks(marks)).toMatchObject({ bold: true, color: "FF0000", size: 14 });
  });

  it("reconstructs nothing for an empty snapshot (plain before-state)", () => {
    expect(runPropsToMarks({})).toEqual([]);
  });
});
