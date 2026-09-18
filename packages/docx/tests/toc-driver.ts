// Fixture driver for the toc-oracles matrix (item 8): discovery, model
// loading, generation with the fixture's field context, and the expectation
// shape consumed by the generation and LibreOffice oracle specs.
import { execFileSync } from "node:child_process";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

import { unzipSync } from "@office-open/core";

import type { JSONContent } from "../src/core";
import { generateDOCXSync, resolveDocument, type FieldCacheOptions } from "../src/index";

/** The fixture root (this file lives in packages/docx/tests). */
export const TOC_FIXTURES_ROOT = join(import.meta.dirname, "toc-fixtures");

export interface FixtureFieldContext {
  pageCount?: number;
  pageOf?: Record<string, number>;
  bookmarkPages?: Record<string, number>;
  sectionOf?: Record<string, number>;
  sectionPages?: number;
  tocPageOf?: Record<string, number>;
}

export interface ExpectedTocEntry {
  style: string;
  text: string;
  page: string;
}

export interface PythonExpectation {
  simple?: [string, string][];
  instrTexts?: string[];
  tocEntries?: ExpectedTocEntry[];
  parts?: Record<string, { simple?: [string, string][]; instrTexts?: string[] }>;
  headers?: { instrTexts?: string[] };
  footers?: { instrTexts?: string[] };
}

export interface PdfExpectation {
  contains?: string[];
  patterns?: string[];
  absent?: string[];
}

export interface FixtureExpectation {
  description: string;
  fields?: FixtureFieldContext;
  python?: PythonExpectation;
  xmlPresent?: string[];
  xmlAbsent?: string[];
  tocSliceAbsent?: string[];
  pdf?: PdfExpectation;
  update?: PdfExpectation & { skip?: string };
  resavePython?: PythonExpectation;
  /** Python checks for the *updated* package; defaults to `resavePython`.
   *  An empty object means "a deliberate update changes this fixture's caches"
   *  (e.g. a non-empty TOC recomputed by Word's F9). */
  updatePython?: PythonExpectation;
  loLimitations?: string[];
}

export interface TocFixture {
  name: string;
  dir: string;
  model: JSONContent;
  expected: FixtureExpectation;
}

const isTiptapDocument = (model: unknown): model is JSONContent =>
  typeof model === "object" && model !== null && (model as { type?: unknown }).type === "doc";

/** Every fixture directory under the root, sorted by name. */
export function discoverTocFixtures(root: string = TOC_FIXTURES_ROOT): TocFixture[] {
  return readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && !entry.name.startsWith("."))
    .map((entry) => entry.name)
    .sort()
    .map((name) => {
      const dir = join(root, name);
      const model = JSON.parse(readFileSync(join(dir, "model.json"), "utf8")) as unknown;
      const expected = JSON.parse(
        readFileSync(join(dir, "expected.json"), "utf8"),
      ) as FixtureExpectation;
      return {
        name,
        dir,
        model: isTiptapDocument(model) ? model : resolveDocument(model as never),
        expected,
      };
    });
}

/** The `fields` context a fixture's expectations describe. */
export function fixtureFieldOptions(expected: FixtureExpectation): FieldCacheOptions | undefined {
  const fields = expected.fields;
  if (!fields) return undefined;
  const pageOf = fields.pageOf ?? {};
  const bookmarks = fields.bookmarkPages ?? {};
  const sections = fields.sectionOf ?? {};
  const tocPages = fields.tocPageOf ?? {};
  return {
    ...(fields.pageCount != null ? { pageCount: fields.pageCount } : {}),
    ...(Object.keys(pageOf).length > 0 || Object.keys(bookmarks).length > 0
      ? {
          pageOf: ({ index, bookmark }: { index: number; bookmark?: string }) =>
            bookmark != null ? bookmarks[bookmark] : pageOf[index],
        }
      : {}),
    ...(Object.keys(sections).length > 0
      ? { sectionOf: ({ index }: { index: number }) => sections[index] }
      : {}),
    ...(fields.sectionPages != null ? { sectionPages: fields.sectionPages } : {}),
    ...(Object.keys(tocPages).length > 0
      ? {
          tocPageOf: ({ index, kind }: { index: number; kind: "heading" | "caption" }) =>
            tocPages[`${kind}:${index}`],
        }
      : {}),
  };
}

/** Generate one fixture to bytes (with its field context). */
export function generateFixture(fixture: TocFixture): Uint8Array {
  return generateDOCXSync(fixture.model, {
    prepare: false,
    fields: fixtureFieldOptions(fixture.expected),
  }) as Uint8Array;
}

export interface PythonReport {
  simple: [string, string][];
  complex: [string, string][];
  instrTexts: string[];
  tocEntries: ExpectedTocEntry[];
  paragraphs: string[];
  parts: Record<string, { simple: [string, string][]; instrTexts: string[]; paragraphs: string[] }>;
}

/** Run the checked-in python-docx/lxml extractor over DOCX paths. */
export function runPython(paths: string[]): Record<string, PythonReport> {
  const out = execFileSync("python3", [join(import.meta.dirname, "toc-oracle.py"), ...paths], {
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  });
  return JSON.parse(out) as Record<string, PythonReport>;
}

export const documentXmlOf = (bytes: Uint8Array): string =>
  new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]);

/** Concatenated TOC sdt slices of `document.xml` (entry assertions must not
 *  match the body headings the entries quote). */
export function tocSlicesOf(xml: string): string {
  let out = "";
  let cursor = 0;
  for (;;) {
    const start = xml.indexOf("<w:sdt>", cursor);
    if (start < 0) break;
    const end = xml.indexOf("</w:sdt>", start);
    if (end < 0) break;
    out += xml.slice(start, end + 8);
    cursor = end + 8;
  }
  return out;
}

/** True when the fixture's external tools are available. */
export const haveSoffice = (): boolean =>
  existsSync("/usr/bin/soffice") || existsSync("/usr/local/bin/soffice");

/** Expected instructions missing from `actual`: every whitespace-separated
 *  token of an expected string must appear inside one actual instruction
 *  (LibreOffice reorders switches — `\h \o "1-3"` becomes `\o "1-3" \h` — so
 *  the comparison is per token, still anchored to a single field). */
export function expectationMissing(expected: string[], actual: string[]): string[] {
  return expected.filter((needle) => {
    const tokens = needle.split(/\s+/).filter(Boolean);
    return !actual.some((value) => tokens.every((token) => value.includes(token)));
  });
}

/** Toc entries (text+page) expected but missing from the actual list. */
export function tocEntriesMissing(
  expected: ExpectedTocEntry[],
  actual: ExpectedTocEntry[],
): string[] {
  return expected
    .filter(
      (entry) =>
        !actual.some((candidate) => candidate.text === entry.text && candidate.page === entry.page),
    )
    .map((entry) => `${entry.text}@${entry.page}`);
}
