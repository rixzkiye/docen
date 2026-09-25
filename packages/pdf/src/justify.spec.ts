// @vitest-environment node
// W31 regression: a justified (w:jc both) paragraph must paint flush to the
// line's stretch interval in the headless Text lane. Before the fix the node
// kit ignored `textAlign` "both-justify"/"both-letter", so every line ended at
// its natural advance — ragged right edges in the official PDF while the
// browser canvas and the DOCX (`Normal jc=both`) were correct.
import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { compileDocument } from "@docen/docx";
import { projectDocumentOptions } from "@docen/docx/layout";
import {
  browserFontMetrics,
  createMeasurer,
  layoutFlowSections,
  type LaidOutParagraph,
  type ShapedMeasurer,
  type TextMeasurer,
} from "@docen/layout";
import { describe, expect, it } from "vitest";

import { renderPdf } from "./render";

const TEXT =
  "lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur excepteur sint occaecat cupidatat non proident sunt in culpa qui officia deserunt mollit anim id est laborum";

interface DocOptions {
  readonly alignment?: "both";
  readonly defaults?: boolean;
}

function buildDoc({ alignment, defaults }: DocOptions = {}) {
  return {
    type: "doc",
    attrs: defaults
      ? { styles: { default: { document: { paragraph: { alignment: "both" } } } } }
      : {},
    content: [
      {
        type: "paragraph",
        ...(alignment ? { attrs: { alignment } } : {}),
        content: [
          {
            type: "text",
            text: TEXT,
            marks: [{ type: "textStyle", attrs: { font: "Times New Roman" } }],
          },
        ],
      },
    ],
  };
}

/** Shaped advance metrics with the shaped-run lane hidden: the painter paints
 *  this document through Text nodes — the node lane W31 fixes (runs the
 *  shaper does not outline, and the lane the production repro's PDF used) —
 *  instead of glyph outlines. */
function textLaneMeasurer(): TextMeasurer {
  const measurer = createMeasurer(browserFontMetrics);
  (measurer as ShapedMeasurer).glyphRunOf = () => undefined;
  return measurer;
}

/** Exactly `renderPdf`'s default (`options.measurer` unset): shaping on, so
 *  the painter takes the glyph-outline lane and the layout spans are the
 *  extraction layer. */
function productionMeasurer(): TextMeasurer {
  return createMeasurer(browserFontMetrics);
}

/** Per-line rightmost word edge (pt), from the PDF's own painted glyphs. */
function rightEdges(pdf: Uint8Array): number[] {
  const dir = mkdtempSync(join(tmpdir(), "docen-justify-"));
  const file = join(dir, "doc.pdf");
  writeFileSync(file, pdf);
  const html = execFileSync("pdftotext", ["-bbox-layout", file, "-"], { encoding: "utf8" });
  return [...html.matchAll(/<line [^>]*xMax="([\d.]+)"/g)].map((m) => Number(m[1]));
}

/** Per-line words extracted from the PDF's own text layer (`-bbox-layout`).
 *  Words are grouped by baseline: a heavily justified line's gaps make
 *  poppler emit one `<line>` per word, so the `<line>` wrapper is not a
 *  reliable grouping here. */
interface ExtractedLine {
  xMax: number;
  words: { text: string; xMin: number; xMax: number; width: number }[];
}

function linesOf(pdf: Uint8Array): ExtractedLine[] {
  const dir = mkdtempSync(join(tmpdir(), "docen-justify-lines-"));
  const file = join(dir, "doc.pdf");
  writeFileSync(file, pdf);
  const html = execFileSync("pdftotext", ["-bbox-layout", file, "-"], { encoding: "utf8" });
  const byBaseline = new Map<string, ExtractedLine["words"]>();
  for (const m of html.matchAll(
    /<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="[\d.]+">([^<]*)<\/word>/g,
  )) {
    const xMin = Number(m[1]);
    const xMax = Number(m[3]);
    const key = Number(m[2]).toFixed(1);
    const line = byBaseline.get(key) ?? [];
    line.push({ text: m[4]!, xMin, xMax, width: xMax - xMin });
    byBaseline.set(key, line);
  }
  return [...byBaseline.values()].map((words) => ({
    xMax: Math.max(...words.map((w) => w.xMax)),
    words,
  }));
}

/** The layout's own right-edge target per line (pt) — the interval the
 *  painter's `width` prop comes from, independent of the node lane. */
function layoutTargetsOf(doc: unknown, measurer: TextMeasurer): number[] {
  const { sections } = projectDocumentOptions(compileDocument(doc as never));
  const section = sections[0]!;
  const { pages } = layoutFlowSections([{ blocks: section.blocks, opts: section.flow }], measurer);
  const first = pages[0]!.items.find((item) => item.block.kind === "paragraph")!;
  const para = first.block as LaidOutParagraph;
  const x = section.flow.contentLeftPx + (first.xPx ?? 0) + (para.indent?.leftPx ?? 0);
  return para.lines.map((line) => (x + (line.xOffsetPx ?? 0) + (line.maxWidthPx ?? 0)) * 0.75);
}

function layoutTargetsPt(options: DocOptions, measurer: TextMeasurer): number[] {
  return layoutTargetsOf(buildDoc(options), measurer);
}

/** A paragraph whose first line is short enough that justification must spread
 *  a large slack — the production repro (a line ending in a long unbreakable
 *  token; the marker/footnote case leaves the same shape). */
const STRETCHED_TEXT = "dan cinta damai. Mubarok " + "Mubarok".repeat(9);

function buildStretchedDoc(alignment?: "both") {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        ...(alignment ? { attrs: { alignment } } : {}),
        content: [
          {
            type: "text",
            text: STRETCHED_TEXT,
            marks: [{ type: "textStyle", attrs: { font: "Times New Roman" } }],
          },
        ],
      },
    ],
  };
}

/** CJK text with word gaps, long enough to wrap: the per-grapheme
 *  (both-letter) justification model. */
const CJK_TEXT = "天地玄黃宇宙洪荒 日月盈昃辰宿列張 寒來暑往秋收冬藏 閏餘成歲律呂調陽 ".repeat(3);

function buildCjkDoc(alignment?: "both") {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        ...(alignment ? { attrs: { alignment } } : {}),
        content: [
          {
            type: "text",
            text: CJK_TEXT,
            marks: [{ type: "textStyle", attrs: { font: "SimSun" } }],
          },
        ],
      },
    ],
  };
}

describe("justified paragraphs fill the layout interval (W31)", () => {
  it("fills every full line of a justified paragraph (attr and docDefaults)", async () => {
    for (const variant of [{ alignment: "both" as const }, { defaults: true }]) {
      const measurer = textLaneMeasurer();
      const edges = rightEdges(await renderPdf(buildDoc(variant) as never, { measurer }));
      const targets = layoutTargetsPt(variant, measurer);
      expect(edges).toHaveLength(targets.length);
      const full = edges.length - 1;
      for (let i = 0; i < full; i++) {
        expect(Math.abs(edges[i]! - targets[i]!), `line ${i}`).toBeLessThanOrEqual(1);
      }
      // The paragraph's last line is not justified — it stays at its natural
      // advance, well inside the interval.
      expect(targets[full]! - edges[full]!).toBeGreaterThan(5);
    }
  }, 120_000);

  it("leaves a left-aligned paragraph ragged (no regression)", async () => {
    const measurer = textLaneMeasurer();
    const edges = rightEdges(await renderPdf(buildDoc() as never, { measurer }));
    const targets = layoutTargetsPt({}, measurer);
    expect(edges).toHaveLength(targets.length);
    const full = edges.length - 1;
    for (let i = 0; i < full; i++) {
      expect(targets[i]! - edges[i]!, `line ${i}`).toBeGreaterThan(1);
    }
  }, 120_000);

  it("does not stretch the paragraph's last line", async () => {
    const measurer = textLaneMeasurer();
    const justified = rightEdges(
      await renderPdf(buildDoc({ alignment: "both" }) as never, { measurer }),
    );
    const left = rightEdges(await renderPdf(buildDoc() as never, { measurer }));
    expect(justified.at(-1)!).toBeCloseTo(left.at(-1)!, 0);
    expect(justified.at(-1)!).toBeLessThan(justified[0]! - 5);
  }, 120_000);
});

describe("justified paragraphs fill the layout interval with the production measurer", () => {
  it("fills every full line (shaped glyph-outline lane, attr and docDefaults)", async () => {
    for (const variant of [{ alignment: "both" as const }, { defaults: true }]) {
      const measurer = productionMeasurer();
      const edges = rightEdges(await renderPdf(buildDoc(variant) as never, { measurer }));
      const targets = layoutTargetsPt(variant, measurer);
      expect(edges).toHaveLength(targets.length);
      const full = edges.length - 1;
      for (let i = 0; i < full; i++) {
        expect(Math.abs(edges[i]! - targets[i]!), `line ${i}`).toBeLessThanOrEqual(1);
      }
      // Last line stays natural with the production lane too.
      expect(targets[full]! - edges[full]!).toBeGreaterThan(5);
    }
  }, 120_000);

  it("leaves a left-aligned paragraph ragged (shaped glyph-outline lane)", async () => {
    const measurer = productionMeasurer();
    const edges = rightEdges(await renderPdf(buildDoc() as never, { measurer }));
    const targets = layoutTargetsPt({}, measurer);
    expect(edges).toHaveLength(targets.length);
    const full = edges.length - 1;
    for (let i = 0; i < full; i++) {
      expect(targets[i]! - edges[i]!, `line ${i}`).toBeGreaterThan(1);
    }
  }, 120_000);
});

describe("justified word advances stay natural (W35)", () => {
  it("distributes a Latin line's slack through its spaces (shaped outline lane)", async () => {
    const measurer = productionMeasurer();
    const justified = linesOf(await renderPdf(buildStretchedDoc("both") as never, { measurer }));
    const left = linesOf(await renderPdf(buildStretchedDoc() as never, { measurer }));
    // Justification never changes the breaks: the stretched first line carries
    // the same words in both renders.
    const jw = justified[0]!.words;
    const lw = left[0]!.words;
    expect(jw.length).toBeGreaterThan(2);
    expect(jw.map((w) => w.text)).toEqual(lw.map((w) => w.text));
    // Each word keeps the font's own advance — no letter-spread inside a word.
    for (let i = 0; i < jw.length; i++) {
      expect(Math.abs(jw[i]!.width - lw[i]!.width), jw[i]!.text).toBeLessThan(0.2);
    }
    // ...while the spaces absorb the slack: the line still reaches its
    // interval target and the inter-word gap grew far past the natural one.
    const target = layoutTargetsOf(buildStretchedDoc("both"), measurer)[0]!;
    expect(Math.abs(justified[0]!.xMax - target)).toBeLessThanOrEqual(1);
    const jGap = jw[2]!.xMin - jw[1]!.xMax;
    const lGap = lw[2]!.xMin - lw[1]!.xMax;
    expect(jGap - lGap).toBeGreaterThan(20);
  }, 120_000);

  it("keeps a CJK line's uniform both-letter stretch (not word-gap mode)", async () => {
    const measurer = productionMeasurer();
    const justified = linesOf(await renderPdf(buildCjkDoc("both") as never, { measurer }));
    const left = linesOf(await renderPdf(buildCjkDoc() as never, { measurer }));
    const jw = justified[0]!.words;
    const lw = left[0]!.words;
    expect(jw.length).toBeGreaterThan(2);
    expect(jw.map((w) => w.text)).toEqual(lw.map((w) => w.text));
    // Per-grapheme justification scales the words themselves (a word-gap
    // distribution would leave every word at its natural advance). The line
    // may span several items with slightly different shares, so the ratios
    // only need to stay in one uniform band.
    const ratios = jw.map((w, i) => w.width / lw[i]!.width);
    for (const [i, ratio] of ratios.entries()) {
      expect(ratio, jw[i]!.text).toBeGreaterThan(1.005);
    }
    expect(Math.max(...ratios) - Math.min(...ratios)).toBeLessThan(0.05);
  }, 120_000);
});

describe("footnote-marker lines fill their interval (W36)", () => {
  const run = (text: string) => ({
    type: "text",
    text,
    marks: [{ type: "textStyle", attrs: { font: "Times New Roman" } }],
  });
  const marker = (id: number) => ({
    type: "inlinePassthrough",
    attrs: { data: JSON.stringify({ footnoteReference: id }) },
  });
  const note = (id: number) => ({
    id,
    children: [
      {
        paragraph: {
          style: "FootnoteText",
          children: [
            { style: "FootnoteReference", children: [{ footnoteRef: true }] },
            { text: `Catatan ${id}.` },
          ],
        },
      },
    ],
  });

  /** The production repro: a justified paragraph whose footnote markers sit
   *  mid-sentence; the run after a marker is long enough that its greedy
   *  slice overshoots the remaining line by a hair (the walker's forced
   *  unit), which used to wrap the whole run and leave a ~70pt-gap line. */
  function buildMarkerDoc() {
    return {
      type: "doc",
      attrs: {
        documentExtras: { footnotes: [1, 2, 3, 4, 5].map(note) },
        sectionProperties: {
          pageSize: { width: 11906, height: 16838, orientation: "portrait" },
          pageMargin: {
            top: 2268,
            bottom: 1701,
            left: 2268,
            right: 1701,
            header: 720,
            footer: 720,
          },
        },
        styles: {
          default: {
            document: {
              run: {
                font: { ascii: "Times New Roman", hAnsi: "Times New Roman", cs: "Times New Roman" },
                size: 12,
              },
              paragraph: {
                alignment: "both",
                spacing: { line: 360, lineRule: "multiple", after: 120 },
                indent: { firstLine: 567 },
              },
            },
          },
        },
      },
      content: [
        {
          type: "paragraph",
          content: [
            run(
              "Novel Ayat-Ayat Cinta karya Habiburrahman El Shirazy merupakan salah satu karya sastra Islam populer yang telah banyak dikaji dalam penelitian akademik. Sebagai contoh, Supratno ",
            ),
            marker(1),
            run(" mengkaji nilai multikultural dalam novel tersebut, sedangkan Yasid dan Juhdi "),
            marker(2),
            run(" menelaah Islam sebagai agama toleransi dan cinta damai. Mubarok "),
            marker(3),
            run(
              " menganalisis aspek stilistika dan implikasinya sebagai bahan ajar bahasa Indonesia, sementara Yulianto ",
            ),
            marker(4),
            run(" meneliti konstruksi maskulinitas Islam. Rosdiana "),
            marker(5),
            run(
              " membongkar ideologi wacana pada novel yang sama. Keragaman pendekatan tersebut menunjukkan bahwa novel Ayat-Ayat Cinta memuat lapisan makna yang produktif untuk dianalisis dari sisi nilai, kebahasaan, dan ideologi.",
            ),
          ],
        },
      ],
    };
  }

  it("keeps the following run on the marker's line instead of wrapping it whole", async () => {
    const measurer = productionMeasurer();
    const lines = linesOf(await renderPdf(buildMarkerDoc() as never, { measurer }));
    const line = lines.find((entry) => entry.words.some((word) => word.text === "Mubarok"));
    expect(line).toBeDefined();
    // The marker must not end the line: the next run's prefix joins it.
    expect(line!.words.some((word) => word.text === "menganalisis")).toBe(true);
    // ...and the line's gaps stay normal (the defect stretched them to ~70pt).
    const gaps = line!.words.slice(1).map((word, index) => word.xMin - line!.words[index]!.xMax);
    expect(Math.max(...gaps)).toBeLessThan(20);
  }, 120_000);
});
