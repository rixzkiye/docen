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

/** The layout's own right-edge target per line (pt) — the interval the
 *  painter's `width` prop comes from, independent of the node lane. */
function layoutTargetsPt(options: DocOptions, measurer: TextMeasurer): number[] {
  const { sections } = projectDocumentOptions(compileDocument(buildDoc(options) as never));
  const section = sections[0]!;
  const { pages } = layoutFlowSections([{ blocks: section.blocks, opts: section.flow }], measurer);
  const first = pages[0]!.items.find((item) => item.block.kind === "paragraph")!;
  const para = first.block as LaidOutParagraph;
  const x = section.flow.contentLeftPx + (first.xPx ?? 0) + (para.indent?.leftPx ?? 0);
  return para.lines.map((line) => (x + (line.xOffsetPx ?? 0) + (line.maxWidthPx ?? 0)) * 0.75);
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
