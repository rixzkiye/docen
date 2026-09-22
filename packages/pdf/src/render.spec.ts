// @vitest-environment node
import { execFileSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { Document, Paragraph, SectionBreak } from "@docen/docx";
import { Editor, Node as TextNode } from "@docen/docx/core";
import {
  browserFontMetrics,
  clearMissingFontWarnings,
  createMeasurer,
  getShapingFontManager,
  ShapedMeasurer,
} from "@docen/layout";
import { describe, expect, it, vi } from "vitest";

import { renderPdf } from "./render";

const DUMMY_PNG_1X1 =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const DUMMY_JPEG_1X1 =
  "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wgALCAABAAEBAREA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxA=";

describe("renderPdf headless entry point", () => {
  it("renders a simple JSON document to valid PDF bytes without DOM", async () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [{ type: "text", text: "Headless PDF Generation" }],
        },
        {
          type: "paragraph",
          content: [
            { type: "text", text: "This PDF was rendered entirely in pure Node.js using " },
            {
              type: "text",
              marks: [{ type: "bold" }],
              text: "nodeKit and PaintKit.",
            },
          ],
        },
      ],
    };

    const pdfBytes = await renderPdf(doc, {
      title: "Headless Test",
    });

    expect(pdfBytes).toBeInstanceOf(Uint8Array);
    expect(pdfBytes.length).toBeGreaterThan(500);

    const header = new TextDecoder().decode(pdfBytes.subarray(0, 8));
    expect(header).toContain("%PDF-1.4");

    const tail = new TextDecoder().decode(pdfBytes.subarray(-100));
    expect(tail).toContain("%%EOF");
  });

  it("registers the bundled production faces before layout (zero missing-font warnings)", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      clearMissingFontWarnings();
      await renderPdf({
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [{ type: "text", text: "AVATAR office fi fl" }],
          },
        ],
      });
      const missingFontWarnings = warn.mock.calls
        .map((call) => String(call[0]))
        .filter((message) => message.includes("no face is registered"));
      expect(missingFontWarnings).toEqual([]);

      // The shared shaping manager holds the bundled families and the default
      // measurer really shapes them — not the canvas shim fallback.
      for (const family of ["Calibri", "Cambria", "Arial", "Times New Roman"]) {
        expect(getShapingFontManager().getActiveFont(family), family).toBeDefined();
      }
      const measurer = createMeasurer(browserFontMetrics);
      expect(measurer).toBeInstanceOf(ShapedMeasurer);
      const run = (measurer as ShapedMeasurer).glyphRunOf("AVATAR office fi fl", {
        family: "Calibri",
        sizePx: 16,
      });
      expect(run).toBeDefined();
      expect(run!.totalAdvancePx).toBeGreaterThan(0);
    } finally {
      warn.mockRestore();
    }
  }, 30000);

  it("extracts text byte-exact with pdftotext from headless render", async () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [{ type: "text", text: "Architecture Overview" }],
        },
        {
          type: "paragraph",
          content: [
            {
              type: "text",
              text: "PaintKit decouples document layout painting from browser canvas.",
            },
          ],
        },
      ],
    };

    const pdfBytes = await renderPdf(doc, {
      title: "Architecture",
    });

    const tmpPdf = path.join(os.tmpdir(), `headless-pdftotext-${Date.now()}.pdf`);
    fs.writeFileSync(tmpPdf, pdfBytes);

    try {
      const textOutput = execFileSync("pdftotext", [tmpPdf, "-"], {
        encoding: "utf-8",
      });
      expect(textOutput).toContain("Architecture Overview");
      expect(textOutput).toContain(
        "PaintKit decouples document layout painting from browser canvas.",
      );
    } finally {
      if (fs.existsSync(tmpPdf)) fs.unlinkSync(tmpPdf);
    }
  });

  it("renders documents with images via NodeImageCache", async () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Document with embedded PNG image:" }],
        },
        {
          type: "image",
          attrs: {
            src: DUMMY_PNG_1X1,
            width: 100,
            height: 100,
          },
        },
        {
          type: "paragraph",
          content: [{ type: "text", text: "Document with embedded JPEG image:" }],
        },
        {
          type: "image",
          attrs: {
            src: DUMMY_JPEG_1X1,
            width: 100,
            height: 100,
          },
        },
      ],
    };

    const pdfBytes = await renderPdf(doc, {
      title: "Image Test",
    });

    expect(pdfBytes.length).toBeGreaterThan(1000);
    const text = new TextDecoder("latin1").decode(pdfBytes);
    expect(text).toContain("/XObject");
  }, 30000); // image decode + full render: headroom under full-suite load

  it("renders an SVG image through its raster fallbackSrc", async () => {
    const svg =
      '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10">' +
      '<rect width="10" height="10" fill="#ff0000"/></svg>';
    const svgSrc = `data:image/svg+xml;base64,${Buffer.from(svg, "utf8").toString("base64")}`;
    const doc = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Document with an SVG image and its raster fallback:" }],
        },
        {
          type: "image",
          attrs: {
            src: svgSrc,
            fallbackSrc: DUMMY_PNG_1X1,
            width: 100,
            height: 100,
          },
        },
      ],
    };

    const pdfBytes = await renderPdf(doc, { title: "SVG Fallback" });

    // The headless renderer has no SVG rasterizer: the fallback PNG decodes
    // into an embedded image XObject (the SVG primary would paint the empty
    // placeholder frame instead).
    const text = new TextDecoder("latin1").decode(pdfBytes);
    expect(text).toContain("/Subtype /Image");
    expect(text).toContain("/ColorSpace /DeviceRGB");
  }, 30000);

  it("renders documents with PDF/A-2b and PDF/UA-1 conformance", async () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [{ type: "text", text: "Accessible PDF Document" }],
        },
        {
          type: "paragraph",
          content: [{ type: "text", text: "Compliant with ISO 19005-2 and ISO 14289-1." }],
        },
      ],
    };

    const pdfBytes = await renderPdf(doc, {
      pdfa: "2b",
      pdfUa: true,
      metadata: {
        title: "Accessible PDF",
        author: "Docen Team",
      },
    });

    const text = new TextDecoder("latin1").decode(pdfBytes);
    expect(text).toContain("<pdfaExtension:schemas>");
    expect(text).toContain("<pdfaProperty:name>part</pdfaProperty:name>");
    expect(text).toContain("/OutputIntent");
    expect(text).toContain("/StructTreeRoot");
    expect(text).toContain("/DisplayDocTitle true");

    let verapdfPath: string | null = null;
    try {
      execFileSync("verapdf", ["--version"], { stdio: "ignore" });
      verapdfPath = "verapdf";
    } catch {
      if (fs.existsSync("/home/rixzkiye/.local/bin/verapdf")) {
        verapdfPath = "/home/rixzkiye/.local/bin/verapdf";
      }
    }

    if (!verapdfPath && process.env.DOCEN_REQUIRE_VERAPDF === "1") {
      throw new Error(
        "[docen/pdf] DOCEN_REQUIRE_VERAPDF=1 but the veraPDF CLI is not available on PATH " +
          "or at /home/rixzkiye/.local/bin/verapdf",
      );
    }
    if (!verapdfPath) {
      console.warn(
        "[docen/pdf] veraPDF CLI not found (PATH or /home/rixzkiye/.local/bin/verapdf) — " +
          "skipping the PDF/A-2b + PDF/UA-1 conformance validation. " +
          "Set DOCEN_REQUIRE_VERAPDF=1 to make a missing veraPDF a hard failure.",
      );
    }

    if (verapdfPath) {
      const tmpPdf = path.join(os.tmpdir(), `headless-verapdf-${Date.now()}.pdf`);
      fs.writeFileSync(tmpPdf, pdfBytes);
      try {
        const out2b = execFileSync(verapdfPath, ["--format", "text", "--flavour", "2b", tmpPdf], {
          encoding: "utf-8",
        });
        expect(out2b).toContain("PASS");

        const outUa = execFileSync(verapdfPath, ["--format", "text", "--flavour", "ua1", tmpPdf], {
          encoding: "utf-8",
        });
        expect(outUa).toContain("PASS");
      } finally {
        if (fs.existsSync(tmpPdf)) fs.unlinkSync(tmpPdf);
      }
    }
  }, 30000);

  it("verifies memory discipline: sustained loop of 50 documents stays flat", async () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [{ type: "text", text: "Performance & Memory Benchmark" }],
        },
        {
          type: "paragraph",
          content: [
            {
              type: "text",
              text: "Testing RSS memory overhead and transient heap allocation per document.",
            },
          ],
        },
      ],
    };

    // Warm-up run
    await renderPdf(doc);
    if (typeof global.gc === "function") global.gc();

    const initialRss = process.memoryUsage().rss / (1024 * 1024);

    const iterations = 50;
    for (let i = 0; i < iterations; i++) {
      const bytes = await renderPdf(doc);
      expect(bytes.length).toBeGreaterThan(500);
    }

    if (typeof global.gc === "function") global.gc();
    const finalRss = process.memoryUsage().rss / (1024 * 1024);
    const rssDelta = finalRss - initialRss;

    console.log(
      `[Memory Benchmark] Initial RSS: ${initialRss.toFixed(2)} MB, Final RSS: ${finalRss.toFixed(2)} MB, Delta: ${rssDelta.toFixed(2)} MB over ${iterations} docs`,
    );

    // Memory should stay flat (less than 30 MB delta across 50 iterations)
    expect(rssDelta).toBeLessThan(30);
  }, 30000);
});

describe("section break semantics reach the headless flow", () => {
  const pageCount = (bytes: Uint8Array): number => {
    const text = new TextDecoder("latin1").decode(bytes);
    return (text.match(/\/Type\s*\/Page(?![a-zA-Z])/g) ?? []).length;
  };

  const twoSections = (type: "continuous" | "oddPage") => ({
    type: "doc" as const,
    // The following (final) section's sectPr declares how IT starts; a break
    // type on section A would make A itself start on that parity.
    attrs: { sectionProperties: { type } },
    content: [
      {
        type: "paragraph",
        attrs: { sectionProperties: {} },
        content: [{ type: "text", text: "Section A" }],
      },
      { type: "paragraph", content: [{ type: "text", text: "Section B" }] },
    ],
  });

  it("keeps a continuous section on the same page", async () => {
    const bytes = await renderPdf(twoSections("continuous"), { title: "Continuous" });
    expect(pageCount(bytes)).toBe(1);
  }, 30000);

  it("inserts Word's blank interleave for an oddPage section start", async () => {
    // Section A occupies physical page 1 (odd). An oddPage start must land on
    // page 3, so Word inserts blank page 2 — the Node pipeline must project the
    // sectPr type into the flow to produce the same three pages.
    const bytes = await renderPdf(twoSections("oddPage"), { title: "Odd Page" });
    expect(pageCount(bytes)).toBe(3);
  }, 30000);

  it("paginates the UI command's output the same way", async () => {
    const Text = TextNode.create({ name: "text", group: "inline" });
    const commandDoc = (type?: "continuous" | "oddPage") => {
      const editor = new Editor({
        element: null,
        extensions: [Document, Paragraph, Text, SectionBreak],
        content: {
          type: "doc",
          content: [
            { type: "paragraph", content: [{ type: "text", text: "Section A" }] },
            { type: "paragraph", content: [{ type: "text", text: "Section B" }] },
          ],
        },
      });
      // element:null skips mount() and with it plugin installation.
      for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
      editor.commands.setSectionBreak(type ? { type } : undefined);
      return editor.getJSON();
    };

    const continuous = await renderPdf(commandDoc("continuous"), {
      title: "Command Continuous",
    });
    expect(pageCount(continuous)).toBe(1);

    // Before the type moved to the FOLLOWING section this rendered 2 pages:
    // section A claimed the oddPage start instead of the inserted section.
    const odd = await renderPdf(commandDoc("oddPage"), { title: "Command Odd Page" });
    expect(pageCount(odd)).toBe(3);
  }, 60000);
});

describe("headless PDF structure (outline + page labels)", () => {
  it("derives /Outlines and /PageLabels from the projected document", async () => {
    const doc = {
      type: "doc",
      attrs: {
        sectionProperties: { pageNumberType: { start: 1, format: "lowerRoman" } },
      },
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [{ type: "text", text: "Chapter One" }],
        },
        { type: "paragraph", content: [{ type: "text", text: "Body text." }] },
        {
          type: "heading",
          attrs: { level: 2 },
          content: [{ type: "text", text: "Section 1.1" }],
        },
        { type: "paragraph", content: [{ type: "text", text: "More text." }] },
      ],
    };

    const pdfBytes = await renderPdf(doc, { title: "Structure" });
    const text = new TextDecoder("latin1").decode(pdfBytes);

    // The outline root + first heading entry come from the laid blocks (the
    // server view has no PM document).
    expect(text).toMatch(/\/Type \/Outlines \/First \d+ 0 R/);
    expect(text).toContain("(Chapter One)");
    expect(text).toContain("(Section 1.1)");
    // The section's w:pgNumType maps to the roman page-label style.
    expect(text).toContain("/PageLabels");
    expect(text).toContain("/S /r");
  }, 30000);
});
