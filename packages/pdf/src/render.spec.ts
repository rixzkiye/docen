// @vitest-environment node
import { execFileSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { describe, expect, it } from "vitest";

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
