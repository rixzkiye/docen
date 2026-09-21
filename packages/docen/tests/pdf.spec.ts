// @vitest-environment node
import { execFileSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { describe, expect, it } from "vitest";

import {
  generatePDF,
  getPdfBackend,
  nodeBackendAdapter,
  registerPdfBackend,
  type PdfBackendAdapter,
} from "../src";

const FIXTURE_PATH = path.resolve(
  __dirname,
  "../../../scripts/fixtures/pdf-render/demo-basic.json",
);

describe("docen generatePDF unified API", () => {
  it("exports nodeBackendAdapter and default registry", () => {
    expect(nodeBackendAdapter.name).toBe("node");
    expect(getPdfBackend("node")).toBe(nodeBackendAdapter);
    expect(getPdfBackend("chromium")).toBeDefined();
  });

  it("renders a simple JSON document to valid PDF bytes via default node backend", async () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { heading: "Heading1" },
          content: [{ type: "text", text: "Docen Unified PDF API" }],
        },
        {
          type: "paragraph",
          content: [
            {
              type: "text",
              text: "Unified generatePDF mirrors generateDOCX with pluggable backends.",
            },
          ],
        },
      ],
    };

    const bytes = await generatePDF(doc, {
      title: "Unified API Test",
    });

    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(500);

    const header = new TextDecoder().decode(bytes.subarray(0, 8));
    expect(header).toContain("%PDF-1.4");

    const tail = new TextDecoder().decode(bytes.subarray(-100));
    expect(tail).toContain("%%EOF");
  });

  it("renders demo-basic fixture end-to-end and validates text with pdftotext", async () => {
    const fixtureJson = JSON.parse(fs.readFileSync(FIXTURE_PATH, "utf-8"));

    const bytes = await generatePDF(fixtureJson, {
      title: "Demo Basic Headless",
    });

    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(2000);

    const tmpPdf = path.join(os.tmpdir(), `docen-fixture-${Date.now()}.pdf`);
    fs.writeFileSync(tmpPdf, bytes);

    try {
      const textOutput = execFileSync("pdftotext", [tmpPdf, "-"], {
        encoding: "utf-8",
      });
      expect(textOutput).toContain("Server PDF Render Fixture");
      expect(textOutput).toContain("Character formatting");
      expect(textOutput).toContain("Deterministic headless export");
    } finally {
      if (fs.existsSync(tmpPdf)) fs.unlinkSync(tmpPdf);
    }
  });

  it("supports PDF/A-2b and PDF/UA-1 conformance options", async () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          attrs: { heading: "Heading1" },
          content: [{ type: "text", text: "ISO Conformance Test" }],
        },
      ],
    };

    const bytes = await generatePDF(doc, {
      pdfa: "2b",
      pdfUa: true,
      metadata: {
        title: "Compliant Document",
        author: "Docen Team",
      },
    });

    const text = new TextDecoder("latin1").decode(bytes);
    expect(text).toContain("%PDF-1.7");
    expect(text).toContain("<pdfaid:part>2</pdfaid:part>");
    expect(text).toContain("<pdfaid:conformance>B</pdfaid:conformance>");
    expect(text).toContain("/DisplayDocTitle true");
  });

  it("supports textMode embedded and outlines", async () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Text mode test" }],
        },
      ],
    };

    const outlinesPdf = await generatePDF(doc, { textMode: "outlines" });
    const embeddedPdf = await generatePDF(doc, { textMode: "embedded" });

    expect(outlinesPdf).toBeInstanceOf(Uint8Array);
    expect(embeddedPdf).toBeInstanceOf(Uint8Array);
    expect(outlinesPdf.length).toBeGreaterThan(0);
    expect(embeddedPdf.length).toBeGreaterThan(0);
  });

  it("supports custom backend adapter registration and dispatch", async () => {
    const mockAdapter: PdfBackendAdapter = {
      name: "mock-backend",
      async render(input) {
        return new TextEncoder().encode(`MOCK-PDF:${JSON.stringify(input)}`);
      },
    };

    registerPdfBackend(mockAdapter);
    expect(getPdfBackend("mock-backend")).toBe(mockAdapter);

    const doc = { type: "doc", content: [] };
    const bytes = await generatePDF(doc, { backend: "mock-backend" });
    const text = new TextDecoder().decode(bytes);
    expect(text).toContain('MOCK-PDF:{"type":"doc","content":[]}');
  });

  it("supports direct inline custom adapter object", async () => {
    const inlineAdapter: PdfBackendAdapter = {
      name: "inline",
      async render() {
        return new Uint8Array([1, 2, 3, 4]);
      },
    };

    const bytes = await generatePDF({ type: "doc", content: [] }, { backend: inlineAdapter });
    expect(Array.from(bytes)).toEqual([1, 2, 3, 4]);
  });

  it("throws clear error on unknown backend", async () => {
    await expect(
      generatePDF({ type: "doc", content: [] }, { backend: "unknown-xyz" as any }),
    ).rejects.toThrow('Unknown PDF backend: "unknown-xyz"');
  });
});
