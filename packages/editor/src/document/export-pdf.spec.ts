// @vitest-environment node
import { execFileSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import type { FlowPage, LaidOutParagraph, LaidOutTable } from "@docen/layout";
import { describe, expect, it } from "vitest";

import type { CanvasStageSection } from "./canvas/stage";
import {
  encodeHexUtf16,
  escapePdfString,
  extractPdfPageLayers,
  pagesToPdf,
  type PdfLinkAnnotation,
  type PdfPageShot,
  type PdfTextSpan,
} from "./export-pdf";

// A minimal valid 1x1 white JPEG image byte sequence
const DUMMY_JPEG = new Uint8Array([
  0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x01, 0x00, 0x48,
  0x00, 0x48, 0x00, 0x00, 0xff, 0xdb, 0x00, 0x43, 0x00, 0x08, 0x06, 0x06, 0x07, 0x06, 0x05, 0x08,
  0x07, 0x07, 0x07, 0x09, 0x09, 0x08, 0x0a, 0x0c, 0x14, 0x0d, 0x0c, 0x0b, 0x0b, 0x0c, 0x19, 0x12,
  0x13, 0x0f, 0x14, 0x1d, 0x1a, 0x1f, 0x1e, 0x1d, 0x1a, 0x1c, 0x1c, 0x20, 0x24, 0x2e, 0x27, 0x20,
  0x22, 0x2c, 0x23, 0x1c, 0x1c, 0x28, 0x37, 0x29, 0x2c, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1f, 0x27,
  0x39, 0x3d, 0x38, 0x32, 0x3c, 0x2e, 0x33, 0x34, 0x32, 0xff, 0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01,
  0x00, 0x01, 0x01, 0x01, 0x11, 0x00, 0xff, 0xc4, 0x00, 0x1f, 0x00, 0x00, 0x01, 0x05, 0x01, 0x01,
  0x01, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04,
  0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f,
  0x00, 0xbf, 0x00, 0xff, 0xd9,
]);

describe("export-pdf string and encoding helpers", () => {
  it("escapes special characters in PDF literal strings", () => {
    expect(escapePdfString("Hello (World) \\ [Test]")).toBe("Hello \\(World\\) \\\\ [Test]");
    expect(escapePdfString("Line 1\nLine 2\rDone")).toBe("Line 1\\nLine 2\\rDone");
  });

  it("encodes UTF-16 code units as 4-character hex CIDs", () => {
    expect(encodeHexUtf16("A")).toBe("0041");
    expect(encodeHexUtf16("Hello")).toBe("00480065006c006c006f");
    expect(encodeHexUtf16("中文")).toBe("4e2d6587");
  });
});

describe("pagesToPdf", () => {
  it("generates a valid PDF with header, xref table, and trailer", async () => {
    const shot: PdfPageShot = {
      width: 800,
      height: 600,
      jpeg: DUMMY_JPEG,
    };
    const blob = await pagesToPdf([shot]);
    expect(blob.type).toBe("application/pdf");
    const text = await blob.text();

    expect(text.startsWith("%PDF-1.4\n")).toBe(true);
    expect(text).toContain("/Type /Catalog");
    expect(text).toContain("/Type /Pages");
    expect(text).toContain("/Type /Page");
    expect(text).toContain("/Type /XObject");
    expect(text).toContain("xref\n");
    expect(text).toContain("trailer\n");
    expect(text).toContain("startxref\n");
    expect(text.trim().endsWith("%%EOF")).toBe(true);
  });

  it("includes document metadata in /Info dictionary", async () => {
    const shot: PdfPageShot = {
      width: 800,
      height: 600,
      jpeg: DUMMY_JPEG,
    };
    const blob = await pagesToPdf([shot], {
      metadata: {
        title: "Annual Report 2026",
        author: "Jane Doe",
      },
    });
    const text = await blob.text();

    expect(text).toContain("/Title (Annual Report 2026)");
    expect(text).toContain("/Author (Jane Doe)");
    expect(text).toContain("/Creator (Docen Word Processor)");
    expect(text).toContain("/Producer (Docen PDF Engine)");
    expect(text).toContain("/CreationDate (D:");
  });

  it("generates invisible searchable text layer with 3 Tr mode and exact positioning", async () => {
    const textSpans: PdfTextSpan[] = [
      {
        text: "Chapter 1: Parity",
        x: 72,
        y: 720,
        width: 150,
        height: 18,
        fontSize: 16,
        bold: true,
      },
      {
        text: "This is a selectable paragraph.",
        x: 72,
        y: 695,
        width: 200,
        height: 14,
        fontSize: 12,
      },
    ];

    const shot: PdfPageShot = {
      width: 800,
      height: 600,
      jpeg: DUMMY_JPEG,
      textSpans,
    };

    const blob = await pagesToPdf([shot]);
    const text = await blob.text();

    // Text rendering mode 3 = invisible text
    expect(text).toContain("3 Tr");
    // Text positioning
    expect(text).toContain("1 0 0 1 72.00 720.00 Tm");
    expect(text).toContain("(Chapter 1: Parity) Tj");
    expect(text).toContain("1 0 0 1 72.00 695.00 Tm");
    expect(text).toContain("(This is a selectable paragraph.) Tj");
  });

  it("generates /F_Uni with ToUnicode CMap for extended Unicode and CJK text", async () => {
    const textSpans: PdfTextSpan[] = [
      {
        text: "中文文档测试",
        x: 50,
        y: 500,
        width: 120,
        height: 16,
        fontSize: 14,
      },
      {
        text: "Café résumé — 100% paritas",
        x: 50,
        y: 475,
        width: 160,
        height: 14,
        fontSize: 12,
      },
    ];

    const shot: PdfPageShot = {
      width: 800,
      height: 600,
      jpeg: DUMMY_JPEG,
      textSpans,
    };

    const blob = await pagesToPdf([shot]);
    const text = await blob.text();

    expect(text).toContain("/Subtype /Type0");
    expect(text).toContain("/ToUnicode");
    expect(text).toContain("/CIDInit /ProcSet findresource begin");
    expect(text).toContain("<0000> <FFFF> <0000>");

    // Chinese hex encoded CIDs
    const hexCjk = encodeHexUtf16("中文文档测试");
    expect(text).toContain(`<${hexCjk}> Tj`);
  });

  it("generates PDF Link Annotations for hyperlinks and internal bookmarks", async () => {
    const links: PdfLinkAnnotation[] = [
      {
        rect: [72, 700, 180, 715],
        url: "https://example.com/docs",
        title: "Documentation Link",
      },
      {
        rect: [72, 650, 150, 665],
        url: "#section-toc",
      },
    ];

    const shot: PdfPageShot = {
      width: 800,
      height: 600,
      jpeg: DUMMY_JPEG,
      links,
    };

    const blob = await pagesToPdf([shot]);
    const text = await blob.text();

    expect(text).toContain("/Type /Annot");
    expect(text).toContain("/Subtype /Link");
    expect(text).toContain("/Rect [ 72.00 700.00 180.00 715.00 ]");
    expect(text).toContain("/URI (https://example.com/docs)");
    expect(text).toContain("/Dest (section-toc)");
  });

  it("generates Tagged PDF /MarkInfo and /StructTreeRoot for accessibility", async () => {
    const textSpans: PdfTextSpan[] = [
      {
        text: "Accessible Heading",
        x: 72,
        y: 720,
        width: 140,
        height: 18,
        fontSize: 16,
        tag: "H1",
      },
    ];

    const shot: PdfPageShot = {
      width: 800,
      height: 600,
      jpeg: DUMMY_JPEG,
      textSpans,
    };

    const blob = await pagesToPdf([shot], { tagged: true });
    const text = await blob.text();

    expect(text).toContain("/MarkInfo << /Marked true >>");
    expect(text).toContain("/StructTreeRoot");
    expect(text).toContain("/Type /StructElem");
    expect(text).toContain("/P << /MCID 0 >> BDC");
    expect(text).toContain("EMC");
  });
});

describe("extractPdfPageLayers", () => {
  it("extracts text spans and links from paragraphs, tables, and furniture", () => {
    const mockParagraph: LaidOutParagraph = {
      kind: "paragraph",
      heightPx: 40,
      beforePx: 0,
      afterPx: 0,
      lines: [
        {
          yPx: 0,
          heightPx: 20,
          naturalPx: 16,
          endInlineIndex: 0,
          items: [
            {
              kind: "text",
              inlineIndex: 0,
              text: "Hello Word",
              xPx: 0,
              widthPx: 80,
            },
          ],
        },
      ],
      inline: [
        {
          kind: "text",
          text: "Hello Word",
          style: { family: "Calibri", sizePx: 16 },
          link: { url: "https://word.com" },
        },
      ],
    };

    const mockTable: LaidOutTable = {
      kind: "table",
      widthPx: 300,
      columnWidthsPx: [150, 150],
      heightPx: 50,
      rows: [
        {
          heightPx: 50,
          cells: [
            {
              colspan: 1,
              rowspan: 1,
              innerWidthPx: 140,
              insets: { top: 4, right: 4, bottom: 4, left: 4 },
              stack: [
                {
                  yPx: 0,
                  block: {
                    kind: "paragraph",
                    heightPx: 20,
                    beforePx: 0,
                    afterPx: 0,
                    lines: [
                      {
                        yPx: 0,
                        heightPx: 20,
                        naturalPx: 16,
                        endInlineIndex: 0,
                        items: [
                          {
                            kind: "text",
                            inlineIndex: 0,
                            text: "Cell content",
                            xPx: 0,
                            widthPx: 70,
                          },
                        ],
                      },
                    ],
                    inline: [
                      {
                        kind: "text",
                        text: "Cell content",
                        style: { family: "Arial", sizePx: 14 },
                      },
                    ],
                  },
                },
              ],
            },
          ],
        },
      ],
    };

    const mockPage: FlowPage = {
      items: [
        { yPx: 0, block: mockParagraph },
        { yPx: 50, block: mockTable },
      ],
    };

    const mockSection: CanvasStageSection = {
      flow: {
        pageWidthPx: 816,
        pageHeightPx: 1056,
        contentLeftPx: 96,
        contentTopPx: 96,
        contentWidthPx: 624,
        contentHeightPx: 864,
      },
    };

    const layers = extractPdfPageLayers([mockPage], [mockSection], [0]);
    expect(layers).toHaveLength(1);
    const { textSpans, links } = layers[0]!;

    expect(textSpans).toHaveLength(2);
    expect(textSpans[0]!.text).toBe("Hello Word");
    expect(textSpans[1]!.text).toBe("Cell content");

    expect(links).toHaveLength(1);
    expect(links[0]!.url).toBe("https://word.com");
  });
});

describe("pdftotext and pdfinfo integration verification", () => {
  it("allows pdftotext to extract Latin and Unicode text with 100% accuracy", async () => {
    const textSpans: PdfTextSpan[] = [
      {
        text: "Welcome to Docen Word Parity",
        x: 72,
        y: 750,
        width: 220,
        height: 18,
        fontSize: 16,
        bold: true,
      },
      {
        text: "Testing searchable PDF text layer with CJK 中文测试 and accents: café naïve.",
        x: 72,
        y: 720,
        width: 380,
        height: 14,
        fontSize: 12,
      },
    ];

    const shot: PdfPageShot = {
      width: 816,
      height: 1056,
      jpeg: DUMMY_JPEG,
      textSpans,
    };

    const blob = await pagesToPdf([shot], {
      metadata: {
        title: "Integration Test PDF",
        author: "Docen Team",
      },
      tagged: true,
    });

    const arrayBuf = await blob.arrayBuffer();
    const buffer = Buffer.from(arrayBuf);

    const tmpPdf = path.join(os.tmpdir(), `docen-test-${Date.now()}.pdf`);
    fs.writeFileSync(tmpPdf, buffer);

    try {
      // 1. Verify text extraction via pdftotext
      const stdout = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(stdout).toContain("Welcome to Docen Word Parity");
      expect(stdout).toContain(
        "Testing searchable PDF text layer with CJK 中文测试 and accents: café naïve.",
      );

      // 2. Verify metadata and tagged info via pdfinfo
      const infoOut = execFileSync("pdfinfo", [tmpPdf], { encoding: "utf-8" });
      expect(infoOut).toMatch(/Title:\s+Integration Test PDF/);
      expect(infoOut).toMatch(/Author:\s+Docen Team/);
      expect(infoOut).toMatch(/Creator:\s+Docen Word Processor/);
      expect(infoOut).toMatch(/Producer:\s+Docen PDF Engine/);
      expect(infoOut).toMatch(/Tagged:\s+yes/);
      expect(infoOut).toMatch(/Pages:\s+1/);
    } finally {
      if (fs.existsSync(tmpPdf)) {
        if (fs.existsSync(tmpPdf)) {
          fs.unlinkSync(tmpPdf);
        }
      }
    }
  });

  it("embeds subset font stream and custom ToUnicode CMap when embeddedFonts option is provided", async () => {
    const dummyFontData = new Uint8Array([0x00, 0x01, 0x00, 0x00, 0x00, 0x04]);
    const toUnicodeMap = new Map<number, number>([
      [1, 0x0041], // 'A'
      [2, 0x0042], // 'B'
    ]);
    const shot: PdfPageShot = {
      width: 800,
      height: 600,
      jpeg: DUMMY_JPEG,
    };
    const blob = await pagesToPdf([shot], {
      embeddedFonts: [
        {
          fontName: "CustomSubsetFont",
          fontData: dummyFontData,
          toUnicodeMap,
        },
      ],
    });
    const text = await blob.text();
    expect(text).toContain("/FontFile2");
    expect(text).toContain("/FontName /CustomSubsetFont");
    expect(text).toContain("/CMapName /Custom-ToUnicode def");
    expect(text).toContain("<0001> <0041>");
  });
});
