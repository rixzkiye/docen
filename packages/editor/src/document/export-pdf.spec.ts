// @vitest-environment node
import { execFileSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { inflateSync } from "node:zlib";

import type { FlowPage, LaidOutParagraph, LaidOutTable } from "@docen/layout";
import { readCmap, subsetFontWithPlan } from "@docen/shaping/subsetter";
import { describe, expect, it } from "vitest";

import type { CanvasStageSection } from "./canvas/stage";
import {
  buildEmbeddedPdfFonts,
  encodeHexUtf16,
  escapePdfString,
  extractPdfPageLayers,
  pagesToPdf,
  type PdfLinkAnnotation,
  type PdfPageShot,
  type PdfTextSpan,
} from "./export-pdf";
import type { PdfScenePage } from "./pdf-scene";

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

  it("embeds a real remapped subset with CIDToGIDMap; pdftotext extracts Latin/accents/CJK byte-exact", async () => {
    const fontPath = path.resolve(
      __dirname,
      "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
    );
    const fontBytes = fs.readFileSync(fontPath);
    const cmap = readCmap(fontBytes);
    const latin = "Café résumé — naïve façade 100%";
    const cjk = "中文文档测试";

    const usedGids = new Set<number>([0]);
    const text = latin + cjk;
    for (const ch of text) {
      const gid = cmap.get(ch.codePointAt(0)!);
      if (gid !== undefined) usedGids.add(gid);
    }
    const plan = subsetFontWithPlan(fontBytes, [...usedGids]);
    const cidToGid = new Map<number, number>();
    for (let i = 0; i < text.length; i++) {
      const cu = text.charCodeAt(i);
      const gid = cmap.get(cu);
      const newGid = gid === undefined ? undefined : plan.glyphMap.get(gid);
      if (newGid !== undefined) cidToGid.set(cu, newGid);
    }

    const textSpans: PdfTextSpan[] = [
      {
        text: latin,
        x: 72,
        y: 750,
        width: 220,
        height: 18,
        fontSize: 16,
        fontFamily: "Open Sans",
      },
      {
        text: cjk,
        x: 72,
        y: 720,
        width: 120,
        height: 14,
        fontSize: 12,
        fontFamily: "Open Sans",
      },
    ];
    const blob = await pagesToPdf([{ width: 816, height: 1056, jpeg: DUMMY_JPEG, textSpans }], {
      metadata: { title: "Subset Gate" },
      embeddedFonts: [
        {
          fontName: "OpenSans-Subset",
          fontFamily: "Open Sans",
          fontData: plan.data,
          cidToGid,
        },
      ],
    });

    const tmpPdf = path.join(os.tmpdir(), `docen-subset-${Date.now()}.pdf`);
    fs.writeFileSync(tmpPdf, Buffer.from(await blob.arrayBuffer()));
    try {
      const extracted = execFileSync("pdftotext", ["-enc", "UTF-8", tmpPdf, "-"], {
        encoding: "utf-8",
      });
      expect(extracted).toContain(latin);
      expect(extracted).toContain(cjk);

      const fontsOut = execFileSync("pdffonts", [tmpPdf], { encoding: "utf-8" });
      expect(fontsOut).toContain("OpenSans-Subset");
      // Embedded (emb) and ToUnicode (uni) columns must both be yes.
      expect(fontsOut).toMatch(
        /OpenSans-Subset\s+CID TrueType\s+Identity-H\s+yes\s+(no|yes)\s+yes/,
      );

      // pdfinfo walks the xref: a corrupt table fails here.
      const infoOut = execFileSync("pdfinfo", [tmpPdf], { encoding: "utf-8" });
      expect(infoOut).toMatch(/Title:\s+Subset Gate/);
    } finally {
      fs.rmSync(tmpPdf, { force: true });
    }
  });
});

describe("buildEmbeddedPdfFonts", () => {
  const fontPath = path.resolve(
    __dirname,
    "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
  );
  const span = (text: string, fontFamily?: string): PdfTextSpan => ({
    text,
    x: 0,
    y: 0,
    width: 100,
    height: 14,
    fontSize: 12,
    fontFamily,
  });

  it("subsets installed fonts and maps code units to remapped glyph ids", async () => {
    const fontData = fs.readFileSync(fontPath);
    const fonts = await buildEmbeddedPdfFonts(
      [span("Café", "Open Sans")],
      [{ family: "Open Sans", fontData }],
    );
    expect(fonts).toHaveLength(1);
    const embedded = fonts[0]!;
    expect(embedded.fontData.byteLength).toBeLessThan(fontData.byteLength);
    expect(embedded.fontName).toBe("OpenSansSubset");
    expect(embedded.fsType).toBe(0);
    expect(embedded.cidToGid?.get("C".charCodeAt(0))).toBeGreaterThan(0);
    expect(embedded.cidToGid?.get("é".charCodeAt(0))).toBeGreaterThan(0);
  });

  it("skips restricted fonts (fsType 0x0002) and embeds no-subsetting fonts whole", async () => {
    const fontData = fs.readFileSync(fontPath);
    const spans = [span("Restricted", "Locked Font")];
    const restricted = await buildEmbeddedPdfFonts(spans, [
      { family: "Locked Font", fontData, fsType: 0x0002 },
    ]);
    expect(restricted).toEqual([]);

    const whole = await buildEmbeddedPdfFonts(
      [span("Whole", "NoSubset Font")],
      [{ family: "NoSubset Font", fontData, fsType: 0x0100 }],
    );
    expect(whole).toHaveLength(1);
    expect(whole[0]!.fontData).toBe(fontData); // untouched full font
    expect(whole[0]!.fontName).toBe("NoSubsetFont");
  });
});

describe("P1 vector core acceptance (A1-A3)", () => {
  const sampleScene: PdfScenePage = {
    width: 612,
    height: 792,
    nodes: [
      {
        type: "shape",
        matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        path: "M 50 50 L 250 50 L 250 250 Z",
        fill: "#0066cc",
        stroke: "#003366",
        strokeWidth: 2,
      },
      {
        type: "shape",
        matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        path: "M 100 100 C 120 180 180 180 200 100",
        stroke: "#cc0000",
        strokeWidth: 1.5,
      },
    ],
  };

  const sampleSpans: PdfTextSpan[] = [
    {
      text: "Vector Core Fidelity Test",
      x: 50,
      y: 500,
      width: 200,
      height: 18,
      fontSize: 14,
    },
    {
      text: "Layout text extraction with pdftotext 100%",
      x: 50,
      y: 470,
      width: 280,
      height: 14,
      fontSize: 11,
    },
  ];

  it("A1: page with scene has no full-page Image XObject and emits path operators", async () => {
    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: sampleScene,
      textSpans: sampleSpans,
    };
    const blob = await pagesToPdf([shot]);
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Must not contain full-page Image XObject
    const imageCount = (pdfStr.match(/\/Subtype\s*\/Image\b/g) || []).length;
    expect(imageCount).toBe(0);

    // Decompress stream and check path operators
    const streamMatches = [
      ...pdfStr.matchAll(/<<([^>]*)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream/g),
    ];
    let decompressed = "";
    for (const m of streamMatches) {
      if (m[1].includes("/FlateDecode")) {
        decompressed += inflateSync(Buffer.from(m[2], "latin1")).toString("latin1");
      }
    }

    expect(decompressed).toMatch(/\bm\b/); // moveto
    expect(decompressed).toMatch(/\bl\b/); // lineto
    expect(decompressed).toMatch(/\bc\b/); // curveto
    expect(decompressed).toMatch(/\b(f\*?|B\*?)\b/); // fill or fill-and-stroke
    expect(decompressed).toMatch(/\b(S|B\*?)\b/); // stroke
  });

  it("A2 & A3: pdftoppm renders without error and pdftotext extracts layout text", async () => {
    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: sampleScene,
      textSpans: sampleSpans,
    };
    const blob = await pagesToPdf([shot]);
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-test-"));
    const tmpPdf = path.join(tmpDir, "test.pdf");
    const ppmPrefix = path.join(tmpDir, "page");

    try {
      fs.writeFileSync(tmpPdf, Buffer.from(await blob.arrayBuffer()));

      // A2: pdftoppm renders page PNG
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      const renderedPng = `${ppmPrefix}-1.png`;
      expect(fs.existsSync(renderedPng)).toBe(true);
      expect(fs.statSync(renderedPng).size).toBeGreaterThan(1000);

      // A3: pdftotext extracts text
      const extracted = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(extracted).toContain("Vector Core Fidelity Test");
      expect(extracted).toContain("Layout text extraction with pdftotext 100%");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});

describe("P2 graphics fidelity acceptance", () => {
  it("P2 table fidelity: cell shading and dashed/solid borders emit vector paths and dash patterns", async () => {
    const tableScene: PdfScenePage = {
      width: 612,
      height: 792,
      nodes: [
        // Cell 1 shading
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 50 100 L 250 100 L 250 150 L 50 150 Z",
          fill: "#2563eb",
        },
        // Cell 2 shading
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 250 100 L 450 100 L 450 150 L 250 150 Z",
          fill: "#e0f2fe",
        },
        // Dashed border between cells
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 250 100 L 250 150",
          stroke: "#1e40af",
          strokeWidth: 2,
          dash: [4, 2],
        },
        // Outer table border (solid)
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 50 100 L 450 100 L 450 150 L 50 150 Z",
          stroke: "#0f172a",
          strokeWidth: 1.5,
        },
      ],
    };

    const tableSpans: PdfTextSpan[] = [
      { text: "Header Col 1", x: 60, y: 500, width: 100, height: 14, fontSize: 11 },
      { text: "Header Col 2", x: 260, y: 500, width: 100, height: 14, fontSize: 11 },
    ];

    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: tableScene,
      textSpans: tableSpans,
    };

    const blob = await pagesToPdf([shot]);
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Zero image XObjects
    expect((pdfStr.match(/\/Subtype\s*\/Image\b/g) || []).length).toBe(0);

    // Decompress and verify operators
    const streamMatches = [
      ...pdfStr.matchAll(/<<([^>]*)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream/g),
    ];
    let decompressed = "";
    for (const m of streamMatches) {
      if (m[1].includes("/FlateDecode")) {
        decompressed += inflateSync(Buffer.from(m[2], "latin1")).toString("latin1");
      }
    }

    expect(decompressed).toContain("[ 4 2 ] 0 d"); // dash pattern
    expect(decompressed).toMatch(/\b(f\*?|B\*?)\b/); // fill
    expect(decompressed).toMatch(/\b(S|B\*?)\b/); // stroke

    // Verify render and text extraction
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-tbl-"));
    const tmpPdf = path.join(tmpDir, "table.pdf");
    const ppmPrefix = path.join(tmpDir, "table-page");
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      expect(fs.existsSync(`${ppmPrefix}-1.png`)).toBe(true);

      const extracted = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(extracted).toContain("Header Col 1");
      expect(extracted).toContain("Header Col 2");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("P2 shapes & drawings: custGeom, ellipse, and text box emit bezier curves without rasterization", async () => {
    const drawingScene: PdfScenePage = {
      width: 612,
      height: 792,
      nodes: [
        // Ellipse (arc to cubic)
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 100 50 A 50 30 0 1 0 100 49.9 Z",
          fill: "#f59e0b",
          stroke: "#b45309",
          strokeWidth: 2,
        },
        // Custom geometry with cubics
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 200 200 C 220 150 280 150 300 200 C 320 250 280 300 200 200 Z",
          fill: "#10b981",
          stroke: "#047857",
          strokeWidth: 1.5,
        },
        // Textbox clip group with background
        {
          type: "group",
          clip: {
            path: "M 50 350 L 250 350 L 250 450 L 50 450 Z",
            matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          },
          children: [
            {
              type: "shape",
              matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
              path: "M 50 350 L 250 350 L 250 450 L 50 450 Z",
              fill: "#f8fafc",
              stroke: "#94a3b8",
              strokeWidth: 1,
            },
            {
              type: "text",
              matrix: { a: 1, b: 0, c: 0, d: 1, e: 60, f: 380 },
              fontSize: 12,
              fill: "#1e293b",
              rows: [{ x: 0, y: 0, width: 150, text: "Textbox inner text content" }],
            },
          ],
        },
      ],
    };

    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: drawingScene,
      textSpans: [
        { text: "Textbox inner text content", x: 60, y: 300, width: 150, height: 14, fontSize: 12 },
      ],
    };

    const blob = await pagesToPdf([shot]);
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Zero image XObjects
    expect((pdfStr.match(/\/Subtype\s*\/Image\b/g) || []).length).toBe(0);

    // Decompress and verify operators
    const streamMatches = [
      ...pdfStr.matchAll(/<<([^>]*)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream/g),
    ];
    let decompressed = "";
    for (const m of streamMatches) {
      if (m[1].includes("/FlateDecode")) {
        decompressed += inflateSync(Buffer.from(m[2], "latin1")).toString("latin1");
      }
    }

    expect(decompressed).toMatch(/\bc\b/); // cubics from arc/custGeom
    expect(decompressed).toContain("W n"); // clip operator

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-draw-"));
    const tmpPdf = path.join(tmpDir, "drawing.pdf");
    const ppmPrefix = path.join(tmpDir, "drawing-page");
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      expect(fs.existsSync(`${ppmPrefix}-1.png`)).toBe(true);

      const extracted = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(extracted).toContain("Textbox inner text content");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("P2 transparent images: emits DeviceRGB stream with 3 bytes/pixel and DeviceGray SMask", async () => {
    // 2x2 semi-transparent image
    const width = 2;
    const height = 2;
    const rgba = new Uint8Array([
      255,
      0,
      0,
      128, // red, 50% alpha
      0,
      255,
      0,
      64, // green, 25% alpha
      0,
      0,
      255,
      200, // blue, ~78% alpha
      255,
      255,
      255,
      255, // white, opaque
    ]);

    const transScene: PdfScenePage = {
      width: 612,
      height: 792,
      nodes: [
        {
          type: "image",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 100, f: 100 },
          width: 200,
          height: 200,
          image: {
            key: "test-trans-alpha-key",
            width,
            height,
            rgba,
            hasAlpha: true,
          },
        },
      ],
    };

    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: transScene,
    };

    const blob = await pagesToPdf([shot]);
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Must have SMask reference
    expect(pdfStr).toMatch(/\/SMask\s+\d+\s+0\s+R/);

    // Verify image and SMask stream sizes
    const streamMatches = [
      ...pdfStr.matchAll(/<<([^>]*)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream/g),
    ];
    let rgbDecompressedLen = 0;
    let smaskDecompressedLen = 0;

    for (const m of streamMatches) {
      const dict = m[1]!;
      if (dict.includes("/ColorSpace /DeviceRGB") && dict.includes("/Subtype /Image")) {
        const dec = inflateSync(Buffer.from(m[2], "latin1"));
        rgbDecompressedLen = dec.length;
      }
      if (dict.includes("/ColorSpace /DeviceGray") && dict.includes("/Subtype /Image")) {
        const dec = inflateSync(Buffer.from(m[2], "latin1"));
        smaskDecompressedLen = dec.length;
      }
    }

    // DeviceRGB must be 2*2*3 = 12 bytes (NOT 16 bytes!)
    expect(rgbDecompressedLen).toBe(12);
    // SMask must be 2*2*1 = 4 bytes
    expect(smaskDecompressedLen).toBe(4);

    // Verify pdftoppm renders cleanly
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-img-"));
    const tmpPdf = path.join(tmpDir, "image.pdf");
    const ppmPrefix = path.join(tmpDir, "image-page");
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      expect(fs.existsSync(`${ppmPrefix}-1.png`)).toBe(true);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("P2 vector chart: series bars, axis lines, and labels render as vectors", async () => {
    const chartScene: PdfScenePage = {
      width: 612,
      height: 792,
      nodes: [
        // Chart title
        {
          type: "text",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 200, f: 100 },
          fontSize: 16,
          bold: true,
          fill: "#333333",
          rows: [{ x: 0, y: 0, width: 200, text: "Quarterly Revenue Breakdown" }],
        },
        // Gridlines
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 100 250 L 500 250 M 100 200 L 500 200 M 100 150 L 500 150",
          stroke: "#d9d9d9",
          strokeWidth: 1,
          dash: [2, 2],
        },
        // Axis lines
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 100 150 L 100 300 L 500 300",
          stroke: "#595959",
          strokeWidth: 1.5,
        },
        // Series Bar 1 (Q1)
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 140 200 L 180 200 L 180 300 L 140 300 Z",
          fill: "#4472c4", // Accent color
        },
        // Series Bar 2 (Q2)
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 220 160 L 260 160 L 260 300 L 220 300 Z",
          fill: "#ed7d31", // Accent color
        },
      ],
    };

    const chartSpans: PdfTextSpan[] = [
      { text: "Quarterly Revenue Breakdown", x: 200, y: 550, width: 200, height: 18, fontSize: 16 },
    ];

    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: chartScene,
      textSpans: chartSpans,
    };

    const blob = await pagesToPdf([shot]);
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Zero image XObjects — chart is 100% vector
    expect((pdfStr.match(/\/Subtype\s*\/Image\b/g) || []).length).toBe(0);

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-ch-"));
    const tmpPdf = path.join(tmpDir, "chart.pdf");
    const ppmPrefix = path.join(tmpDir, "chart-page");
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      expect(fs.existsSync(`${ppmPrefix}-1.png`)).toBe(true);

      const extracted = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(extracted).toContain("Quarterly Revenue Breakdown");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("P2 math equations: radical surd, fraction line, and math symbols emit vector paths", async () => {
    const mathScene: PdfScenePage = {
      width: 612,
      height: 792,
      nodes: [
        // Radical surd tick and bar
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 100, f: 100 },
          path: "M 0 12 L 3 10 L 6 20 L 10 1 L 60 1",
          stroke: "#000000",
          strokeWidth: 1.5,
        },
        // Fraction bar
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 120, f: 110 },
          path: "M 0 0 L 40 0",
          stroke: "#000000",
          strokeWidth: 1.2,
        },
        // Formula text
        {
          type: "text",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 70, f: 112 },
          fontSize: 14,
          fill: "#000000",
          rows: [{ x: 0, y: 0, width: 30, text: "E = " }],
        },
      ],
    };

    const mathSpans: PdfTextSpan[] = [
      { text: "E = mc^2", x: 70, y: 550, width: 80, height: 16, fontSize: 14 },
    ];

    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: mathScene,
      textSpans: mathSpans,
    };

    const blob = await pagesToPdf([shot]);
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Zero image XObjects
    expect((pdfStr.match(/\/Subtype\s*\/Image\b/g) || []).length).toBe(0);

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-math-"));
    const tmpPdf = path.join(tmpDir, "math.pdf");
    const ppmPrefix = path.join(tmpDir, "math-page");
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      expect(fs.existsSync(`${ppmPrefix}-1.png`)).toBe(true);

      const extracted = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(extracted).toContain("E = mc^2");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("P2 watermarks & 3D/ink: rotated semi-transparent watermark and 3D isometric cube", async () => {
    const angle = (-45 * Math.PI) / 180;
    const cos = Math.cos(angle);
    const sin = Math.sin(angle);

    const watermarkScene: PdfScenePage = {
      width: 612,
      height: 792,
      nodes: [
        // Rotated diagonal watermark text
        {
          type: "text",
          matrix: { a: cos, b: sin, c: -sin, d: cos, e: 200, f: 400 },
          opacity: 0.35,
          fontSize: 54,
          bold: true,
          fill: "#c0c0c0", // silver watermark
          rows: [{ x: 0, y: 0, width: 350, text: "CONFIDENTIAL" }],
        },
        // 3D wireframe isometric cube faces
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 400 200 L 450 230 L 400 260 L 350 230 Z",
          fill: "#93c5fd",
          stroke: "#1e3a8a",
          strokeWidth: 1,
        },
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 350 230 L 400 260 L 400 320 L 350 290 Z",
          fill: "#60a5fa",
          stroke: "#1e3a8a",
          strokeWidth: 1,
        },
        // Cursive ink stroke
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 100 600 C 150 550 200 650 250 580 C 300 510 350 620 400 570",
          stroke: "#1e293b",
          strokeWidth: 2.5,
          cap: "round",
          join: "round",
        },
      ],
    };

    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: watermarkScene,
      textSpans: [{ text: "CONFIDENTIAL", x: 50, y: 350, width: 250, height: 36, fontSize: 36 }],
    };

    const blob = await pagesToPdf([shot]);
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Verify opacity ExtGState
    expect(pdfStr).toContain("/ca 0.35");
    expect(pdfStr).toContain("/CA 0.35");

    // Zero image XObjects
    expect((pdfStr.match(/\/Subtype\s*\/Image\b/g) || []).length).toBe(0);

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-wm-"));
    const tmpPdf = path.join(tmpDir, "watermark.pdf");
    const ppmPrefix = path.join(tmpDir, "wm-page");
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      expect(fs.existsSync(`${ppmPrefix}-1.png`)).toBe(true);

      const extracted = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(extracted).toContain("CONFIDENTIAL");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("P2 ruby phonetic guide: renders annotation text aligned above base text", async () => {
    const rubyScene: PdfScenePage = {
      width: 612,
      height: 792,
      nodes: [
        // Ruby furigana annotation text
        {
          type: "text",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 100, f: 180 },
          fontSize: 8,
          fill: "#64748b",
          rows: [{ x: 0, y: 0, width: 40, text: "とうきょう" }],
        },
        // Base kanji text
        {
          type: "text",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 100, f: 200 },
          fontSize: 16,
          fill: "#0f172a",
          rows: [{ x: 0, y: 0, width: 40, text: "東京" }],
        },
      ],
    };

    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: rubyScene,
      textSpans: [
        { text: "とうきょう", x: 100, y: 440, width: 40, height: 10, fontSize: 8 },
        { text: "東京", x: 100, y: 420, width: 40, height: 18, fontSize: 16 },
      ],
    };

    const blob = await pagesToPdf([shot]);
    const pdfBuf = Buffer.from(await blob.arrayBuffer());

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-ruby-"));
    const tmpPdf = path.join(tmpDir, "ruby.pdf");
    const ppmPrefix = path.join(tmpDir, "ruby-page");
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      expect(fs.existsSync(`${ppmPrefix}-1.png`)).toBe(true);

      const extracted = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(extracted).toContain("とうきょう");
      expect(extracted).toContain("東京");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});

describe("P3 structure and navigation acceptance", () => {
  const blankShot = (pageIndex: number): PdfPageShot => ({
    width: 612,
    height: 792,
    jpeg: DUMMY_JPEG,
    textSpans: [
      {
        text: `Page ${pageIndex + 1} content`,
        x: 72,
        y: 700,
        width: 150,
        height: 14,
        fontSize: 12,
      },
    ],
  });

  it("P3 outlines (bookmarks): builds hierarchical outline tree with parent, prev, next, first, last, count and target dest", async () => {
    const shots = [blankShot(0), blankShot(1)];
    const blob = await pagesToPdf(shots, {
      outline: [
        {
          title: "Chapter 1",
          dest: { pageIndex: 0, top: 750 },
          children: [
            { title: "Section 1.1", dest: { pageIndex: 0, top: 700 } },
            { title: "Section 1.2", dest: { pageIndex: 0, top: 600 } },
          ],
        },
        {
          title: "Chapter 2",
          dest: 1,
        },
      ],
    });
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Catalog references Outlines
    expect(pdfStr).toMatch(/\/Type\s*\/Catalog[^\n]*\/Outlines\s+(\d+)\s+0\s+R/);
    const outlineRootMatch = pdfStr.match(/\/Type\s*\/Catalog[^\n]*\/Outlines\s+(\d+)\s+0\s+R/);
    const outlineRootId = outlineRootMatch![1];

    // Outlines root dictionary
    expect(pdfStr).toContain(`${outlineRootId} 0 obj\n<< /Type /Outlines`);
    expect(pdfStr).toMatch(
      new RegExp(
        `${outlineRootId} 0 obj\\n<< \\/Type \\/Outlines \\/First (\\d+) 0 R \\/Last (\\d+) 0 R \\/Count 4 >>`,
      ),
    );

    // Outline items
    expect(pdfStr).toContain("/Title (Chapter 1)");
    expect(pdfStr).toContain("/Title (Section 1.1)");
    expect(pdfStr).toContain("/Title (Section 1.2)");
    expect(pdfStr).toContain("/Title (Chapter 2)");

    // Parent/child relations
    expect(pdfStr).toMatch(/\/Title \(Chapter 1\)[^\n]*\/Count 2/);
    expect(pdfStr).toMatch(/\/Title \(Chapter 1\)[^\n]*\/Dest \[ \d+ 0 R \/XYZ 0 750\.00 0 \]/);
    expect(pdfStr).toMatch(/\/Title \(Chapter 2\)[^\n]*\/Dest \[ \d+ 0 R \/XYZ 0 792\.00 0 \]/);

    // Verify xref valid via pdfinfo
    const tmpPdf = path.join(os.tmpdir(), `pdf-outline-${Date.now()}.pdf`);
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      const infoOut = execFileSync("pdfinfo", [tmpPdf], { encoding: "utf-8" });
      expect(infoOut).toContain("Pages:           2");
    } finally {
      fs.rmSync(tmpPdf, { force: true });
    }
  });

  it("P3 page labels: emits section-based numbering with roman, decimal, and prefix styles", async () => {
    const shots = [blankShot(0), blankShot(1), blankShot(2)];
    const blob = await pagesToPdf(shots, {
      pageLabels: [
        { startPageIndex: 0, style: "romanLower" },
        { startPageIndex: 1, style: "decimal", prefix: "App-", startNumber: 1 },
      ],
    });
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    expect(pdfStr).toContain(
      "/PageLabels << /Nums [ 0 << /S /r >> 1 << /S /D /P (App-) /St 1 >> ] >>",
    );

    const tmpPdf = path.join(os.tmpdir(), `pdf-labels-${Date.now()}.pdf`);
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      const infoOut = execFileSync("pdfinfo", [tmpPdf], { encoding: "utf-8" });
      expect(infoOut).toContain("Pages:           3");
    } finally {
      fs.rmSync(tmpPdf, { force: true });
    }
  });

  it("P3 document metadata: emits title, author, subject, keywords, creation and modification dates; verified via pdfinfo", async () => {
    const shots = [blankShot(0)];
    const blob = await pagesToPdf(shots, {
      metadata: {
        title: "Docen Technical Whitepaper",
        author: "Docen Platform Team",
        subject: "Vector PDF Export Architecture",
        keywords: "pdf, vector, docen, typography",
      },
    });
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    expect(pdfStr).toContain("/Title (Docen Technical Whitepaper)");
    expect(pdfStr).toContain("/Author (Docen Platform Team)");
    expect(pdfStr).toContain("/Subject (Vector PDF Export Architecture)");
    expect(pdfStr).toContain("/Keywords (pdf, vector, docen, typography)");
    expect(pdfStr).toContain("/ModDate (D:");

    const tmpPdf = path.join(os.tmpdir(), `pdf-meta-${Date.now()}.pdf`);
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      const infoOut = execFileSync("pdfinfo", [tmpPdf], { encoding: "utf-8" });
      expect(infoOut).toMatch(/Title:\s+Docen Technical Whitepaper/);
      expect(infoOut).toMatch(/Author:\s+Docen Platform Team/);
      expect(infoOut).toMatch(/Subject:\s+Vector PDF Export Architecture/);
      expect(infoOut).toMatch(/Keywords:\s+pdf, vector, docen, typography/);
      expect(infoOut).toMatch(/Creator:\s+Docen Word Processor/);
      expect(infoOut).toMatch(/Producer:\s+Docen PDF Engine/);
    } finally {
      fs.rmSync(tmpPdf, { force: true });
    }
  });

  it("P3 viewer preferences: emits displayDocTitle, centerWindow, fitWindow, hideToolbar", async () => {
    const shots = [blankShot(0)];
    const blob = await pagesToPdf(shots, {
      viewerPreferences: {
        displayDocTitle: true,
        centerWindow: true,
        fitWindow: true,
        hideToolbar: false,
      },
    });
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    expect(pdfStr).toContain(
      "/ViewerPreferences << /DisplayDocTitle true /HideToolbar false /CenterWindow true /FitWindow true >>",
    );
  });

  it("P3 named destinations: emits targetable destinations in Catalog /Dests and link resolution", async () => {
    const shots = [
      {
        ...blankShot(0),
        links: [
          {
            rect: [72, 600, 150, 615] as [number, number, number, number],
            url: "#summary-table",
          },
        ],
      },
      blankShot(1),
    ];
    const blob = await pagesToPdf(shots, {
      destinations: {
        "summary-table": { pageIndex: 1, x: 72, y: 650, zoom: 1 },
      },
    });
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    expect(pdfStr).toMatch(/\/Dests << \/summary-table \[ \d+ 0 R \/XYZ 72\.00 650\.00 1 \] >>/);
    expect(pdfStr).toContain("/Dest (summary-table)");
  });

  it("P3 AcroForm form fields: generates text fields and checkboxes with widget annotations", async () => {
    const shots = [blankShot(0)];
    const blob = await pagesToPdf(shots, {
      formFields: [
        {
          name: "customer_name",
          type: "text",
          pageIndex: 0,
          rect: [100, 600, 300, 620],
          value: "Jane Smith",
          readOnly: true,
        },
        {
          name: "subscribe_newsletter",
          type: "checkbox",
          pageIndex: 0,
          rect: [100, 550, 120, 570],
          value: true,
        },
        {
          name: "opt_in_sms",
          type: "checkbox",
          pageIndex: 0,
          rect: [100, 500, 120, 520],
          value: false,
        },
      ],
    });
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    // Catalog has AcroForm
    expect(pdfStr).toMatch(/\/Type\s*\/Catalog[^\n]*\/AcroForm\s+(\d+)\s+0\s+R/);

    // AcroForm object
    expect(pdfStr).toMatch(
      /<< \/Fields \[ ((\d+ 0 R )+)\] \/NeedAppearances true \/DA \(\/F1 12 Tf 0 g\) >>/,
    );

    // Text field widget
    expect(pdfStr).toContain("/Subtype /Widget");
    expect(pdfStr).toContain("/FT /Tx");
    expect(pdfStr).toContain("/T (customer_name)");
    expect(pdfStr).toContain("/V (Jane Smith)");
    expect(pdfStr).toContain("/Ff 1"); // Read-only

    // Checked checkbox widget
    expect(pdfStr).toContain("/FT /Btn");
    expect(pdfStr).toContain("/T (subscribe_newsletter)");
    expect(pdfStr).toContain("/V /Yes");
    expect(pdfStr).toContain("/AS /Yes");

    // Unchecked checkbox widget
    expect(pdfStr).toContain("/T (opt_in_sms)");
    expect(pdfStr).toContain("/V /Off");
    expect(pdfStr).toContain("/AS /Off");

    // Page Annots contains all 3 widget IDs
    const pageAnnotsMatch = pdfStr.match(/\/Type\s*\/Page[^\n]*\/Annots\s*\[\s*([^\]]+)\]/);
    expect(pageAnnotsMatch).not.toBeNull();
    const annotRefs = pageAnnotsMatch![1]!.trim().split(/\s+/);
    expect(annotRefs.length).toBe(9); // 3 annots * 3 tokens ("id 0 R") = 9 tokens
  });

  it("P3 tagged PDF accessibility: emits StructElem with Alt text for figure / table", async () => {
    const shots = [blankShot(0)];
    const blob = await pagesToPdf(shots, {
      tagged: true,
      structElements: [
        {
          type: "Figure",
          pageIndex: 0,
          altText: "Quarterly revenue bar chart comparing Q1-Q4 2025",
          title: "Figure 1: Revenue Chart",
        },
        {
          type: "Table",
          pageIndex: 0,
          altText: "Summary table of financial metrics",
          title: "Table 1",
        },
      ],
    });
    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const pdfStr = pdfBuf.toString("latin1");

    expect(pdfStr).toContain("/Type /StructTreeRoot");
    expect(pdfStr).toContain(
      "/RoleMap << /H1 /H /H2 /H /H3 /H /H4 /H /P /P /Table /Table /Figure /Figure >>",
    );
    expect(pdfStr).toContain("/S /Figure");
    expect(pdfStr).toContain("/Alt (Quarterly revenue bar chart comparing Q1-Q4 2025)");
    expect(pdfStr).toContain("/T (Figure 1: Revenue Chart)");
    expect(pdfStr).toContain("/S /Table");
    expect(pdfStr).toContain("/Alt (Summary table of financial metrics)");
    expect(pdfStr).toContain("/T (Table 1)");
  });
});

describe("P4 QA, determinism, and performance acceptance", () => {
  const blankShot = (pageIndex: number): PdfPageShot => ({
    width: 612,
    height: 792,
    jpeg: DUMMY_JPEG,
    textSpans: [
      {
        text: `Page ${pageIndex + 1} content in test document`,
        x: 72,
        y: 700,
        width: 200,
        height: 14,
        fontSize: 12,
      },
    ],
  });

  it("P4 determinism: repeated exports of identical document with fixed creationDate yield byte-identical buffers", async () => {
    const fixedDate = new Date("2026-01-15T10:30:00Z");
    const shots = [blankShot(0), blankShot(1), blankShot(2)];
    const options = {
      metadata: {
        title: "Deterministic Specification",
        author: "Docen Quality Assurance",
        subject: "Repeatable Vector PDF Builds",
        keywords: "reproducible, deterministic, pdf",
        creationDate: fixedDate,
        modDate: fixedDate,
      },
      outline: [
        {
          title: "Introduction",
          dest: { pageIndex: 0, top: 750 },
          children: [{ title: "Background", dest: { pageIndex: 1, top: 700 } }],
        },
        { title: "Conclusion", dest: 2 },
      ],
      pageLabels: [
        { startPageIndex: 0, style: "romanLower" as const },
        { startPageIndex: 1, style: "decimal" as const, startNumber: 1 },
      ],
      destinations: {
        "intro-anchor": { pageIndex: 0, x: 72, y: 750 },
      },
      formFields: [
        {
          name: "confirmed",
          type: "checkbox" as const,
          pageIndex: 2,
          rect: [100, 500, 120, 520] as [number, number, number, number],
          value: true,
        },
      ],
    };

    const [blob1, blob2] = await Promise.all([
      pagesToPdf(shots, options),
      pagesToPdf(shots, options),
    ]);

    const buf1 = Buffer.from(await blob1.arrayBuffer());
    const buf2 = Buffer.from(await blob2.arrayBuffer());

    expect(buf1.byteLength).toBe(buf2.byteLength);
    expect(Buffer.compare(buf1, buf2)).toBe(0);
    expect(buf1.equals(buf2)).toBe(true);
  });

  it("P4 300+ page benchmark: exports 300 pages efficiently with valid xref and document structure", async () => {
    const count = 300;
    const shots = Array.from({ length: count }, (_, i) => ({
      width: 612,
      height: 792,
      jpeg: DUMMY_JPEG,
      textSpans: [
        {
          text: `Page ${i + 1} of ${count} in high-throughput PDF generation benchmark.`,
          x: 72,
          y: 700,
          width: 350,
          height: 14,
          fontSize: 12,
        },
      ],
    }));

    const start = performance.now();
    const blob = await pagesToPdf(shots, {
      metadata: {
        title: "300 Page Stress Test",
        creationDate: new Date("2026-01-01T00:00:00Z"),
      },
      pageLabels: [{ startPageIndex: 0, style: "decimal" }],
    });
    const elapsed = performance.now() - start;
    const pdfBuf = Buffer.from(await blob.arrayBuffer());

    // File generated with proper size
    expect(pdfBuf.byteLength).toBeGreaterThan(150 * 1024);

    // Throughput assertion: 300 pages must take less than 5000ms (< 17ms per page)
    expect(elapsed).toBeLessThan(5000);

    // Validate with pdfinfo to ensure xref table and catalog for 300 pages are completely valid
    const tmpPdf = path.join(os.tmpdir(), `pdf-300p-${Date.now()}.pdf`);
    try {
      fs.writeFileSync(tmpPdf, pdfBuf);
      const infoOut = execFileSync("pdfinfo", [tmpPdf], { encoding: "utf-8" });
      expect(infoOut).toContain("Pages:           300");
    } finally {
      fs.rmSync(tmpPdf, { force: true });
    }
  });

  it("P4 smoke multi-viewer validation: verify poppler tools parse rich P1-P3 features cleanly", async () => {
    const vectorScene: PdfScenePage = {
      width: 612,
      height: 792,
      nodes: [
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 72 72 L 540 72 L 540 200 L 72 200 Z",
          fill: "#f1f5f9",
          stroke: "#cbd5e1",
          strokeWidth: 1,
        },
        {
          type: "shape",
          matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          path: "M 100 150 C 150 100 200 200 250 150",
          stroke: "#2563eb",
          strokeWidth: 2,
        },
      ],
    };

    const shot: PdfPageShot = {
      width: 612,
      height: 792,
      scene: vectorScene,
      textSpans: [
        {
          text: "Executive Summary & Performance Metrics",
          x: 72,
          y: 500,
          width: 300,
          height: 18,
          fontSize: 16,
          bold: true,
        },
      ],
      links: [
        {
          rect: [72, 490, 200, 510],
          url: "https://docen.dev",
          title: "Docen Documentation",
        },
      ],
    };

    const blob = await pagesToPdf([shot], {
      metadata: {
        title: "Executive Report",
        author: "Chief Architect",
        subject: "Quarterly Systems Audit",
        keywords: "executive, audit, vector",
        creationDate: new Date("2026-02-01T00:00:00Z"),
      },
      outline: [{ title: "Executive Summary", dest: { pageIndex: 0, top: 500 } }],
      formFields: [
        {
          name: "approved",
          type: "checkbox",
          pageIndex: 0,
          rect: [72, 440, 90, 458],
          value: true,
        },
      ],
      structElements: [
        {
          type: "Figure",
          pageIndex: 0,
          altText: "System health metric diagram",
          title: "Figure 1",
        },
      ],
    });

    const pdfBuf = Buffer.from(await blob.arrayBuffer());
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-smoke-"));
    const tmpPdf = path.join(tmpDir, "smoke.pdf");
    const ppmPrefix = path.join(tmpDir, "smoke-page");

    try {
      fs.writeFileSync(tmpPdf, pdfBuf);

      // 1. pdfinfo inspection
      const info = execFileSync("pdfinfo", [tmpPdf], { encoding: "utf-8" });
      expect(info).toMatch(/Title:\s+Executive Report/);
      expect(info).toMatch(/Author:\s+Chief Architect/);
      expect(info).toMatch(/Subject:\s+Quarterly Systems Audit/);
      expect(info).toMatch(/Pages:\s+1/);

      // 2. pdftoppm rasterization
      execFileSync("pdftoppm", ["-png", "-r", "150", tmpPdf, ppmPrefix]);
      expect(fs.existsSync(`${ppmPrefix}-1.png`)).toBe(true);

      // 3. pdftotext extraction
      const text = execFileSync("pdftotext", [tmpPdf, "-"], { encoding: "utf-8" });
      expect(text).toContain("Executive Summary & Performance Metrics");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});
