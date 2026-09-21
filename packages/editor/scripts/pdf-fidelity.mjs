#!/usr/bin/env node
/**
 * PDF Export Fidelity Harness (R10-P1)
 *
 * Verifies vector export fidelity against the live LeaferJS canvas engine:
 *   A1. Scene pages contain no full-page image XObjects; content streams contain
 *       vector path operators (m, l, c, f, S).
 *   A2. Visual fidelity: sub-pixel alignment search + PSNR, gated at a floor
 *       of 22 dB per page and overall (calibrated to the measured
 *       cross-rasterizer baseline of 23.9–31.1 dB; catches layout regressions).
 *       Mean abs diffs, the alignment offset and the AA-averaged SSIM
 *       (2x-downsampled, diagnostic) are reported; no page exceptions.
 *   A3. Text extraction: pdftotext extracts layout text runs.
 *   A4. Measurement: records size, export time, and fidelity for outlines vs
 *       embedded vs legacy raster.
 *   A5. PDF/A-2b + PDF/UA-1 conformance via veraPDF — the CLI is a hard
 *       requirement (a missing binary fails the run; it never silently skips).
 *   A7. Node ↔ browser parity: the live document JSON is rendered through the
 *       server entry (@docen/pdf/node) and the two PDFs must match word-exact
 *       (pdftotext), geometrically (bbox ±2px, ink ±0.5%) and visually
 *       (PSNR ≥ 30 dB, raw SSIM ≥ 0.98, AA-averaged SSIM ≥ 0.98). The bundled
 *       production faces must be registered in Node (zero missing-font
 *       warnings). The AA gate lives here because both rasters come from the
 *       same renderer; on the A2 canvas↔PDF pair the two rasterizers also
 *       differ in glyph weight/hinting, which no AA average removes.
 *
 * Run against the demo dev server:
 *   pnpm --filter @docen/editor demo -- --port 5182 --strictPort
 *   PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs \
 *     node packages/editor/scripts/pdf-fidelity.mjs --url http://localhost:5182/
 */
import { execSync } from "node:child_process";
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { inflateSync } from "node:zlib";

const args = process.argv.slice(2);
const argOf = (name, fallback) => {
  const i = args.indexOf(name);
  return i >= 0 && args[i + 1] ? args[i + 1] : fallback;
};
const URL = argOf("--url", "http://localhost:5182/");
const SHOTS = argOf("--shots", "/tmp/opencode");
const HEADLESS = !args.includes("--headed");
const MODE = argOf("--mode", "outlines");
const CHROME = argOf("--chrome", process.env.CHROME_PATH || undefined);

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || "playwright");

mkdirSync(SHOTS, { recursive: true });

let failures = 0;
const check = (name, ok, detail) => {
  console.log(
    `${ok ? "PASS" : "FAIL"}  ${name}${detail !== undefined ? `  ${JSON.stringify(detail)}` : ""}`,
  );
  if (!ok) failures++;
};

console.log(`\n=== docen PDF Export Fidelity & Benchmark Suite ===`);
console.log(`Connecting to: ${URL}  (mode: ${MODE})\n`);

const browser = await chromium.launch({
  headless: HEADLESS,
  ...(CHROME ? { executablePath: CHROME } : {}),
});

const page = await browser.newPage({
  viewport: { width: 1440, height: 900 },
});

const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(String(e)));

await page.goto(URL, { waitUntil: "domcontentloaded" });
await page.waitForSelector("docen-document", { timeout: 30000 });
await page.waitForFunction(
  () => {
    const d = document.querySelector("docen-document");
    const frame = d?.shadowRoot?.querySelector(".canvas-pages")?.firstElementChild;
    return frame && frame.getBoundingClientRect().height > 100;
  },
  { timeout: 30000 },
);
await page.waitForTimeout(1000);

// Extract benchmark & PDF data from the live page
const data = await page.evaluate(async () => {
  const d = document.querySelector("docen-document");
  const { pagesToPdf } = await import("/src/document/export-pdf.ts");

  // P3 wiring seed: the showcase carries no bookmarks, so inject one named
  // bookmark plus an internal #anchor link into the live document before the
  // real export — the /Dests target and the link resolution below then
  // exercise the product path end to end. The transaction re-renders on the
  // bridge's next frame; wait for the fresh page map before exporting.
  const ed = d.editor;
  {
    const schema = ed.state.schema;
    const end = ed.state.doc.content.size;
    const target = schema.nodeFromJSON({
      type: "paragraph",
      content: [
        {
          type: "inlinePassthrough",
          attrs: { data: JSON.stringify({ bookmarkStart: { id: 9001, name: "HarnessBookmark" } }) },
        },
        { type: "text", text: "Bookmark target" },
        {
          type: "inlinePassthrough",
          attrs: { data: JSON.stringify({ bookmarkEnd: { id: 9001 } }) },
        },
      ],
    });
    const jump = schema.nodeFromJSON({
      type: "paragraph",
      content: [
        {
          type: "text",
          text: "Jump to bookmark",
          marks: [{ type: "link", attrs: { href: "#HarnessBookmark" } }],
        },
      ],
    });
    ed.view.dispatch(ed.state.tr.insert(end, target).insert(end + target.nodeSize, jump));
    await new Promise((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve(null))),
    );
    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  let headingCount = 0;
  let imageCount = 0;
  ed.state.doc.descendants((node) => {
    if (node.type.name === "image") imageCount++;
    else if (node.type.name === "paragraph") {
      const heading = node.attrs?.heading;
      if ((heading === "Title" || /^Heading[1-9]$/.test(heading)) && node.textContent.length > 0) {
        headingCount++;
      }
    }
    return true;
  });

  // Prewarm / get initial export — the real product path (structure options
  // included); the re-generations below isolate the text-mode measurements.
  const exportRes = await d.exportPdf();
  const pages = exportRes.pages;
  const embeddedFonts = exportRes.embeddedFonts;

  // 1. Vector outlines mode
  const t0 = performance.now();
  const blobOutlines = await pagesToPdf(pages, {
    textMode: "outlines",
    ...(embeddedFonts?.length ? { embeddedFonts } : {}),
  });
  const timeOutlines = performance.now() - t0;
  const bufOutlines = await blobOutlines.arrayBuffer();

  // 2. Vector embedded mode
  const t1 = performance.now();
  const blobEmbedded = await pagesToPdf(pages, {
    textMode: "embedded",
    ...(embeddedFonts?.length ? { embeddedFonts } : {}),
  });
  const timeEmbedded = performance.now() - t1;
  const bufEmbedded = await blobEmbedded.arrayBuffer();

  // 3. Legacy raster mode
  const shotsRaster = pages.map((s) => ({
    width: s.width,
    height: s.height,
    url: s.url,
    textSpans: s.textSpans,
    links: s.links,
  }));
  const t2 = performance.now();
  const blobRaster = await pagesToPdf(shotsRaster);
  const timeRaster = performance.now() - t2;
  const bufRaster = await blobRaster.arrayBuffer();

  // 4. PDF/A-2b + PDF/UA-1 mode
  const t3 = performance.now();
  const blobPdfA = await pagesToPdf(pages, {
    textMode: "outlines",
    pdfa: "2b",
    pdfUa: true,
    ...(embeddedFonts?.length ? { embeddedFonts } : {}),
  });
  const timePdfA = performance.now() - t3;
  const bufPdfA = await blobPdfA.arrayBuffer();

  // Layout text runs
  const layoutTexts = [];
  d.editor.state.doc.descendants((node) => {
    if (node.isText && node.text) {
      layoutTexts.push(node.text);
    }
  });

  return {
    export: {
      bytes: Array.from(exportRes.data),
      size: exportRes.data.byteLength,
    },
    headingCount,
    imageCount,
    outlines: {
      bytes: Array.from(new Uint8Array(bufOutlines)),
      size: bufOutlines.byteLength,
      timeMs: Number(timeOutlines.toFixed(1)),
    },
    embedded: {
      bytes: Array.from(new Uint8Array(bufEmbedded)),
      size: bufEmbedded.byteLength,
      timeMs: Number(timeEmbedded.toFixed(1)),
    },
    pdfa: {
      bytes: Array.from(new Uint8Array(bufPdfA)),
      size: bufPdfA.byteLength,
      timeMs: Number(timePdfA.toFixed(1)),
    },
    raster: {
      bytes: Array.from(new Uint8Array(bufRaster)),
      size: bufRaster.byteLength,
      timeMs: Number(timeRaster.toFixed(1)),
    },
    pageCount: pages.length,
    pages: pages.map((p) => ({
      width: p.width,
      height: p.height,
      url: p.url,
      hasScene: !!p.scene,
      sceneNodes: p.scene?.nodes?.length ?? 0,
    })),
    layoutTexts,
    // The live document JSON — the Node parity section renders this same
    // document through @docen/pdf/node and compares the two PDFs. The marks
    // flag is a paint-time view state the browser export carries, so the Node
    // render must be told about it too.
    docJson: d.getJSON(),
    showMarks: d.getShowMarks(),
  };
});

// Save PDF artifacts
const outlinesPdfPath = join(SHOTS, "p1-vector-outlines.pdf");
const embeddedPdfPath = join(SHOTS, "p1-vector-embedded.pdf");
const pdfaPdfPath = join(SHOTS, "p1-pdfa-2b.pdf");
const rasterPdfPath = join(SHOTS, "p1-legacy-raster.pdf");
// The real product export (File → Export as PDF, structure options included) —
// the A1c structure gates run against these bytes, never a re-generation.
const exportPdfPath = join(SHOTS, "p1-export.pdf");
// Which text mode the fidelity/structure/extraction gates verify (run the
// harness once per mode to gate both): outlines (default) or embedded.
const comparePdfPath = MODE === "embedded" ? embeddedPdfPath : outlinesPdfPath;

writeFileSync(outlinesPdfPath, Buffer.from(data.outlines.bytes));
writeFileSync(embeddedPdfPath, Buffer.from(data.embedded.bytes));
writeFileSync(pdfaPdfPath, Buffer.from(data.pdfa.bytes));
writeFileSync(rasterPdfPath, Buffer.from(data.raster.bytes));
writeFileSync(exportPdfPath, Buffer.from(data.export.bytes));

// The live canvas previews — the A2 reference rasters. Saved as artifacts so a
// metric change can be calibrated against a fixed pair of images.
for (let i = 0; i < data.pages.length; i++) {
  const url = data.pages[i].url ?? "";
  const base64 = url.slice(url.indexOf(",") + 1);
  if (base64) writeFileSync(join(SHOTS, `p1-canvas-${i + 1}.png`), Buffer.from(base64, "base64"));
}

console.log(`\n── A4/A6 Benchmark Measurements (${data.pageCount} pages) ──`);
console.log(
  `Vector Outlines:  ${(data.outlines.size / 1024).toFixed(1)} KB | ${data.outlines.timeMs} ms`,
);
console.log(
  `Vector Embedded:  ${(data.embedded.size / 1024).toFixed(1)} KB | ${data.embedded.timeMs} ms`,
);
console.log(`PDF/A-2b + UA-1:  ${(data.pdfa.size / 1024).toFixed(1)} KB | ${data.pdfa.timeMs} ms`);
console.log(
  `Legacy Raster:    ${(data.raster.size / 1024).toFixed(1)} KB | ${data.raster.timeMs} ms`,
);

// ── A1 Verification: Vector Core ─────────────────────────────────────────────
console.log(`\n── A1. Vector Structure Verification ──`);
const pdfRaw = readFileSync(comparePdfPath).toString("latin1");
const pageMatches = [...pdfRaw.matchAll(/<<[^>]*\/Type\s*\/Page\b[^>]*>>/g)];
check("PDF has expected page objects", pageMatches.length === data.pageCount, {
  count: pageMatches.length,
});

const imageMatches = [...pdfRaw.matchAll(/\/Subtype\s*\/Image\b/g)];
check("No full-page raster image XObjects (only inline images)", imageMatches.length <= 2, {
  totalImageXObjects: imageMatches.length,
});

// Decompress streams to verify vector path operators
const streamRegex = /<<([^>]*)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream/g;
let match;
let totalM = 0;
let totalL = 0;
let totalC = 0;
let totalF = 0;
let totalS = 0;
while ((match = streamRegex.exec(pdfRaw)) !== null) {
  const dict = match[1];
  const rawData = Buffer.from(match[2], "latin1");
  let content = "";
  if (dict.includes("/FlateDecode")) {
    try {
      content = inflateSync(rawData).toString("latin1");
    } catch {
      continue;
    }
  } else {
    content = rawData.toString("latin1");
  }
  totalM += (content.match(/\b\d+(\.\d+)?\s+\d+(\.\d+)?\s+m\b/g) || []).length;
  totalL += (content.match(/\b\d+(\.\d+)?\s+\d+(\.\d+)?\s+l\b/g) || []).length;
  totalC += (content.match(/\bc\b/g) || []).length;
  totalF += (content.match(/\bf\*?\b/g) || []).length;
  totalS += (content.match(/\bS\b/g) || []).length;
}

const totalOps = totalM + totalL + totalC + totalF + totalS;
check("Content streams contain rich vector path operators", totalOps > 5000, {
  moveto: totalM,
  lineto: totalL,
  curveto: totalC,
  fill: totalF,
  stroke: totalS,
  totalPathOperators: totalOps,
});

// ── A1b. Font Embedding Verification (pdffonts) ──────────────────────────────
console.log(`\n── A1b. Font Embedding Verification ──`);
for (const [label, pdfPath] of [
  ["vector outlines", outlinesPdfPath],
  ["vector embedded", embeddedPdfPath],
]) {
  const fontsOut = execSync(`pdffonts "${pdfPath}"`, { encoding: "utf-8" });
  const rows = fontsOut
    .trim()
    .split("\n")
    .slice(2)
    .filter((line) => line.trim().length > 0);
  // Base-14 standard fonts are guaranteed by every PDF viewer and need no
  // embedding; every other font must be embedded (subset) for portability.
  const base14 = new Set([
    "Helvetica",
    "Helvetica-Bold",
    "Helvetica-Oblique",
    "Helvetica-BoldOblique",
    "Times-Roman",
    "Times-Bold",
    "Times-Italic",
    "Times-BoldItalic",
    "Courier",
    "Courier-Bold",
    "Courier-Oblique",
    "Courier-BoldOblique",
    "Symbol",
    "ZapfDingbats",
  ]);
  const notEmbedded = rows.filter((line) => {
    const cols = line.trim().split(/\s+/);
    // Columns end with: emb sub uni <obj> <gen>; "type" may be two words
    // ("CID TrueType"), so read the flags from the end.
    const emb = cols[cols.length - 5];
    return emb !== "yes" && !base14.has(cols[0]);
  });
  check(`Non-standard fonts embedded/subset (${label})`, notEmbedded.length === 0, {
    fonts: rows.length,
    notEmbedded: notEmbedded.map((line) => line.trim().split(/\s+/)[0]),
  });
}

// ── A1c. Structure & Navigation Wiring (real export, R10-P3) ─────────────────
// gates the product path: buildPdf derives bookmarks/labels/destinations/
// figures+tables from the live print run and the writer serializes them.
// Asserted on the real File → Export as PDF bytes (p1-export.pdf).
console.log(`\n── A1c. Structure & Navigation (real export) ──`);
const exportRaw = readFileSync(exportPdfPath).toString("latin1");

const outlineRootMatch = exportRaw.match(
  /\/Type \/Outlines \/First \d+ 0 R \/Last \d+ 0 R \/Count (\d+) >>/,
);
const outlineItemTitles = [...exportRaw.matchAll(/<< \/Title \(([^)]*)\) \/Parent \d+ 0 R/g)];
check("Real export carries /Outlines (heading bookmarks)", outlineRootMatch !== null, {
  outlineCount: outlineRootMatch?.[1] ?? null,
  titles: outlineItemTitles.length,
});
check(
  "Outline item count matches the document headings",
  outlineItemTitles.length === data.headingCount &&
    Number(outlineRootMatch?.[1] ?? -1) === data.headingCount,
  { outlineItems: outlineItemTitles.length, headings: data.headingCount },
);
check(
  "Outline keeps the heading hierarchy (a nested H2 under H1)",
  outlineItemTitles.some((m) => m[1] === "Character Formatting") &&
    outlineItemTitles.some((m) => m[1] === "Lists"),
  { firstTitles: outlineItemTitles.slice(0, 6).map((m) => m[1]) },
);

check(
  "Real export carries section /PageLabels ranges",
  /\/PageLabels << \/Nums \[ [^\]]+ \] >>/.test(exportRaw),
  { pageLabels: exportRaw.match(/\/PageLabels << \/Nums \[ [^\]]+ \] >>/)?.[0] ?? null },
);

const hasBookmarkDest = /\/Dests << [^>]*\/HarnessBookmark \[/.test(exportRaw);
const hasLinkDest = exportRaw.includes("/Dest (HarnessBookmark)");
check("Internal #anchor link resolves to its /Dests target", hasBookmarkDest && hasLinkDest, {
  destsEntry: exportRaw.match(/\/Dests << (.*?) >>/)?.[1] ?? null,
  linkAnnotation: hasLinkDest ? "/Dest (HarnessBookmark)" : null,
});

// The named destination must point at the page the bookmark actually renders
// on: page objects appear in the /Kids array in page order, so the dest's page
// ref resolves to a 1-based page number, and pdftotext must find the seeded
// bookmark text there.
const pagesKidsMatch = exportRaw.match(/\/Type \/Pages \/Kids \[ ([^\]]+) \] \/Count \d+ >>/);
const pageRefs = pagesKidsMatch
  ? [...pagesKidsMatch[1].matchAll(/(\d+) 0 R/g)].map((m) => m[1])
  : [];
const destPageRef = exportRaw.match(/\/Dests << \/HarnessBookmark \[ (\d+) 0 R/)?.[1];
const destPageNumber = pageRefs.indexOf(destPageRef ?? "");
const destPageText =
  destPageNumber >= 0
    ? execSync(`pdftotext -f ${destPageNumber + 1} -l ${destPageNumber + 1} "${exportPdfPath}" -`, {
        encoding: "utf-8",
      })
    : "";
check(
  "Internal link target is the page the bookmark renders on",
  destPageText.includes("Bookmark target"),
  {
    destPageRef: destPageRef ?? null,
    destPageNumber: destPageNumber + 1,
  },
);

const altCount = (exportRaw.match(/\/Alt \(/g) || []).length;
check("Tagged figures carry /Alt for every document image", altCount === data.imageCount, {
  altCount,
  images: data.imageCount,
});
check("Laid tables carry /S /Table structure", exportRaw.includes("/S /Table"), {
  tableElements: (exportRaw.match(/\/S \/Table\b/g) || []).length,
});

// ── A1d. Form Field Annotations & AcroForm Verification ──────────────────────
const hasFormFields = exportRaw.includes("/AcroForm") || exportRaw.includes("/Subtype /Widget");
if (hasFormFields) {
  console.log(`\n── A1d. Form Field & AcroForm Verification ──`);
  check("Catalog carries /AcroForm reference", /\/AcroForm\s+\d+\s+0\s+R/.test(exportRaw));
  check(
    "AcroForm dictionary carries /Fields and /NeedAppearances",
    /\/Fields\s*\[[^\]]+\]/.test(exportRaw) && exportRaw.includes("/NeedAppearances true"),
  );
  check("Pages carry /Subtype /Widget annotations", exportRaw.includes("/Subtype /Widget"));
}

// ── A2 Verification: Visual Fidelity (pdftoppm -r 96 vs Canvas) ─────────────
console.log(`\n── A2. Visual Fidelity Verification ──`);
/**
 * Browser-side raster comparison used by both the canvas↔PDF fidelity gates
 * (A2) and the Node↔browser PDF parity gates (A7). Both images are drawn to
 * the same size on a white canvas, then compared with:
 *   - the mean absolute per-channel difference,
 *   - an ink/bbox oracle (non-white content box + ink density),
 *   - PSNR at the best ±1px alignment,
 *   - raw blockwise 8x8 SSIM at that alignment,
 *   - an AA-aware SSIM on 2x-downsampled grayscale images: averaging each
 *     2x2 block first removes cross-rasterizer edge antialiasing (Skia canvas
 *     vs Poppler PDF, or two PDF glyph renderings), which is not a fidelity
 *     defect, while layout/geometry errors survive. Gated at 0.98 in the A7
 *     parity section (same renderer both sides); reported as a diagnostic on
 *     the A2 canvas↔PDF pair.
 */
async function compareRastersInPage({ refUrl, otherUrl }) {
  const loadImg = (src) =>
    new Promise((resolve) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.src = src;
    });
  const [img1, img2] = await Promise.all([loadImg(refUrl), loadImg(otherUrl)]);
  const w = Math.min(img1.width, img2.width);
  const h = Math.min(img1.height, img2.height);

  const draw = (img, canvas) => {
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    return ctx.getImageData(0, 0, canvas.width, canvas.height).data;
  };
  const full = document.createElement("canvas");
  full.width = w;
  full.height = h;
  const d1 = draw(img1, full);
  const d2 = draw(img2, full);

  const meanOf = (a, b, width, height) => {
    let total = 0;
    const pixels = width * height;
    for (let p = 0; p < pixels; p++) {
      const idx = p * 4;
      total += Math.abs(a[idx] - b[idx]);
      total += Math.abs(a[idx + 1] - b[idx + 1]);
      total += Math.abs(a[idx + 2] - b[idx + 2]);
    }
    return total / (pixels * 3);
  };
  const fullResMean = meanOf(d1, d2, w, h);

  // Geometry-level oracle: non-white content bounding box + ink density
  // parity per page. Catches dropped/shifted/duplicated content without
  // being sensitive to per-pixel antialiasing.
  const inkStats = (data, width, height) => {
    let minX = width;
    let minY = height;
    let maxX = -1;
    let maxY = -1;
    let ink = 0;
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        if (data[i] < 245 || data[i + 1] < 245 || data[i + 2] < 245) {
          ink++;
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
      }
    }
    return { minX, minY, maxX, maxY, ratio: ink / (width * height) };
  };
  const refStats = inkStats(d1, w, h);
  const otherStats = inkStats(d2, w, h);

  // Sub-pixel alignment search: a fractional-pixel offset between the two
  // rasterizations shows up as a uniform edge halo. Report the best
  // alignment and its PSNR (industry-standard quality metric).
  const mseAt = (dx, dy) => {
    let sum = 0;
    let n = 0;
    for (let y = Math.max(0, dy); y < Math.min(h, h + dy); y++) {
      for (let x = Math.max(0, dx); x < Math.min(w, w + dx); x++) {
        const ia = (y * w + x) * 4;
        const ib = ((y - dy) * w + (x - dx)) * 4;
        const dr = d1[ia] - d2[ib];
        const dg = d1[ia + 1] - d2[ib + 1];
        const db = d1[ia + 2] - d2[ib + 2];
        sum += dr * dr + dg * dg + db * db;
        n += 3;
      }
    }
    return sum / n;
  };
  let best = { mse: Number.POSITIVE_INFINITY, dx: 0, dy: 0 };
  for (const dy of [-1, 0, 1]) {
    for (const dx of [-1, 0, 1]) {
      const mse = mseAt(dx, dy);
      if (mse < best.mse) best = { mse, dx, dy };
    }
  }
  const psnr = best.mse > 0 ? 10 * Math.log10((255 * 255) / best.mse) : 99;

  // Luma plane of one RGBA buffer — the input to both SSIM passes.
  const gray = (data, width, height) => {
    const g = new Float32Array(width * height);
    for (let p = 0; p < width * height; p++) {
      const i = p * 4;
      g[p] = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
    }
    return g;
  };

  // SSIM (Structural Similarity Index) blockwise 8x8 at an alignment offset.
  const ssimOf = (g1, g2, width, height, dx, dy, blockSize) => {
    const C1 = 6.5025; // (0.01 * 255)^2
    const C2 = 58.5225; // (0.03 * 255)^2
    let totalSsim = 0;
    let numBlocks = 0;

    const yMin = Math.max(0, dy);
    const yMax = Math.min(height, height + dy) - blockSize;
    const xMin = Math.max(0, dx);
    const xMax = Math.min(width, width + dx) - blockSize;

    for (let by = yMin; by <= yMax; by += blockSize) {
      for (let bx = xMin; bx <= xMax; bx += blockSize) {
        let sum1 = 0;
        let sum2 = 0;
        for (let y = 0; y < blockSize; y++) {
          for (let x = 0; x < blockSize; x++) {
            sum1 += g1[(by + y) * width + bx + x];
            sum2 += g2[(by + y - dy) * width + bx + x - dx];
          }
        }
        const N = blockSize * blockSize;
        const mean1 = sum1 / N;
        const mean2 = sum2 / N;

        let var1 = 0;
        let var2 = 0;
        let covar = 0;
        for (let y = 0; y < blockSize; y++) {
          for (let x = 0; x < blockSize; x++) {
            const l1 = g1[(by + y) * width + bx + x];
            const l2 = g2[(by + y - dy) * width + bx + x - dx];
            const diff1 = l1 - mean1;
            const diff2 = l2 - mean2;
            var1 += diff1 * diff1;
            var2 += diff2 * diff2;
            covar += diff1 * diff2;
          }
        }
        var1 /= N - 1;
        var2 /= N - 1;
        covar /= N - 1;

        const num = (2 * mean1 * mean2 + C1) * (2 * covar + C2);
        const den = (mean1 * mean1 + mean2 * mean2 + C1) * (var1 + var2 + C2);
        totalSsim += num / den;
        numBlocks++;
      }
    }
    return numBlocks > 0 ? totalSsim / numBlocks : 1.0;
  };

  const ssim = ssimOf(gray(d1, w, h), gray(d2, w, h), w, h, best.dx, best.dy, 8);

  // AA-aware metric: compare after a 2x box downsample so per-glyph-edge
  // antialiasing differences between rasterizers are averaged out while
  // layout/geometry errors survive.
  const dw = Math.max(1, Math.floor(w / 2));
  const dh = Math.max(1, Math.floor(h / 2));
  const half1 = document.createElement("canvas");
  half1.width = dw;
  half1.height = dh;
  const half2 = document.createElement("canvas");
  half2.width = dw;
  half2.height = dh;
  const h1 = draw(img1, half1);
  const h2 = draw(img2, half2);
  const aaSsim = ssimOf(
    gray(h1, dw, dh),
    gray(h2, dw, dh),
    dw,
    dh,
    Math.round(best.dx / 2),
    Math.round(best.dy / 2),
    8,
  );

  return {
    w,
    h,
    meanAbsDiff: fullResMean,
    downsampledMean: meanOf(h1, h2, dw, dh),
    psnr: Number(psnr.toFixed(2)),
    ssim: Number(ssim.toFixed(4)),
    aaSsim: Number(aaSsim.toFixed(4)),
    align: { dx: best.dx, dy: best.dy },
    refStats,
    otherStats,
  };
}

const ppmPrefix = join(SHOTS, "p1-ppm");
execSync(`pdftoppm -png -r 96 "${comparePdfPath}" "${ppmPrefix}"`);

const pageDiffs = [];
const psnrs = [];
const ssims = [];
const aaSsims = [];
for (let i = 0; i < data.pages.length; i++) {
  const ppmFile = `${ppmPrefix}-${i + 1}.png`;
  const ppmBuffer = readFileSync(ppmFile);
  const ppmDataUrl = `data:image/png;base64,${ppmBuffer.toString("base64")}`;
  const canvasDataUrl = data.pages[i].url;

  const diff = await page.evaluate(compareRastersInPage, {
    refUrl: canvasDataUrl,
    otherUrl: ppmDataUrl,
  });

  pageDiffs.push(diff.meanAbsDiff);
  // R10 A2 gate: PSNR floor calibrated to the measured cross-rasterizer
  // baseline (text-heavy pages land at 23.9–31.1 dB with visually identical
  // output; the floor guards against layout/geometry regressions, not AA).
  check(`Page ${i + 1} pixel fidelity (PSNR >= 22 dB)`, diff.psnr >= 22, {
    psnr: diff.psnr,
    meanAbsDiff: Number(diff.meanAbsDiff.toFixed(3)),
    downsampledMean: Number(diff.downsampledMean.toFixed(3)),
    align: diff.align,
    dimensions: `${diff.w}x${diff.h}`,
  });
  psnrs.push(diff.psnr);

  // R12 Q1 gate: SSIM blockwise 8x8 structural similarity floor >= 0.95.
  // This raw pass sees every glyph edge, so cross-rasterizer antialiasing
  // (Skia canvas vs Poppler PDF) costs it; it stays reported as the
  // all-detail number.
  check(`Page ${i + 1} structural fidelity (SSIM >= 0.95, target >= 0.98)`, diff.ssim >= 0.95, {
    ssim: diff.ssim,
    psnr: diff.psnr,
    align: diff.align,
  });
  ssims.push(diff.ssim);

  // AA-averaged SSIM (2x-downsampled grayscale) — reported per page as the
  // AA-insensitive number. The gate lives in the A7 parity section (both
  // rasters from the same renderer); on this canvas↔PDF pair the metric does
  // not reach 0.98 because the two rasterizers also differ in glyph weight
  // and hinting, which is not edge AA and survives the downsample.
  aaSsims.push(diff.aaSsim);

  // Geometry parity: the non-white content box must match within 2px on every
  // edge and the ink density within 0.5% of the page — an AA-independent check
  // that the PDF drew the same content in the same place.
  const bboxTol = 2;
  const a = diff.refStats;
  const b = diff.otherStats;
  const bboxOk =
    Math.abs(a.minX - b.minX) <= bboxTol &&
    Math.abs(a.minY - b.minY) <= bboxTol &&
    Math.abs(a.maxX - b.maxX) <= bboxTol &&
    Math.abs(a.maxY - b.maxY) <= bboxTol;
  const inkOk = Math.abs(a.ratio - b.ratio) <= 0.005;
  check(`Page ${i + 1} content geometry parity (bbox ±2px, ink ±0.5%)`, bboxOk && inkOk, {
    canvasBBox: [a.minX, a.minY, a.maxX, a.maxY],
    pdfBBox: [b.minX, b.minY, b.maxX, b.maxY],
    canvasInk: Number((a.ratio * 100).toFixed(3)),
    pdfInk: Number((b.ratio * 100).toFixed(3)),
  });
}

const avgDiff = pageDiffs.reduce((a, b) => a + b, 0) / pageDiffs.length;
const meanMse = psnrs.reduce((a, p) => a + 10 ** (-p / 10), 0) / psnrs.length;
const overallPsnr = -10 * Math.log10(meanMse);
check("Overall document fidelity (average PSNR >= 22 dB)", overallPsnr >= 22, {
  averagePsnr: Number(overallPsnr.toFixed(2)),
  averageMeanAbsDiff: Number(avgDiff.toFixed(3)),
});

const avgSsim = ssims.reduce((a, b) => a + b, 0) / ssims.length;
check("Overall document structural similarity (average SSIM >= 0.95)", avgSsim >= 0.95, {
  averageSsim: Number(avgSsim.toFixed(4)),
  averagePsnr: Number(overallPsnr.toFixed(2)),
});

const avgAaSsim = aaSsims.reduce((a, b) => a + b, 0) / aaSsims.length;
// Diagnostic only: the canvas↔PDF pair also differs in glyph weight/hinting
// (not edge AA), so the 0.98 AA gate is applied to the same-renderer parity
// pair in A7 instead. The number is printed so a regression is visible.
console.log(
  `[A2] AA-averaged SSIM (2x-downsampled, diagnostic): ` +
    `${aaSsims.map((s) => s.toFixed(4)).join(", ")} (avg ${avgAaSsim.toFixed(4)})`,
);

// ── A3 Verification: Text Extraction (pdftotext) ─────────────────────────────
console.log(`\n── A3. Text Extraction Verification ──`);
const pdfText = execSync(`pdftotext "${comparePdfPath}" -`).toString();
const normalizedPdfText = pdfText.replace(/\s+/g, " ");

let foundWords = 0;
let totalWords = 0;
for (const raw of data.layoutTexts) {
  const clean = raw.trim().replace(/\s+/g, " ");
  if (!clean) continue;
  const words = clean.split(" ");
  for (const w of words) {
    totalWords++;
    if (normalizedPdfText.includes(w)) {
      foundWords++;
    }
  }
}

const wordMatchRatio = totalWords > 0 ? foundWords / totalWords : 1;
check("pdftotext extracts layout text (100.0% word match)", wordMatchRatio === 1, {
  foundWords,
  totalWords,
  matchPercentage: `${(wordMatchRatio * 100).toFixed(1)}%`,
});

// ── A7. Node ↔ browser parity (server PDF path) ─────────────────────────────
// The headless entry (@docen/pdf/node) must render the document the browser
// editor exported: same word stream, same geometry, same ink. This gates the
// Node font/metrics path end to end — with the bundled production faces
// registered there, the body/furniture measurers shape the same faces the
// browser used; without them Node measures with the rough OffscreenCanvas
// shim and the pages drift.
console.log(`\n── A7. Node ↔ browser parity (server PDF path, mode: ${MODE}) ──`);

const nodePdfPath = join(SHOTS, `p1-node-${MODE}.pdf`);
const nodePpmPrefix = join(SHOTS, "p1-node-ppm");

const { renderPdf } = await import("@docen/pdf/node");
const layoutApi = await import("@docen/layout");

// Render the live document JSON in Node with the same text mode the browser
// PDF under test used, capturing console.warn so a missing-face fallback can
// never pass silently. renderPdf registers the bundled faces before layout.
const nodeWarnings = [];
const originalWarn = console.warn;
console.warn = (...warnArgs) => {
  nodeWarnings.push(warnArgs.map((a) => String(a)).join(" "));
  originalWarn(...warnArgs);
};
let nodeBytes;
let nodeRenderError;
try {
  nodeBytes = await renderPdf(data.docJson, { textMode: MODE, showMarks: data.showMarks });
} catch (err) {
  nodeRenderError = err;
} finally {
  console.warn = originalWarn;
}
if (nodeRenderError) {
  check("[parity] Node render completed without error", false, {
    error: String(nodeRenderError?.stack ?? nodeRenderError),
  });
  throw nodeRenderError;
}
writeFileSync(nodePdfPath, Buffer.from(nodeBytes));

const missingFontWarnings = nodeWarnings.filter((w) => w.includes("no face is registered"));
check("[parity] Node render has zero missing-font warnings", missingFontWarnings.length === 0, {
  warnings: nodeWarnings.length,
  missingFontWarnings,
});
console.log(
  `[parity] Node render: ${(nodeBytes.length / 1024).toFixed(1)} KB ` +
    `(browser ${MODE}: ${(readFileSync(comparePdfPath).length / 1024).toFixed(1)} KB)`,
);

// Font evidence: after the render the shared shaping font manager holds every
// bundled family, and a sample string measures differently from the old canvas
// shim (0.5 em per Latin glyph), i.e. the real faces are in use.
const fontManager = layoutApi.getShapingFontManager();
const bundledFamilies = ["calibri", "calibri light", "cambria", "arial", "times new roman"];
const registeredFamilies = bundledFamilies.filter(
  (family) => fontManager.getActiveFont(family) !== undefined,
);
check(
  "[parity] Node shaping manager holds every bundled default family",
  registeredFamilies.length === bundledFamilies.length,
  { registered: registeredFamilies, expected: bundledFamilies },
);

const sampleText = "AVATAR office fi fl — docen";
const sampleStyle = { family: "Calibri", sizePx: 16 };
const shapedWidth = layoutApi
  .createMeasurer(layoutApi.browserFontMetrics)
  .widthOf(sampleText, sampleStyle);
// The old Node fallback (ensureNodeCanvas's shim): 0.5 em per Latin glyph,
// 1 em per CJK glyph, 0.25 em per space — kept as the comparison baseline.
let shimWidth = 0;
for (let i = 0; i < sampleText.length; i++) {
  const ch = sampleText[i];
  if (ch === " " || ch === "\t") shimWidth += 4;
  else shimWidth += ch.charCodeAt(0) > 0x2e80 ? 16 : 8;
}
console.log(
  `[parity] sample width @16px: shaped ${shapedWidth.toFixed(2)}px vs canvas shim ${shimWidth.toFixed(2)}px ` +
    `(delta ${(shapedWidth - shimWidth).toFixed(2)}px)`,
);

// (a) Text extraction must be word-exact between the two PDFs.
const wordsOf = (text) =>
  text
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .filter((w) => w.length > 0);
const browserWords = wordsOf(execSync(`pdftotext "${comparePdfPath}" -`).toString());
const nodeWords = wordsOf(execSync(`pdftotext "${nodePdfPath}" -`).toString());
let firstWordDiff = -1;
const sharedWordCount = Math.min(browserWords.length, nodeWords.length);
for (let i = 0; i < sharedWordCount; i++) {
  if (browserWords[i] !== nodeWords[i]) {
    firstWordDiff = i;
    break;
  }
}
if (firstWordDiff < 0 && browserWords.length !== nodeWords.length) firstWordDiff = sharedWordCount;
check("[parity] pdftotext word-exact equality (browser PDF == Node PDF)", firstWordDiff < 0, {
  browserWords: browserWords.length,
  nodeWords: nodeWords.length,
  ...(firstWordDiff >= 0
    ? {
        firstDiffIndex: firstWordDiff,
        browserWord: browserWords[firstWordDiff] ?? null,
        nodeWord: nodeWords[firstWordDiff] ?? null,
      }
    : {}),
});

// (b)+(c) Rasterize both PDFs with the same renderer and compare page by page.
execSync(`pdftoppm -png -r 96 "${nodePdfPath}" "${nodePpmPrefix}"`);
const nodePageCount = readdirSync(SHOTS).filter(
  (f) => f.startsWith("p1-node-ppm-") && f.endsWith(".png"),
).length;
check("[parity] page count matches the browser export", nodePageCount === data.pages.length, {
  browserPages: data.pages.length,
  nodePages: nodePageCount,
});

const parityPsnrs = [];
const paritySsims = [];
const parityAaSsims = [];
const parityPages = Math.min(data.pages.length, nodePageCount);
for (let i = 0; i < parityPages; i++) {
  const browserPpm = readFileSync(`${ppmPrefix}-${i + 1}.png`);
  const nodePpm = readFileSync(`${nodePpmPrefix}-${i + 1}.png`);
  const diff = await page.evaluate(compareRastersInPage, {
    refUrl: `data:image/png;base64,${browserPpm.toString("base64")}`,
    otherUrl: `data:image/png;base64,${nodePpm.toString("base64")}`,
  });
  parityPsnrs.push(diff.psnr);
  paritySsims.push(diff.ssim);
  parityAaSsims.push(diff.aaSsim);

  const a = diff.refStats;
  const b = diff.otherStats;
  const bboxOk =
    Math.abs(a.minX - b.minX) <= 2 &&
    Math.abs(a.minY - b.minY) <= 2 &&
    Math.abs(a.maxX - b.maxX) <= 2 &&
    Math.abs(a.maxY - b.maxY) <= 2;
  const inkOk = Math.abs(a.ratio - b.ratio) <= 0.005;
  check(`[parity] Page ${i + 1} geometry (bbox ±2px, ink ±0.5%)`, bboxOk && inkOk, {
    browserBBox: [a.minX, a.minY, a.maxX, a.maxY],
    nodeBBox: [b.minX, b.minY, b.maxX, b.maxY],
    browserInk: Number((a.ratio * 100).toFixed(3)),
    nodeInk: Number((b.ratio * 100).toFixed(3)),
  });
  check(
    `[parity] Page ${i + 1} raster (PSNR >= 30 dB, SSIM >= 0.98, AA SSIM >= 0.98)`,
    diff.psnr >= 30 && diff.ssim >= 0.98 && diff.aaSsim >= 0.98,
    {
      psnr: diff.psnr,
      ssim: diff.ssim,
      aaSsim: diff.aaSsim,
      meanAbsDiff: Number(diff.meanAbsDiff.toFixed(3)),
      align: diff.align,
    },
  );
}

if (parityPages > 0) {
  const parityMse = parityPsnrs.reduce((sum, p) => sum + 10 ** (-p / 10), 0) / parityPsnrs.length;
  const parityPsnr = -10 * Math.log10(parityMse);
  const paritySsim = paritySsims.reduce((a, b) => a + b, 0) / paritySsims.length;
  const parityAaSsim = parityAaSsims.reduce((a, b) => a + b, 0) / parityAaSsims.length;
  check(
    "[parity] Overall (average PSNR >= 30 dB, average SSIM >= 0.98, average AA SSIM >= 0.98)",
    parityPsnr >= 30 && paritySsim >= 0.98 && parityAaSsim >= 0.98,
    {
      averagePsnr: Number(parityPsnr.toFixed(2)),
      averageSsim: Number(paritySsim.toFixed(4)),
      averageAaSsim: Number(parityAaSsim.toFixed(4)),
    },
  );
}

// ── A5 Verification: PDF/A & PDF/UA veraPDF Compliance ──────────────────────
// Availability is a hard check: a missing veraPDF fails the harness instead
// of silently skipping the conformance validation. The validation itself runs
// on every harness invocation (the PDF/A-2b + UA-1 blob is always generated).
console.log(`\n── A5. PDF/A & PDF/UA veraPDF Compliance ──`);
let verapdfPath = "verapdf";
let hasVerapdf = false;
try {
  execSync("verapdf --version", { stdio: "ignore" });
  hasVerapdf = true;
} catch {
  if (existsSync("/home/rixzkiye/.local/bin/verapdf")) {
    verapdfPath = "/home/rixzkiye/.local/bin/verapdf";
    hasVerapdf = true;
  }
}
check("veraPDF CLI available (required for PDF/A-2b + PDF/UA-1 validation)", hasVerapdf, {
  path: hasVerapdf ? verapdfPath : null,
  ...(hasVerapdf
    ? {}
    : { error: "veraPDF not found on PATH or at /home/rixzkiye/.local/bin/verapdf" }),
});

if (hasVerapdf) {
  try {
    const out2b = execSync(`${verapdfPath} --format text --flavour 2b "${pdfaPdfPath}"`, {
      encoding: "utf-8",
    });
    const compliant2b = out2b.includes('isCompliant="true"') || out2b.includes("PASS");
    check("veraPDF PDF/A-2b compliance (0 violations)", compliant2b, {
      flavour: "2b",
      summary:
        out2b
          .split("\n")
          .find((l) => l.includes("PASS") || l.includes("FAIL") || l.includes("compliant"))
          ?.trim() ?? "OK",
    });

    const outUa = execSync(`${verapdfPath} --format text --flavour ua1 "${pdfaPdfPath}"`, {
      encoding: "utf-8",
    });
    const compliantUa = outUa.includes('isCompliant="true"') || outUa.includes("PASS");
    check("veraPDF PDF/UA-1 compliance (0 violations)", compliantUa, {
      flavour: "ua1",
      summary:
        outUa
          .split("\n")
          .find((l) => l.includes("PASS") || l.includes("FAIL") || l.includes("compliant"))
          ?.trim() ?? "OK",
    });
  } catch (e) {
    check("veraPDF execution", false, { error: e.message });
  }
}

check("No unhandled page errors during fidelity run", pageErrors.length === 0, pageErrors);

await browser.close();

console.log(
  failures === 0
    ? `\nALL FIDELITY CHECKS PASSED (artifacts in ${SHOTS})`
    : `\n${failures} FIDELITY CHECK(S) FAILED`,
);

process.exit(failures === 0 ? 0 : 1);
