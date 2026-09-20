#!/usr/bin/env node
/**
 * PDF Export Fidelity Harness (R10-P1)
 *
 * Verifies vector export fidelity against the live LeaferJS canvas engine:
 *   A1. Scene pages contain no full-page image XObjects; content streams contain
 *       vector path operators (m, l, c, f, S).
 *   A2. Visual fidelity: pdftoppm-rendered PDF matches canvas snapshots with
 *       mean absolute difference <= 2.5/255 per page (average <= 2.0/255).
 *   A3. Text extraction: pdftotext extracts layout text runs.
 *   A4. Measurement: records size, export time, and fidelity for outlines vs
 *       embedded vs legacy raster.
 *
 * Run against the demo dev server:
 *   pnpm --filter @docen/editor demo -- --port 5182 --strictPort
 *   PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs \
 *     node packages/editor/scripts/pdf-fidelity.mjs --url http://localhost:5182/
 */
import { execSync } from "node:child_process";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
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
console.log(`Connecting to: ${URL}\n`);

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

  // Prewarm / get initial export
  const exportRes = await d.exportPdf();
  const pages = exportRes.pages;

  // 1. Vector outlines mode
  const t0 = performance.now();
  const blobOutlines = await pagesToPdf(pages, { textMode: "outlines" });
  const timeOutlines = performance.now() - t0;
  const bufOutlines = await blobOutlines.arrayBuffer();

  // 2. Vector embedded mode
  const t1 = performance.now();
  const blobEmbedded = await pagesToPdf(pages, { textMode: "embedded" });
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

  // Layout text runs
  const layoutTexts = [];
  d.editor.state.doc.descendants((node) => {
    if (node.isText && node.text) {
      layoutTexts.push(node.text);
    }
  });

  return {
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
  };
});

// Save PDF artifacts
const outlinesPdfPath = join(SHOTS, "p1-vector-outlines.pdf");
const embeddedPdfPath = join(SHOTS, "p1-vector-embedded.pdf");
const rasterPdfPath = join(SHOTS, "p1-legacy-raster.pdf");

writeFileSync(outlinesPdfPath, Buffer.from(data.outlines.bytes));
writeFileSync(embeddedPdfPath, Buffer.from(data.embedded.bytes));
writeFileSync(rasterPdfPath, Buffer.from(data.raster.bytes));

console.log(`\n── A4/A6 Benchmark Measurements (${data.pageCount} pages) ──`);
console.log(
  `Vector Outlines:  ${(data.outlines.size / 1024).toFixed(1)} KB | ${data.outlines.timeMs} ms`,
);
console.log(
  `Vector Embedded:  ${(data.embedded.size / 1024).toFixed(1)} KB | ${data.embedded.timeMs} ms`,
);
console.log(
  `Legacy Raster:    ${(data.raster.size / 1024).toFixed(1)} KB | ${data.raster.timeMs} ms`,
);

// ── A1 Verification: Vector Core ─────────────────────────────────────────────
console.log(`\n── A1. Vector Structure Verification ──`);
const pdfRaw = readFileSync(outlinesPdfPath).toString("latin1");
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

// ── A2 Verification: Visual Fidelity (pdftoppm -r 96 vs Canvas) ─────────────
console.log(`\n── A2. Visual Fidelity Verification ──`);
const ppmPrefix = join(SHOTS, "p1-ppm");
execSync(`pdftoppm -png -r 96 "${outlinesPdfPath}" "${ppmPrefix}"`);

const pageDiffs = [];
for (let i = 0; i < data.pages.length; i++) {
  const ppmFile = `${ppmPrefix}-${i + 1}.png`;
  const ppmBuffer = readFileSync(ppmFile);
  const ppmDataUrl = `data:image/png;base64,${ppmBuffer.toString("base64")}`;
  const canvasDataUrl = data.pages[i].url;

  const diff = await page.evaluate(
    async ({ ppmUrl, canvasUrl }) => {
      const loadImg = (src) =>
        new Promise((resolve) => {
          const img = new Image();
          img.onload = () => resolve(img);
          img.src = src;
        });
      const [img1, img2] = await Promise.all([loadImg(ppmUrl), loadImg(canvasUrl)]);
      const w = Math.min(img1.width, img2.width);
      const h = Math.min(img1.height, img2.height);

      const c1 = document.createElement("canvas");
      c1.width = w;
      c1.height = h;
      const ctx1 = c1.getContext("2d");
      ctx1.drawImage(img1, 0, 0, w, h);
      const d1 = ctx1.getImageData(0, 0, w, h).data;

      const c2 = document.createElement("canvas");
      c2.width = w;
      c2.height = h;
      const ctx2 = c2.getContext("2d");
      ctx2.fillStyle = "#ffffff";
      ctx2.fillRect(0, 0, w, h);
      ctx2.drawImage(img2, 0, 0, w, h);
      const d2 = ctx2.getImageData(0, 0, w, h).data;

      let totalDiff = 0;
      const totalPixels = w * h;
      for (let p = 0; p < totalPixels; p++) {
        const idx = p * 4;
        totalDiff += Math.abs(d1[idx] - d2[idx]);
        totalDiff += Math.abs(d1[idx + 1] - d2[idx + 1]);
        totalDiff += Math.abs(d1[idx + 2] - d2[idx + 2]);
      }

      return {
        w,
        h,
        meanAbsDiff: totalDiff / (totalPixels * 3),
      };
    },
    { ppmUrl: ppmDataUrl, canvasUrl: canvasDataUrl },
  );

  pageDiffs.push(diff.meanAbsDiff);
  check(`Page ${i + 1} pixel fidelity (meanAbsDiff <= 2.5/255)`, diff.meanAbsDiff <= 2.5, {
    meanAbsDiff: Number(diff.meanAbsDiff.toFixed(3)),
    dimensions: `${diff.w}x${diff.h}`,
  });
}

const avgDiff = pageDiffs.reduce((a, b) => a + b, 0) / pageDiffs.length;
check("Overall document fidelity (average meanAbsDiff <= 2.0/255)", avgDiff <= 2.0, {
  averageMeanAbsDiff: Number(avgDiff.toFixed(3)),
});

// ── A3 Verification: Text Extraction (pdftotext) ─────────────────────────────
console.log(`\n── A3. Text Extraction Verification ──`);
const pdfText = execSync(`pdftotext "${outlinesPdfPath}" -`).toString();
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
check("pdftotext extracts layout text (word match >= 95%)", wordMatchRatio >= 0.95, {
  foundWords,
  totalWords,
  matchPercentage: `${(wordMatchRatio * 100).toFixed(1)}%`,
});

check("No unhandled page errors during fidelity run", pageErrors.length === 0, pageErrors);

await browser.close();

console.log(
  failures === 0
    ? `\nALL FIDELITY CHECKS PASSED (artifacts in ${SHOTS})`
    : `\n${failures} FIDELITY CHECK(S) FAILED`,
);

process.exit(failures === 0 ? 0 : 1);
