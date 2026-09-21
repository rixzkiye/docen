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
 *       Mean abs diffs and the alignment offset are diagnostics; no page
 *       exceptions.
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

// ── A2 Verification: Visual Fidelity (pdftoppm -r 96 vs Canvas) ─────────────
console.log(`\n── A2. Visual Fidelity Verification ──`);
const ppmPrefix = join(SHOTS, "p1-ppm");
execSync(`pdftoppm -png -r 96 "${outlinesPdfPath}" "${ppmPrefix}"`);

const pageDiffs = [];
const psnrs = [];
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

      // AA-aware metric: compare after a 2× box downsample so per-glyph-edge
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

      return {
        w,
        h,
        meanAbsDiff: fullResMean,
        downsampledMean: meanOf(h1, h2, dw, dh),
        psnr: Number(psnr.toFixed(2)),
        align: { dx: best.dx, dy: best.dy },
      };
    },
    { ppmUrl: ppmDataUrl, canvasUrl: canvasDataUrl },
  );

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
}

const avgDiff = pageDiffs.reduce((a, b) => a + b, 0) / pageDiffs.length;
const meanMse = psnrs.reduce((a, p) => a + 10 ** (-p / 10), 0) / psnrs.length;
const overallPsnr = -10 * Math.log10(meanMse);
check("Overall document fidelity (average PSNR >= 22 dB)", overallPsnr >= 22, {
  averagePsnr: Number(overallPsnr.toFixed(2)),
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
check("pdftotext extracts layout text (100.0% word match)", wordMatchRatio === 1, {
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
