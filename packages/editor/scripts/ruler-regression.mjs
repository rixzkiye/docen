#!/usr/bin/env node
/**
 * Ruler runtime regression — Playwright checks that need a real layout engine,
 * a scrolling pane and pointer input (the vitest suite covers the pure
 * geometry, not the pane clipping).
 *
 * Run against the demo dev server:
 *   pnpm --filter @docen/editor demo -- --port 5180 --strictPort
 *   PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs \
 *     node packages/editor/scripts/ruler-regression.mjs --url http://localhost:5180/
 *
 * `PLAYWRIGHT_MODULE` reuses an existing Playwright install (the repo does not
 * depend on Playwright); without it, `playwright` must be resolvable. Checks:
 *   1. the horizontal band pins to the true top of the scrollport and nothing
 *      of the page is ever visible through it at any scroll offset,
 *   2. the band spans the pane and its scale tracks the centered page column,
 *   3. the vertical ruler is Print-Layout-only, option-gated, fixed, and shows
 *      the section of the page nearest the pane top,
 *   4. its top/bottom margin handles reflow the section live and commit twips
 *      (multi-section documents target the section at the pane top),
 *   5. R8 behavior stays green: indent drags, tab-stop click/drag/dblclick,
 *      bidi mirroring.
 *
 * Exits non-zero when any check fails. Screenshots land in `--shots` (default
 * /tmp/opencode) as r9-ruler-*.png.
 */
import { mkdirSync } from "node:fs";
import { join } from "node:path";

const args = process.argv.slice(2);
const argOf = (name, fallback) => {
  const i = args.indexOf(name);
  return i >= 0 && args[i + 1] ? args[i + 1] : fallback;
};
const URL = argOf("--url", "http://localhost:5180/");
const SHOTS = argOf("--shots", "/tmp/opencode");
const HEADLESS = !args.includes("--headed");
// A pre-installed browser can be pointed at explicitly (CHROME_PATH or
// --chrome) when the Playwright package's pinned revision is not installed.
const CHROME = argOf("--chrome", process.env.CHROME_PATH || undefined);

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || "playwright");

let failures = 0;
const check = (name, ok, detail) => {
  console.log(
    `${ok ? "PASS" : "FAIL"}  ${name}${detail !== undefined ? `  ${JSON.stringify(detail)}` : ""}`,
  );
  if (!ok) failures++;
};

const browser = await chromium.launch({
  headless: HEADLESS,
  ...(CHROME ? { executablePath: CHROME } : {}),
});
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
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
await page.evaluate(() => {
  const d = document.querySelector("docen-document");
  if (!d.getShowRuler()) {
    d.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: { event: "toggle-ruler" },
      }),
    );
  }
});
await page.waitForTimeout(400);

mkdirSync(SHOTS, { recursive: true });
const shot = async (name) => {
  await page.screenshot({ path: join(SHOTS, `r9-ruler-${name}.png`) });
};

// ── 1+2. Band pinning, clipping, width and page alignment ───────────────────
{
  const result = await page.evaluate(async () => {
    const d = document.querySelector("docen-document");
    const sr = d.shadowRoot;
    const area = sr.querySelector("docen-document-area");
    const ruler = sr.querySelector("docen-ruler");
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const pierce = (x, y) => {
      let el = document.elementFromPoint(x, y);
      while (el && el.shadowRoot) {
        const inner = el.shadowRoot.elementFromPoint?.(x, y);
        if (!inner || inner === el) break;
        el = inner;
      }
      return el;
    };
    const inRuler = (el) => {
      for (let cur = el; cur; cur = cur.parentElement ?? cur.getRootNode?.().host) {
        if (cur === ruler) return true;
        if (cur === document.documentElement) break;
      }
      return false;
    };
    const rows = [];
    const offsets = [0, 150, 400, 900, 1500, 2400];
    for (const st of offsets) {
      area.scrollTop = st;
      await wait(140);
      const areaRect = area.getBoundingClientRect();
      const rulerRect = ruler.getBoundingClientRect();
      const bandHits = [];
      for (const dy of [1, 5, 10, 15, 19]) {
        for (const fx of [0.01, 0.5, 0.99]) {
          const x = areaRect.left + areaRect.width * fx;
          const y = rulerRect.top + dy;
          const el = pierce(x, y);
          bandHits.push({ dy, fx, tag: el?.tagName?.toLowerCase(), inRuler: inRuler(el) });
        }
      }
      rows.push({
        scrollTop: st,
        pinned: Math.abs(rulerRect.bottom - areaRect.top) <= 1,
        bandHeight: Math.round(rulerRect.height),
        span: rulerRect.width >= areaRect.width - 1,
        pageLeftPx: ruler.pageLeftPx,
        zeroXPx: ruler.zeroXPx,
        marginLeftPx: ruler.marginLeftPx,
        scale: ruler.scale,
        misses: bandHits.filter((h) => !h.inRuler),
      });
    }
    area.scrollTop = 0;
    await wait(150);
    return rows;
  });
  for (const row of result) {
    check(`band pinned at scrollTop=${row.scrollTop}`, row.pinned, { top: "area/ruler" });
    check(
      `band clear of page content at scrollTop=${row.scrollTop}`,
      row.misses.length === 0,
      row.misses,
    );
    check(`band spans the pane at scrollTop=${row.scrollTop}`, row.span);
  }
  const first = result[0];
  check(
    "band scale is offset to the centered page column",
    Math.abs(first.zeroXPx - (first.pageLeftPx * first.scale + first.marginLeftPx * first.scale)) <
      1,
    { zeroXPx: first.zeroXPx, pageLeftPx: first.pageLeftPx },
  );
  await shot("band");
}

// ── 3. Vertical ruler: fixed, Print-only, option-gated, gutter-anchored ─────
{
  const matrix = [];
  for (const view of ["print", "web", "draft", "read", "outline"]) {
    await page.evaluate(
      (v) => document.querySelector("docen-document").setAttribute("view", v),
      view,
    );
    await page.waitForTimeout(300);
    const row = await page.evaluate(() => {
      const d = document.querySelector("docen-document");
      const sr = d.shadowRoot;
      const area = sr.querySelector("docen-document-area");
      const h = sr.querySelector("docen-ruler");
      const vr = sr.querySelector("docen-vertical-ruler");
      const frame = sr.querySelector(".canvas-pages")?.firstElementChild;
      const a = area.getBoundingClientRect();
      const hr = h.getBoundingClientRect();
      const v = vr.getBoundingClientRect();
      const fr = frame?.getBoundingClientRect();
      return {
        view: d.getAttribute("view") ?? "print",
        h: getComputedStyle(h).display,
        v: getComputedStyle(vr).display,
        fixedTop: vr.style.display === "none" ? null : Math.abs(v.top - a.top) <= 1,
        spansPane: vr.style.display === "none" ? null : Math.abs(v.height - a.height) <= 1,
        inGutter: vr.style.display === "none" ? null : Math.abs(v.right - a.left) <= 1,
        ticks: vr.shadowRoot.querySelectorAll("line").length,
        handles: vr.shadowRoot.querySelectorAll(".margin-handle").length,
        hasTopGap: fr ? fr.top >= hr.bottom + 16 : null,
      };
    });
    matrix.push(row);
  }
  await page.evaluate(() => document.querySelector("docen-document").setAttribute("view", "print"));
  await page.waitForTimeout(300);
  check(
    "horizontal ruler visible in print/web/draft",
    matrix.slice(0, 3).every((r) => r.h !== "none"),
    matrix.map((r) => [r.view, r.h]),
  );
  check(
    "horizontal ruler hidden in read/outline",
    matrix.slice(3).every((r) => r.h === "none"),
  );
  check(
    "vertical ruler Print-Layout-only",
    matrix[0].v !== "none" && matrix.slice(1).every((r) => r.v === "none"),
  );
  check("vertical ruler pinned alongside the document area", matrix[0].fixedTop === true);
  check("vertical ruler spans the pane height", matrix[0].spansPane === true);
  check("vertical ruler sits in the left window gutter", matrix[0].inGutter === true);
  check("workspace has top gap between horizontal ruler and page", matrix[0].hasTopGap === true);
  check(
    "vertical ruler renders the tick scale and two margin handles",
    matrix[0].ticks > 20 && matrix[0].handles === 2,
    { ticks: matrix[0].ticks },
  );

  await page.evaluate(() => document.querySelector("docen-document").setShowVerticalRuler(false));
  await page.waitForTimeout(200);
  const off = await page.evaluate(
    () =>
      getComputedStyle(
        document.querySelector("docen-document").shadowRoot.querySelector("docen-vertical-ruler"),
      ).display,
  );
  await page.evaluate(() => document.querySelector("docen-document").setShowVerticalRuler(true));
  await page.waitForTimeout(200);
  const on = await page.evaluate(
    () =>
      getComputedStyle(
        document.querySelector("docen-document").shadowRoot.querySelector("docen-vertical-ruler"),
      ).display,
  );
  check("vertical ruler option gates visibility", off === "none" && on !== "none", { off, on });
  await shot("vertical-ruler");
}

// ── 3a. Corner Tab Selector button (Word-parity) ────────────────────────────
{
  const tsResult = await page.evaluate(async () => {
    const d = document.querySelector("docen-document");
    const sr = d.shadowRoot;
    const ts = sr.querySelector("docen-tab-selector");
    const h = sr.querySelector("docen-ruler");
    const vr = sr.querySelector("docen-vertical-ruler");
    if (!ts) return { exists: false };
    const tsRect = ts.getBoundingClientRect();
    const hr = h.getBoundingClientRect();
    const v = vr.getBoundingClientRect();

    const initialType = ts.activeType;
    const cycledTypes = [];
    for (let i = 0; i < 7; i++) {
      ts.cycle();
      cycledTypes.push({
        tsType: ts.activeType,
        rulerType: h.activeTabType,
      });
    }

    return {
      exists: true,
      display: getComputedStyle(ts).display,
      width: Math.round(tsRect.width),
      height: Math.round(tsRect.height),
      alignedWithRulers:
        Math.abs(tsRect.right - hr.left) <= 1 &&
        Math.abs(tsRect.bottom - v.top) <= 1 &&
        Math.abs(tsRect.top - hr.top) <= 1 &&
        Math.abs(tsRect.left - v.left) <= 1,
      initialType,
      cycledTypes,
    };
  });

  check(
    "corner Tab Selector exists and is visible",
    tsResult.exists && tsResult.display !== "none",
  );
  check("corner Tab Selector is 20x20px", tsResult.width === 20 && tsResult.height === 20, {
    width: tsResult.width,
    height: tsResult.height,
  });
  check(
    "corner Tab Selector is aligned with horizontal and vertical rulers",
    tsResult.alignedWithRulers === true,
  );
  check(
    "corner Tab Selector cycles through 7 tab types and syncs with horizontal ruler",
    tsResult.cycledTypes?.every((c) => c.tsType === c.rulerType) &&
      tsResult.cycledTypes?.[tsResult.cycledTypes.length - 1]?.tsType === "left",
    tsResult.cycledTypes,
  );
}

// ── 3b. Options → View entry drives the vertical ruler option ───────────────
{
  const result = await page.evaluate(async () => {
    const d = document.querySelector("docen-document");
    d.dispatchEvent(
      new CustomEvent("command", { bubbles: true, composed: true, detail: { event: "options" } }),
    );
    await new Promise((r) => setTimeout(r, 350));
    const od = d.shadowRoot.querySelector("docen-options-dialog");
    const label = od.verticalRulerLabelEl?.textContent ?? null;
    const checked = od.verticalRulerBox?.checked ?? null;
    // Uncheck → OK: the host setting flips and the ruler hides.
    if (od.verticalRulerBox) od.verticalRulerBox.checked = false;
    od.onOk();
    await new Promise((r) => setTimeout(r, 250));
    const vr = d.shadowRoot.querySelector("docen-vertical-ruler");
    const afterOff = {
      attr: d.getAttribute("show-vertical-ruler"),
      display: getComputedStyle(vr).display,
    };
    d.setShowVerticalRuler(true);
    await new Promise((r) => setTimeout(r, 200));
    return { label, checked, afterOff, restored: d.getShowVerticalRuler() };
  });
  check(
    "Options → View shows the Word label",
    typeof result.label === "string" && /vertical ruler/i.test(result.label),
    { label: result.label },
  );
  check("Options → View seeds the current option", result.checked === true);
  check(
    "Options OK applies the vertical-ruler option",
    result.afterOff.attr === "false" && result.afterOff.display === "none",
    result.afterOff,
  );
  check("vertical-ruler option restores on", result.restored === true);
}

// ── 4. Margin handles drag the pane-top section live and commit twips ───────
{
  const margins = () =>
    page.evaluate(() => {
      const j = document.querySelector("docen-document").getJSON();
      const first = j.content.find((n) => n.attrs?.sectionProperties)?.attrs.sectionProperties;
      const body = j.attrs?.sectionProperties;
      const num = (v) => (typeof v === "number" ? v : null);
      return {
        firstTop: num(first?.pageMargin?.top),
        firstBottom: num(first?.pageMargin?.bottom),
        bodyTop: num(body?.pageMargin?.top),
        bodyBottom: num(body?.pageMargin?.bottom),
      };
    });
  // The demo leaves margins implicit (engine defaults): derive the twip
  // baseline from the rendered content origin so the delta is the oracle.
  const baseline = await page.evaluate(() => {
    const d = document.querySelector("docen-document");
    const vr = d.shadowRoot.querySelector("docen-vertical-ruler");
    return {
      top: Math.round(vr.contentTopPx * 15),
      height: Math.round(vr.contentHeightPx * 15),
      page: Math.round(vr.pageHeightPx * 15),
    };
  });
  const dragHandle = async (side, dy) => {
    const box = await page
      .locator(`docen-document docen-vertical-ruler .margin-handle.${side}`)
      .boundingBox();
    if (!box) return null;
    await page.mouse.move(box.x + 5, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 5, box.y + box.height / 2 + dy, { steps: 4 });
    await page.waitForTimeout(150);
    const live = await margins();
    return {
      live,
      box,
      release: async () => {
        await page.mouse.up();
        await page.waitForTimeout(250);
      },
    };
  };
  const before = await margins();
  const top = await dragHandle("top", 48);
  const liveTop = top?.live.firstTop;
  await top?.release();
  await page.waitForTimeout(200);
  const afterTop = await margins();
  // At 100% zoom 48px = 720 twips.
  check("top margin drags live", liveTop === baseline.top + 720, {
    liveTop,
    expected: baseline.top + 720,
  });
  check("top margin commits twips on release", afterTop.firstTop === baseline.top + 720, afterTop);
  check(
    "the other sides survive the drag",
    afterTop.firstBottom === before.firstBottom && afterTop.bodyTop === before.bodyTop,
  );
  // Drag back to the implicit default.
  const back = await dragHandle("top", -48);
  await back?.release();
  await page.waitForTimeout(250);

  // Bottom handle: the pane must show the content-bottom boundary.
  await page.evaluate(() => {
    const area = document
      .querySelector("docen-document")
      .shadowRoot.querySelector("docen-document-area");
    area.scrollTop = 700;
  });
  await page.waitForTimeout(350);
  const bottomBase = await page.evaluate(() => {
    const vr = document
      .querySelector("docen-document")
      .shadowRoot.querySelector("docen-vertical-ruler");
    return Math.round((vr.pageHeightPx - vr.contentTopPx - vr.contentHeightPx) * 15);
  });
  const bottom = await dragHandle("bottom", -24);
  const liveBottom = bottom?.live.firstBottom;
  await bottom?.release();
  await page.waitForTimeout(250);
  const bottomAfter = await margins();
  check("bottom margin drags live (up grows the margin)", liveBottom === bottomBase + 360, {
    liveBottom,
    expected: bottomBase + 360,
  });
  check(
    "bottom margin commits twips on release",
    bottomAfter.firstBottom === bottomBase + 360,
    bottomAfter,
  );
  // Drag back.
  const bottomBack = await dragHandle("bottom", 24);
  await bottomBack?.release();
  await page.waitForTimeout(250);

  // Multi-section: at the last page the body-level sectPr is the target.
  await page.evaluate(() => {
    const area = document
      .querySelector("docen-document")
      .shadowRoot.querySelector("docen-document-area");
    area.scrollTop = area.scrollHeight;
  });
  await page.waitForTimeout(400);
  const bodyBefore = await margins();
  const bodyBase = await page.evaluate(() => {
    const vr = document
      .querySelector("docen-document")
      .shadowRoot.querySelector("docen-vertical-ruler");
    return Math.round((vr.pageHeightPx - vr.contentTopPx - vr.contentHeightPx) * 15);
  });
  const bodyDrag = await dragHandle("bottom", -20);
  await bodyDrag?.release();
  const bodyAfter = await margins();
  check(
    "a multi-section drag targets the pane-top section (body sectPr)",
    bodyDrag != null &&
      bodyAfter.bodyBottom === bodyBase + 300 &&
      bodyAfter.firstBottom === bodyBefore.firstBottom,
    { after: bodyAfter, base: bodyBase },
  );
  // Restore for later checks.
  await page.evaluate(() => {
    const area = document
      .querySelector("docen-document")
      .shadowRoot.querySelector("docen-document-area");
    area.scrollTop = 0;
  });
  await page.waitForTimeout(300);
  await shot("margin-drag");
}

// ── 5. R8 behavior re-verification ─────────────────────────────────────────
{
  // First-line indent marker drag.
  await page.evaluate(() =>
    document.querySelector("docen-document").editor.commands.setTextSelection(2),
  );
  await page.waitForTimeout(200);
  const marker = await page.evaluate(() => {
    const m = document
      .querySelector("docen-document")
      .shadowRoot.querySelector("docen-ruler")
      .shadowRoot.querySelector(".first-line-marker");
    const r = m.getBoundingClientRect();
    return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
  });
  await page.mouse.move(marker.x, marker.y);
  await page.mouse.down();
  await page.mouse.move(marker.x + 24, marker.y, { steps: 3 });
  await page.waitForTimeout(150);
  const dragLive = await page.evaluate(() => {
    const d = document.querySelector("docen-document");
    return {
      ruler: d.shadowRoot.querySelector("docen-ruler").firstLineTwips,
      para: d.editor.state.doc.child(0).attrs.indent?.firstLine ?? null,
    };
  });
  await page.mouse.up();
  await page.waitForTimeout(200);
  check(
    "indent marker drag updates the paragraph live",
    dragLive.ruler === 360 && dragLive.para === 360,
    dragLive,
  );

  // Tab stop: click adds a stop at the clicked offset (band-relative click).
  const tabAdd = await page.evaluate(() => {
    const d = document.querySelector("docen-document");
    const ruler = d.shadowRoot.querySelector("docen-ruler");
    const track = ruler.shadowRoot.querySelector(".ruler-track");
    const r = track.getBoundingClientRect();
    track.dispatchEvent(
      new MouseEvent("click", {
        bubbles: true,
        composed: true,
        clientX: r.left + ruler.zeroXPx + 120,
        clientY: r.top + 10,
      }),
    );
    return {
      stops: ruler.tabStops.length,
      pos: ruler.tabStops[0]?.position ?? null,
      paraStops: d.editor.state.doc.child(0).attrs.tabStops?.length ?? 0,
    };
  });
  check(
    "click adds a tab stop to the ruler and paragraph",
    tabAdd.stops === 1 && tabAdd.pos === 1800 && tabAdd.paraStops === 1,
    tabAdd,
  );

  // The stop glyph renders behind FAST's rAF batch — wait for the row.
  await page.waitForFunction(
    () =>
      !!document
        .querySelector("docen-document")
        .shadowRoot.querySelector("docen-ruler")
        .shadowRoot.querySelector(".tab-stop-item"),
    { timeout: 3000 },
  );
  await page.waitForTimeout(80);

  // Drag the stop 24px right.
  const stop = await page.evaluate(() => {
    const item = document
      .querySelector("docen-document")
      .shadowRoot.querySelector("docen-ruler")
      .shadowRoot.querySelector(".tab-stop-item");
    const r = item.getBoundingClientRect();
    return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
  });
  await page.mouse.move(stop.x, stop.y);
  await page.mouse.down();
  await page.mouse.move(stop.x + 24, stop.y, { steps: 3 });
  await page.mouse.up();
  await page.waitForTimeout(250);
  const tabMoved = await page.evaluate(() => {
    const d = document.querySelector("docen-document");
    return d.editor.state.doc.child(0).attrs.tabStops?.[0]?.position ?? null;
  });
  check("drag moves the tab stop (committed to the paragraph)", tabMoved === 2160, { tabMoved });

  // Double-click the strip opens the Tabs dialog.
  await page.evaluate(() => {
    const d = document.querySelector("docen-document");
    const ruler = d.shadowRoot.querySelector("docen-ruler");
    const track = ruler.shadowRoot.querySelector(".ruler-track");
    const r = track.getBoundingClientRect();
    track.dispatchEvent(
      new MouseEvent("dblclick", {
        bubbles: true,
        composed: true,
        clientX: r.left + ruler.zeroXPx + 300,
        clientY: r.top + 10,
      }),
    );
  });
  await page.waitForTimeout(350);
  const tabsOpen = await page.evaluate(() => {
    const tabs = document
      .querySelector("docen-document")
      .shadowRoot.querySelector("docen-tabs-dialog");
    return tabs?.shadowRoot?.querySelector("docen-dialog")?.hasAttribute("open") ?? false;
  });
  check("double-clicking the ruler opens the Tabs dialog", tabsOpen === true);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(250);

  // RTL mirroring.
  const rtl = await page.evaluate(async () => {
    const d = document.querySelector("docen-document");
    const ed = d.editor;
    const { tr } = ed.state;
    const node = ed.state.doc.child(0);
    tr.setNodeMarkup(0, undefined, { ...node.attrs, bidirectional: true });
    ed.view.dispatch(tr);
    await new Promise((r) => setTimeout(r, 300));
    const ruler = d.shadowRoot.querySelector("docen-ruler");
    return {
      dir: ruler.getAttribute("dir"),
      zero: ruler.zeroXPx,
      unitLabel: ruler.unitLabelX,
      width: ruler.pageWidthPx,
    };
  });
  check("ruler mirrors for a bidi paragraph", rtl.dir === "rtl" && rtl.unitLabel > rtl.zero, rtl);
  await shot("rtl");
  await page.evaluate(async () => {
    const d = document.querySelector("docen-document");
    const ed = d.editor;
    const { tr } = ed.state;
    const node = ed.state.doc.child(0);
    tr.setNodeMarkup(0, undefined, { ...node.attrs, bidirectional: false });
    ed.view.dispatch(tr);
    await new Promise((r) => setTimeout(r, 250));
  });
}

// ── Console cleanliness ─────────────────────────────────────────────────────
check("no page errors during the run", pageErrors.length === 0, pageErrors.slice(0, 3));

await browser.close();
console.log(
  failures === 0
    ? `\nALL CHECKS PASSED (screenshots in ${SHOTS})`
    : `\n${failures} CHECK(S) FAILED`,
);
process.exit(failures === 0 ? 0 : 1);
