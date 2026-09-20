// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from "vitest";

import { DocenVerticalRuler } from "./vertical-ruler";

const created: DocenVerticalRuler[] = [];
afterEach(() => {
  for (const el of created) el.style.display = "none";
  created.length = 0;
});

async function settle(): Promise<void> {
  for (let i = 0; i < 3; i++) await new Promise((r) => setTimeout(r, 0));
}

async function mountRuler(): Promise<DocenVerticalRuler> {
  const ruler = new DocenVerticalRuler();
  created.push(ruler);
  document.body.append(ruler);
  await settle();
  // happy-dom measures nothing; the host supplies the pane height at runtime.
  ruler.hostHeight = 600;
  return ruler;
}

function setGeometry(ruler: DocenVerticalRuler): void {
  ruler.unit = "in";
  ruler.scale = 1;
  ruler.originY = 100;
  ruler.pageHeightPx = 1056;
  ruler.contentTopPx = 96;
  ruler.contentHeightPx = 400;
}

describe.sequential("DocenVerticalRuler (<docen-vertical-ruler>) — Word's fixed vertical ruler", () => {
  it("renders the vertical tick hierarchy with 0 at the content origin", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler);
    ruler.renderTicks();

    const svg = ruler.shadowRoot!.querySelector(".ticks-svg") as SVGSVGElement;
    expect(svg.getAttribute("shape-rendering")).toBe("crispEdges");
    const ticks = Array.from(svg.querySelectorAll("line")).map((line) => ({
      y: Math.floor(Number(line.getAttribute("y1"))),
      len: Number(line.getAttribute("x1")) - Number(line.getAttribute("x2")),
    }));
    expect(ticks.length).toBeGreaterThan(0);
    // The shared four-level hierarchy: unit 7, half 5, quarter 3, eighth 2.
    expect(new Set(ticks.map((t) => t.len))).toEqual(new Set([7, 5, 3, 2]));
    // 0 sits at originY + contentTopPx = 196; the margin gutter above it is
    // labelled negatively.
    expect(svg.textContent).toContain("-1");
    expect(ticks.some((t) => t.y === 196 && t.len === 7)).toBe(true);
  });

  it("keeps ticking across the whole pane when the page scrolls up", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler);
    ruler.originY = -300; // page top far above the pane
    ruler.renderTicks();
    const svg = ruler.shadowRoot!.querySelector(".ticks-svg") as SVGSVGElement;
    const ys = Array.from(svg.querySelectorAll("line")).map((l) =>
      Math.floor(Number(l.getAttribute("y1"))),
    );
    expect(ys.length).toBeGreaterThan(20);
    expect(Math.min(...ys)).toBeLessThan(8); // reaches the pane top
    expect(Math.max(...ys)).toBeGreaterThan(590); // reaches the pane bottom
  });

  it("draws the text column between the margin boundaries", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler);
    await settle();
    const bg = ruler.shadowRoot!.querySelector(".content-bg") as HTMLElement;
    expect(bg.style.top).toBe("196px");
    expect(bg.style.height).toBe("400px");

    const handles = ruler.shadowRoot!.querySelectorAll(".margin-handle");
    expect(handles.length).toBe(2);
    const top = ruler.shadowRoot!.querySelector(".margin-handle.top") as HTMLElement;
    const bottom = ruler.shadowRoot!.querySelector(".margin-handle.bottom") as HTMLElement;
    expect(top.style.top).toBe("196px");
    expect(bottom.style.top).toBe("596px");
    expect(top.hasAttribute("hidden")).toBe(false);
    expect(bottom.hasAttribute("hidden")).toBe(false);
  });

  it("hides a margin handle that has scrolled out of the pane", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler);
    ruler.originY = -200; // content top above the strip
    await settle();
    await settle();
    const top = ruler.shadowRoot!.querySelector(".margin-handle.top") as HTMLElement;
    expect(top.hasAttribute("hidden")).toBe(true);
  });

  it("has no indent or tab-stop markers (Word's vertical ruler)", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler);
    expect(ruler.shadowRoot!.querySelector(".marker")).toBeNull();
    expect(ruler.shadowRoot!.querySelector(".tab-stop-item")).toBeNull();
  });

  it("drags the top margin boundary live and commits twips on release", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler); // top margin = 96px * 15 = 1440 twips

    const inputs: Array<{ side: string; twips: number }> = [];
    const commits: Array<{ side: string; twips: number }> = [];
    let started = 0;
    ruler.addEventListener("v-ruler:margin-start", () => {
      started++;
    });
    ruler.addEventListener("v-ruler:margin-input", (e: any) => inputs.push(e.detail));
    ruler.addEventListener("v-ruler:margin-commit", (e: any) => commits.push(e.detail));

    const handle = ruler.shadowRoot!.querySelector(".margin-handle.top") as HTMLElement;
    handle.setPointerCapture = () => {};
    ruler.onHandlePointerDown("top", {
      preventDefault: () => {},
      stopPropagation: () => {},
      clientX: 5,
      clientY: 196,
      pointerId: 1,
      target: handle,
    } as any);
    expect(started).toBe(1);
    expect(ruler.dragSide).toBe("top");

    window.dispatchEvent(new PointerEvent("pointermove", { clientY: 244 })); // +48px
    expect(ruler.dragValueTwips).toBe(2160);
    // The dragged handle tracks the pointer while the layout settles.
    expect(ruler.topHandleYPx).toBe(244);
    await new Promise((resolve) => setTimeout(resolve, 120));
    expect(inputs.at(-1)).toEqual({ side: "top", twips: 2160 });

    window.dispatchEvent(new PointerEvent("pointerup"));
    expect(commits.at(-1)).toEqual({ side: "top", twips: 2160 });
    expect(ruler.dragSide).toBeNull();
  });

  it("clamps the top margin between the page top and the content bottom", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler); // page 1056px, content 400px → max top 656px = 9840 twips

    const handle = ruler.shadowRoot!.querySelector(".margin-handle.top") as HTMLElement;
    handle.setPointerCapture = () => {};
    ruler.onHandlePointerDown("top", {
      preventDefault: () => {},
      stopPropagation: () => {},
      clientX: 5,
      clientY: 196,
      pointerId: 2,
      target: handle,
    } as any);

    window.dispatchEvent(new PointerEvent("pointermove", { clientY: -5000 }));
    expect(ruler.dragValueTwips).toBe(0);
    window.dispatchEvent(new PointerEvent("pointermove", { clientY: 5000 }));
    expect(ruler.dragValueTwips).toBe(9840);
    window.dispatchEvent(new PointerEvent("pointerup"));
    expect(ruler.dragValueTwips).toBeNull();
  });

  it("drags the bottom margin boundary the opposite way (larger margin moves up)", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler); // bottom margin = (1056 - 96 - 400)px = 560px = 8400 twips

    const commits: Array<{ side: string; twips: number }> = [];
    ruler.addEventListener("v-ruler:margin-commit", (e: any) => commits.push(e.detail));
    const handle = ruler.shadowRoot!.querySelector(".margin-handle.bottom") as HTMLElement;
    handle.setPointerCapture = () => {};
    ruler.onHandlePointerDown("bottom", {
      preventDefault: () => {},
      stopPropagation: () => {},
      clientX: 5,
      clientY: 596,
      pointerId: 3,
      target: handle,
    } as any);

    window.dispatchEvent(new PointerEvent("pointermove", { clientY: 572 })); // up 24px
    expect(ruler.dragValueTwips).toBe(8760);
    window.dispatchEvent(new PointerEvent("pointerup"));
    expect(commits.at(-1)).toEqual({ side: "bottom", twips: 8760 });
  });

  it("formats the drag tooltip in the shared measurement unit", async () => {
    const ruler = await mountRuler();
    setGeometry(ruler);
    expect(ruler.formatMeasurement(1440)).toBe("1.00 in");
    ruler.unit = "cm";
    expect(ruler.formatMeasurement(1440)).toBe("2.54 cm");

    const handle = ruler.shadowRoot!.querySelector(".margin-handle.top") as HTMLElement;
    handle.setPointerCapture = () => {};
    ruler.onHandlePointerDown("top", {
      preventDefault: () => {},
      stopPropagation: () => {},
      clientX: 5,
      clientY: 196,
      pointerId: 4,
      target: handle,
    } as any);
    expect(ruler.showTooltip).toBe(true);
    expect(ruler.tooltipText).toContain("cm");
    window.dispatchEvent(new PointerEvent("pointerup"));
    expect(ruler.showTooltip).toBe(false);
  });
});
