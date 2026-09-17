// @vitest-environment happy-dom
// Navigation pane Pages tab: thumbnails render when provided and the page
// click handler emits the select event the host routes to the page jump.
import { afterEach, describe, expect, it } from "vitest";

import { DocenNavPages } from "./nav-pages";

const PNG = "data:image/png;base64,AAAA";

const created: DocenNavPages[] = [];
afterEach(() => {
  while (created.length) created.pop()!.remove();
});

/** Let FAST's template + repeat directives settle. */
async function settle(): Promise<void> {
  for (let i = 0; i < 3; i++) await new Promise((r) => setTimeout(r, 0));
}

async function mount(): Promise<DocenNavPages> {
  const pane = new DocenNavPages();
  created.push(pane);
  document.body.append(pane);
  return pane;
}

describe("DocenNavPages thumbnails", () => {
  it("renders a raster thumbnail per page and a placeholder when absent", async () => {
    const el = await mount();
    el.setPageCount(3, 1, [PNG, PNG, null]);
    await settle();
    const cards = el.shadowRoot!.querySelectorAll(".page-card");
    expect(cards).toHaveLength(3);
    const imgs = el.shadowRoot!.querySelectorAll<HTMLImageElement>(".thumb-img");
    expect(imgs).toHaveLength(2);
    expect(imgs[0]!.getAttribute("src")).toBe(PNG);
    // The third page falls back to the CSS placeholder lines.
    expect(cards[2]!.querySelectorAll(".thumb-line").length).toBeGreaterThan(0);
  });

  it("selects the clicked page and emits nav-pages:select", async () => {
    const el = await mount();
    el.setPageCount(3, 1, [PNG, PNG, PNG]);
    await settle();
    const emitted: Array<{ type: string; detail: { page: number } }> = [];
    (el as unknown as { $emit: (type: string, detail: { page: number }) => void }).$emit = (
      type,
      detail,
    ) => {
      emitted.push({ type, detail });
    };

    el.onPageClick(2);
    expect(el.activePage).toBe(2);
    expect(el.pages.filter((p) => p.active).map((p) => p.pageNumber)).toEqual([2]);
    expect(emitted).toEqual([{ type: "nav-pages:select", detail: { page: 2 } }]);
  });
});
