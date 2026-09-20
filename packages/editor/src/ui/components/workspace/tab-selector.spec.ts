// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from "vitest";

import { DocenTabSelector, TAB_SELECTOR_TYPES } from "./tab-selector";

const created: DocenTabSelector[] = [];
afterEach(() => {
  for (const el of created) el.remove();
  created.length = 0;
});

async function mount(): Promise<DocenTabSelector> {
  const el = new DocenTabSelector();
  created.push(el);
  document.body.append(el);
  await new Promise((r) => setTimeout(r, 0));
  return el;
}

describe("DocenTabSelector (<docen-tab-selector>)", () => {
  it("defaults to left tab and cycles through Word tab types", async () => {
    const sel = await mount();
    expect(sel.activeType).toBe("left");
    expect(sel.titleText).toContain("Left Tab");

    const events: string[] = [];
    sel.addEventListener("tab-selector:change", (e: any) => events.push(e.detail.type));

    for (let i = 1; i < TAB_SELECTOR_TYPES.length; i++) {
      sel.cycle();
      expect(sel.activeType).toBe(TAB_SELECTOR_TYPES[i]);
      expect(events[events.length - 1]).toBe(TAB_SELECTOR_TYPES[i]);
    }

    // Wrap-around back to first
    sel.cycle();
    expect(sel.activeType).toBe("left");
    expect(events[events.length - 1]).toBe("left");
  });

  it("renders the SVG icon and title", async () => {
    const sel = await mount();
    const btn = sel.shadowRoot!.querySelector(".selector-btn") as HTMLButtonElement;
    expect(btn).not.toBeNull();
    const svg = sel.shadowRoot!.querySelector(".tab-icon");
    expect(svg).not.toBeNull();
  });
});
