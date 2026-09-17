// @vitest-environment happy-dom
// Reveal Formatting pane: the compare-to-selection diff is computed and
// rendered from the two snapshots the host feeds.
import { afterEach, describe, expect, it } from "vitest";

import { DocenRevealFormattingPane, type FormattingInfo } from "./reveal-formatting-pane";

/** Let FAST render before asserting on the shadow DOM. */
const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

const created: DocenRevealFormattingPane[] = [];
afterEach(() => {
  while (created.length) created.pop()!.remove();
});

const reference: FormattingInfo = {
  font: { family: "Calibri", size: "11 pt", bold: false, color: "Auto" },
  paragraph: { alignment: "Left", indentLeft: "0 pt" },
};
const current: FormattingInfo = {
  font: { family: "Calibri", size: "14 pt", bold: true, color: "Auto" },
  paragraph: { alignment: "Left", indentLeft: "0 pt" },
};

async function mount(): Promise<DocenRevealFormattingPane> {
  const pane = new DocenRevealFormattingPane();
  created.push(pane);
  document.body.append(pane);
  await tick();
  return pane;
}

describe("DocenRevealFormattingPane compare-to-selection", () => {
  it("diffs the current selection against the reference snapshot", async () => {
    const pane = await mount();
    pane.setComparison(reference);
    pane.setFormatting(current);

    const rows = pane.diffRows();
    expect(rows.some((row) => row.current === "14 pt")).toBe(true);
    expect(rows.some((row) => row.label === "reveal.fontSize")).toBe(true);
    // The compared section is rendered.
    expect(pane.shadowRoot!.textContent).toContain("reveal.comparedTo");
    expect(pane.shadowRoot!.textContent).toContain("11 pt → 14 pt");
  });

  it("shows no-differences when the snapshots match, and clears on disable", async () => {
    const pane = await mount();
    pane.setComparison({ ...current });
    pane.setFormatting(current);
    expect(pane.diffRows()).toHaveLength(0);
    expect(pane.shadowRoot!.textContent).toContain("reveal.noDifferences");

    pane.setComparison(null);
    expect(pane.shadowRoot!.textContent).not.toContain("reveal.comparedTo");
  });
});
