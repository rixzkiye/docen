// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

import "./chart-type-dialog";
import type DocenChartTypeDialog from "./chart-type-dialog";

async function settle(): Promise<void> {
  for (let i = 0; i < 5; i++) await new Promise((r) => setTimeout(r, 0));
}

describe("<docen-chart-type-dialog>", () => {
  it("registers docen-chart-type-dialog element and has show/hide methods", () => {
    const dialog = document.createElement("docen-chart-type-dialog") as DocenChartTypeDialog;
    document.body.appendChild(dialog);
    expect(typeof dialog.show).toBe("function");
    expect(typeof dialog.hide).toBe("function");
    dialog.remove();
  });

  it("shows dialog with active type preselected and emits chart-type:ok on confirm", async () => {
    const dialog = document.createElement("docen-chart-type-dialog") as DocenChartTypeDialog;
    document.body.appendChild(dialog);
    await settle();

    dialog.show("surface");
    await settle();
    expect(dialog.getAttribute("open")).toBe("");

    const shadow = dialog.shadowRoot!;
    const surfaceCard = shadow.querySelector('[data-type="surface"]');
    expect(surfaceCard).toBeTruthy();
    expect(surfaceCard?.classList.contains("selected")).toBe(true);

    // Click combo card
    const comboCard = shadow.querySelector('[data-type="combo"]') as HTMLElement;
    expect(comboCard).toBeTruthy();
    comboCard.click();
    await settle();
    expect(comboCard.classList.contains("selected")).toBe(true);

    // Listen for chart-type:ok event
    let confirmedType = "";
    dialog.addEventListener("chart-type:ok", (e: Event) => {
      const detail = (e as CustomEvent<{ type: string }>).detail;
      confirmedType = detail.type;
    });

    const okBtn = shadow.querySelector('[data-action="ok"]') as HTMLElement;
    expect(okBtn).toBeTruthy();
    okBtn.click();
    await settle();

    expect(confirmedType).toBe("combo");
    expect(dialog.hasAttribute("open")).toBe(false);

    dialog.remove();
  });

  it("closes on cancel without emitting chart-type:ok", async () => {
    const dialog = document.createElement("docen-chart-type-dialog") as DocenChartTypeDialog;
    document.body.appendChild(dialog);
    await settle();

    dialog.show("pie");
    await settle();

    let emitted = false;
    dialog.addEventListener("chart-type:ok", () => {
      emitted = true;
    });

    const shadow = dialog.shadowRoot!;
    const cancelBtn = shadow.querySelector('[data-action="cancel"]') as HTMLElement;
    cancelBtn.click();
    await settle();

    expect(emitted).toBe(false);
    expect(dialog.hasAttribute("open")).toBe(false);
    dialog.remove();
  });
});
