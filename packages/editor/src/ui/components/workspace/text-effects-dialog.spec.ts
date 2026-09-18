// @vitest-environment happy-dom
// The custom Text Effects dialog: prefill from the selection's effects and the
// per-family patch it commits (unchecked sections clear, checked ones apply).
import { describe, expect, it } from "vitest";

import "./dialog";
import "./text-effects-dialog";
import type DocenTextEffectsDialog from "./text-effects-dialog";

async function settle(): Promise<void> {
  for (let i = 0; i < 5; i++) await new Promise((r) => setTimeout(r, 0));
}

const field = <T extends Element>(dialog: DocenTextEffectsDialog, name: string): T =>
  dialog.shadowRoot!.querySelector(`[data-field="${name}"]`) as T;

describe("<docen-text-effects-dialog>", () => {
  it("registers the element with show/hide/submit", () => {
    const dialog = document.createElement("docen-text-effects-dialog") as DocenTextEffectsDialog;
    document.body.appendChild(dialog);
    expect(typeof dialog.show).toBe("function");
    expect(typeof dialog.hide).toBe("function");
    expect(typeof dialog.submit).toBe("function");
    dialog.remove();
  });

  it("prefills every section from the current effects", async () => {
    const dialog = document.createElement("docen-text-effects-dialog") as DocenTextEffectsDialog;
    document.body.appendChild(dialog);
    await settle();

    dialog.show({
      outline: { color: "FF0000", widthPx: 4 / 3 },
      glow: { color: "00B050", opacity: 0.5, radiusPx: 8 },
      rotation: { x: 0, y: 10, z: 90 },
    });
    await settle();

    expect((field(dialog, "outline-on") as HTMLInputElement).checked).toBe(true);
    expect((field(dialog, "outline-color") as HTMLInputElement).value.toLowerCase()).toBe(
      "#ff0000",
    );
    expect(Number((field(dialog, "outline-width") as HTMLInputElement).value)).toBeCloseTo(1, 5);
    expect((field(dialog, "glow-on") as HTMLInputElement).checked).toBe(true);
    expect(Number((field(dialog, "glow-transparency") as HTMLInputElement).value)).toBeCloseTo(
      50,
      5,
    );
    expect((field(dialog, "rotation-on") as HTMLInputElement).checked).toBe(true);
    expect(Number((field(dialog, "rotation-z") as HTMLInputElement).value)).toBe(90);
    // Families the run does not carry stay off.
    expect((field(dialog, "shadow-on") as HTMLInputElement).checked).toBe(false);
    dialog.remove();
  });

  it("emits one patch entry per family — null for the unchecked sections", async () => {
    const dialog = document.createElement("docen-text-effects-dialog") as DocenTextEffectsDialog;
    document.body.appendChild(dialog);
    await settle();

    dialog.show({});
    await settle();
    let patch: Record<string, unknown> | null = null;
    dialog.addEventListener("text-effects:ok", (event) => {
      patch = (event as CustomEvent<Record<string, unknown>>).detail;
    });
    // Turn on an outline and a shadow.
    (field(dialog, "outline-on") as HTMLInputElement).checked = true;
    (field(dialog, "outline-color") as HTMLInputElement).value = "#c00000";
    (field(dialog, "outline-width") as HTMLInputElement).value = "2";
    (field(dialog, "shadow-on") as HTMLInputElement).checked = true;
    (field(dialog, "shadow-angle") as HTMLInputElement).value = "270";
    dialog.submit();
    await settle();

    expect(patch).not.toBeNull();
    expect(patch!.outline).toMatchObject({ color: "C00000" });
    expect((patch!.outline as { widthPx: number }).widthPx).toBeCloseTo(2.6667, 3);
    expect(patch!.shadow).toMatchObject({ dirDeg: 270 });
    expect(patch!.glow).toBeNull();
    expect(patch!.reflection).toBeNull();
    expect(patch!.bevel).toBeNull();
    expect(patch!.rotation).toBeNull();
    dialog.remove();
  });
});
