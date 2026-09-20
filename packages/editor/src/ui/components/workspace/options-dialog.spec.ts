// @vitest-environment happy-dom
/**
 * Options → View: Word's "Show vertical ruler in Print Layout view" checkbox.
 * The dialog seeds it from the host property, labels it from the en/zh tables,
 * and carries the value on the options:ok commit.
 */
import { afterEach, describe, expect, it } from "vitest";

import "../../../document/i18n"; // registers the business tables (options.*)
import DocenOptionsDialog from "./options-dialog";

const created: DocenOptionsDialog[] = [];
afterEach(() => {
  for (const el of created) el.remove();
  created.length = 0;
});

async function settle(): Promise<void> {
  for (let i = 0; i < 3; i++) await new Promise((r) => setTimeout(r, 0));
}

async function mountDialog(): Promise<DocenOptionsDialog> {
  const dialog = new DocenOptionsDialog();
  created.push(dialog);
  document.body.append(dialog);
  await settle();
  // The modal shell would own show/hide; the unit test drives the seeding and
  // commit paths directly.
  (dialog as unknown as { dialogEl: { show(): void; hide(): void } }).dialogEl = {
    show: () => {},
    hide: () => {},
  };
  return dialog;
}

describe("Options dialog — vertical ruler setting", () => {
  it("labels the View section from the translation tables", async () => {
    const dialog = await mountDialog();
    expect(dialog.viewHeadingEl?.textContent).toBe("View");
    expect(dialog.verticalRulerLabelEl?.textContent).toBe(
      "Show vertical ruler in Print Layout view",
    );
  });

  it("pre-fills from the host property and defaults to on", async () => {
    const dialog = await mountDialog();
    dialog.showVerticalRuler = false;
    dialog.show();
    expect(dialog.verticalRulerBox?.checked).toBe(false);

    const fresh = new DocenOptionsDialog();
    created.push(fresh);
    document.body.append(fresh);
    await settle();
    (fresh as unknown as { dialogEl: { show(): void; hide(): void } }).dialogEl = {
      show: () => {},
      hide: () => {},
    };
    fresh.show();
    expect(fresh.verticalRulerBox?.checked).toBe(true);
  });

  it("carries showVerticalRuler on the options:ok commit", async () => {
    const dialog = await mountDialog();
    let detail: { showVerticalRuler?: boolean } | null = null;
    dialog.addEventListener("options:ok", (event) => {
      detail = (event as CustomEvent<{ showVerticalRuler?: boolean }>).detail;
    });

    dialog.showVerticalRuler = true;
    dialog.show();
    expect(dialog.verticalRulerBox?.checked).toBe(true);
    if (dialog.verticalRulerBox) dialog.verticalRulerBox.checked = false;
    dialog.onOk();
    expect(detail!.showVerticalRuler).toBe(false);
  });
});
