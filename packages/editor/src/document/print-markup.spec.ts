// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest";

import { DocenPrintPreview } from "../ui/components/workspace/print-preview";

describe("Print Markup Toggle (W7.2)", () => {
  it("DocenPrintPreview emits print-preview:markup-change when markup selection changes", () => {
    const preview = new DocenPrintPreview();
    const handler = vi.fn();
    preview.addEventListener("print-preview:markup-change", handler);

    preview.seed({
      snapshots: [{ width: 800, height: 1000, url: "data:image/png;base64,mock" }],
      printMarkup: false,
    });

    // Simulate dropdown change to markup
    (preview as any).markupDropdown = { value: "markup" };
    preview.onMarkupChange();

    expect(handler).toHaveBeenCalledTimes(1);
    expect(handler.mock.calls[0][0].detail).toEqual({ markup: true });

    // Simulate dropdown change back to document (clean)
    (preview as any).markupDropdown = { value: "document" };
    preview.onMarkupChange();

    expect(handler).toHaveBeenCalledTimes(2);
    expect(handler.mock.calls[1][0].detail).toEqual({ markup: false });
  });

  it("DocenPrintPreview accepts printMarkup initial state in seed()", () => {
    const preview = new DocenPrintPreview();
    preview.seed({
      snapshots: [{ width: 800, height: 1000, url: "data:image/png;base64,mock" }],
      printMarkup: true,
    });

    expect(preview.getSnapshots()).toHaveLength(1);
  });
});
