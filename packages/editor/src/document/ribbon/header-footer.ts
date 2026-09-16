import type { RibbonTab } from "../../ui";
import { btn, col, group, opt, tab } from "./shared";

/** Word's Header & Footer Tools — the contextual tab while a header/footer
 *  story is being edited. Marked `contextual` so {@link ribbonTabs} excludes
 *  it from the static render; the host appends it when a story opens and
 *  removes it when the story closes (same lifecycle as
 *  {@link tableContextTabs}). The Options checkboxes mirror the Insert tab's
 *  header/footer drop-down flags (first-page-different / odd-even-different). */
export function headerFooterContextTab(): RibbonTab {
  return {
    id: "header-footer-tab",
    label: tab("header-footer-tools"),
    contextual: true,
    groups: [
      group("navigation", [
        btn("header", "goto-header", { size: "large" }),
        btn("footer", "goto-footer", { size: "large" }),
      ]),
      group("options", [
        col([
          {
            type: "checkbox",
            event: "header-option",
            value: "title-page",
            label: opt("different-first"),
          },
          {
            type: "checkbox",
            event: "header-option",
            value: "odd-even",
            label: opt("odd-even"),
          },
        ]),
      ]),
      group("close", [btn("close", "close-header-footer", { size: "large" })]),
    ],
  };
}
