import type { RibbonTab } from "../../ui";
import { btn, col, grid, group, opt, parsedItems, split, tabNode } from "./shared";

export const zoomItems = (): string =>
  JSON.stringify([
    { text: opt("200"), value: "200" },
    { text: opt("100"), value: "100" },
    { text: opt("75"), value: "75" },
    { text: opt("50"), value: "50" },
    { text: opt("page-width"), value: "page-width" },
    { text: opt("text-width"), value: "text-width" },
    { text: opt("fit-page"), value: "fit-page" },
    { text: opt("zoom-dialog"), value: "zoom-dialog" },
  ]);

export const viewTab = (): RibbonTab =>
  tabNode("view", [
    group("views", [
      btn("document-print", "read-mode", { size: "large" }),
      btn("eye", "focus-mode", { size: "large" }),
      btn("print", "print-layout", { size: "large" }),
      btn("document-print", "web-layout", { size: "large" }),
      btn("group-objects", "outline", { size: "large" }),
      btn("document-print", "draft", { size: "large" }),
    ]),
    group("show", [
      col([
        btn("ruler", "toggle-ruler"),
        btn("gridlines", "toggle-gridlines"),
        btn("replace", "toggle-navigation"),
      ]),
    ]),
    group("zoom", [
      btn("zoom-in", "zoom", { size: "large" }),
      split("zoom-in", "zoom-100", parsedItems(zoomItems()), { size: "large" }),
    ]),
    group("window", [
      btn("grid", "new-window", { size: "large" }),
      col([grid([btn("grid", "arrange-all"), btn("group-objects", "split-window")])]),
    ]),
    group("macros", [
      btn("group-objects", "view-macros", { size: "large" }),
      btn("edit", "record-macro", { size: "large" }),
    ]),
  ]);
