import type { RibbonTab } from "../../ui";
import { btn, col, grid, group, tabNode } from "./shared";

export const drawTab = (): RibbonTab =>
  tabNode("draw", [
    group("pens", [
      btn("pen", "draw-pen", { size: "large" }),
      col([
        grid([
          btn("pencil", "draw-pencil"),
          btn("highlight", "draw-highlighter"),
          btn("eraser", "draw-eraser"),
        ]),
      ]),
    ]),
    group("draw-tools", [
      btn("lasso", "lasso-select", { size: "large" }),
      col([
        grid([btn("board", "select-objects", { toggle: true }), btn("action-pen", "action-pen")]),
      ]),
    ]),
    group("ink-convert", [
      btn("ink-shape", "ink-to-shape", { size: "large" }),
      col([grid([btn("equation", "ink-to-math"), btn("sync", "replay-ink")])]),
    ]),
  ]);
