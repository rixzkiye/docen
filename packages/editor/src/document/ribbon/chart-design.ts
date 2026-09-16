import type { RibbonTab } from "../../ui";
import {
  accessibilityGroup,
  arrangeGroup,
  btn,
  col,
  grid,
  group,
  menu,
  opt,
  parsedItems,
  sizeGroup,
  tab,
} from "./shared";

/** Chart Type's drop-down — every ChartType token in Word's order; the seven
 *  the renderer draws are live, the placeholder-only tail greys out. */
export const chartTypeItems = (): string =>
  JSON.stringify([
    { text: opt("chart-type.column"), value: "column" },
    { text: opt("chart-type.bar"), value: "bar" },
    { text: opt("chart-type.line"), value: "line" },
    { text: opt("chart-type.area"), value: "area" },
    { text: "-" },
    { text: opt("chart-type.pie"), value: "pie" },
    { text: opt("chart-type.doughnut"), value: "doughnut" },
    { text: opt("chart-type.scatter"), value: "scatter" },
    { text: "-" },
    { text: opt("chart-type.radar"), value: "radar" },
    { text: opt("chart-type.stock"), value: "stock" },
    { text: opt("chart-type.surface"), value: "surface", disabled: true },
    { text: opt("chart-type.of-pie"), value: "ofPie", disabled: true },
    { text: opt("chart-type.bubble"), value: "bubble" },
  ]);

/** The Legend placement drop-down — "none" plus LegendPosition's four sides. */
export const chartLegendItems = (): string =>
  JSON.stringify([
    { text: opt("chart-legend.none"), value: "none" },
    { text: opt("chart-legend.bottom"), value: "bottom" },
    { text: opt("chart-legend.left"), value: "left" },
    { text: opt("chart-legend.right"), value: "right" },
    { text: opt("chart-legend.top"), value: "top" },
  ]);

/** Word's Chart Tools — the contextual tab while a chart carries the
 *  selection. Word splits it into Design + Format tabs; the renderer's model
 *  (type/legend/data) fits one: the Type drop-down (placeholder-only types
 *  greyed), the Legend placement drop-down, and Edit Data (the grid dialog
 *  standing in for Word's embedded worksheet). Arrange/Size/Accessibility
 *  repeat the drawing tabs' shared groups. Marked `contextual` like the
 *  picture tab. */
export function chartDesignTab(): RibbonTab {
  return {
    id: "chart-design",
    label: tab("chart-design"),
    contextual: true,
    groups: [
      group("chart-type", [
        col([
          grid([menu("chart", "chart-type", parsedItems(chartTypeItems()), { size: "large" })]),
        ]),
      ]),
      group("chart-layouts", [
        col([
          grid([
            menu("data-area", "chart-legend", parsedItems(chartLegendItems()), { size: "large" }),
          ]),
        ]),
      ]),
      group("chart-data", [btn("table", "chart-edit-data", { size: "large" })]),
      accessibilityGroup(),
      arrangeGroup(),
      sizeGroup("chart-size", false),
    ],
  };
}
