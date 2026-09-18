import type { LayoutDrawingMember } from "@docen/layout";
import { describe, expect, it, vi } from "vitest";

import { paintChartMember } from "./chart";

vi.mock("leafer-ui", () => {
  class StubElement {
    children: StubElement[] = [];
    constructor(attrs: Record<string, unknown> = {}) {
      Object.assign(this, attrs);
    }
    add(...items: StubElement[]): void {
      this.children.push(...items);
    }
  }
  const make = () => class extends StubElement {};
  return {
    Box: make(),
    Ellipse: make(),
    Group: make(),
    Image: make(),
    ImageManager: make(),
    Line: make(),
    Path: make(),
    Rect: make(),
    Resource: make(),
    Text: make(),
  };
});

// Helper to recursively collect all elements in the painted tree
function collectElements(node: any): any[] {
  const result: any[] = [node];
  if (Array.isArray(node.children)) {
    for (const child of node.children) {
      result.push(...collectElements(child));
    }
  }
  return result;
}

const baseMember = (
  chart: Record<string, unknown>,
): Extract<LayoutDrawingMember, { kind: "chart" }> =>
  ({
    kind: "chart",
    x: 10,
    y: 10,
    width: 400,
    height: 300,
    chart,
  }) as unknown as Extract<LayoutDrawingMember, { kind: "chart" }>;

describe("paintChartMember", () => {
  it("paints placeholder for empty series or missing chart model", () => {
    const root = {
      children: [],
      add(...items: any[]) {
        this.children.push(...items);
      },
    } as any;
    paintChartMember(root, baseMember({}));
    const elements = collectElements(root);
    expect(elements.some((el) => el.fill === "#F3F3F3")).toBe(true);
  });

  it("paints surface chart with elevation bands and wireframe", () => {
    const surfaceChart = {
      type: "surface",
      categories: ["C1", "C2", "C3"],
      series: [
        { name: "S1", values: [10, 25, 40] },
        { name: "S2", values: [20, 35, 50] },
      ],
      wireframe: false,
    };
    const root = {
      children: [],
      add(...items: any[]) {
        this.children.push(...items);
      },
    } as any;
    paintChartMember(root, baseMember(surfaceChart));
    const elements = collectElements(root);
    // Should have filled paths for elevation bands
    expect(elements.some((el) => typeof el.path === "string" && el.fill)).toBe(true);

    // Wireframe mode has stroke and no fill
    const wireframeRoot = {
      children: [],
      add(...items: any[]) {
        this.children.push(...items);
      },
    } as any;
    paintChartMember(wireframeRoot, baseMember({ ...surfaceChart, wireframe: true }));
    const wireElements = collectElements(wireframeRoot);
    expect(wireElements.some((el) => typeof el.path === "string" && !el.fill && el.stroke)).toBe(
      true,
    );
  });

  it("paints ofPie chart with primary pie, secondary pie, and tangent connector lines", () => {
    const ofPieChart = {
      type: "ofPie",
      ofPieType: "pie",
      splitPosition: 2,
      categories: ["P1", "P2", "P3", "Sub1", "Sub2"],
      series: [{ name: "Sales", values: [40, 30, 20, 5, 5] }],
    };
    const root = {
      children: [],
      add(...items: any[]) {
        this.children.push(...items);
      },
    } as any;
    paintChartMember(root, baseMember(ofPieChart));
    const elements = collectElements(root);

    // Primary and secondary pie wedges have path strings
    const paths = elements.filter((el) => typeof el.path === "string");
    expect(paths.length).toBeGreaterThan(3);

    // Tangent connector lines between primary pie and secondary pie
    const connectors = elements.filter((el) => el.stroke === "#8C8C8C");
    expect(connectors.length).toBeGreaterThan(0);
  });

  it("paints ofPie chart with secondary bar", () => {
    const barOfPieChart = {
      type: "ofPie",
      ofPieType: "bar",
      splitPosition: 2,
      categories: ["P1", "P2", "Sub1", "Sub2"],
      series: [{ name: "Sales", values: [60, 30, 6, 4] }],
    };
    const root = {
      children: [],
      add(...items: any[]) {
        this.children.push(...items);
      },
    } as any;
    paintChartMember(root, baseMember(barOfPieChart));
    const elements = collectElements(root);
    expect(elements.length).toBeGreaterThan(5);
  });

  it("paints stock chart with Japanese candlesticks (OHLC)", () => {
    const stockChart = {
      type: "stock",
      categories: ["Mon", "Tue", "Wed"],
      series: [
        { name: "Open", values: [100, 105, 110] },
        { name: "High", values: [115, 112, 120] },
        { name: "Low", values: [95, 100, 102] },
        { name: "Close", values: [110, 102, 118] }, // Day 1 bullish, Day 2 bearish, Day 3 bullish
      ],
    };
    const root = {
      children: [],
      add(...items: any[]) {
        this.children.push(...items);
      },
    } as any;
    paintChartMember(root, baseMember(stockChart));
    const elements = collectElements(root);

    // Candle wicks and bodies
    const candleBodies = elements.filter(
      (el) =>
        (el.fill === "#FFFFFF" || el.fill === "#ED7D31") &&
        (el.stroke === "#26A69A" || el.stroke === "#595959") &&
        typeof el.width === "number" &&
        el.width > 0,
    );
    expect(candleBodies.length).toBe(3);
  });

  it("paints combo chart with mixed types and secondary axis", () => {
    const comboChart = {
      type: "combo",
      categories: ["Jan", "Feb", "Mar"],
      series: [
        { name: "Revenue", chartType: "column", values: [100, 120, 150] },
        { name: "Growth %", chartType: "line", axisGroup: "secondary", values: [10, 15, 20] },
      ],
      secondaryValueAxis: {
        numberFormat: "0%",
      },
    };
    const root = {
      children: [],
      add(...items: any[]) {
        this.children.push(...items);
      },
    } as any;
    paintChartMember(root, baseMember(comboChart));
    const elements = collectElements(root);

    // Both column rects and line path should exist
    expect(elements.some((el) => typeof el.path === "string")).toBe(true);
    // Secondary axis percentage tick labels
    const texts = elements.filter((el) => typeof el.text === "string");
    expect(texts.some((t) => t.text.includes("%"))).toBe(true);
  });

  it("renders data labels, trendlines, and error bars", () => {
    const columnChart = {
      type: "column",
      categories: ["A", "B", "C"],
      valueAxis: {
        min: 0,
        max: 100,
        numberFormat: "$#,##0",
      },
      dataLabels: {
        showValue: true,
        position: "outsideEnd",
      },
      series: [
        {
          name: "Sales",
          values: [20, 50, 80],
          trendlines: [{ type: "linear", displayEquation: true }],
          errorBars: { type: "percentage", val: 5, direction: "both" },
        },
      ],
    };
    const root = {
      children: [],
      add(...items: any[]) {
        this.children.push(...items);
      },
    } as any;
    paintChartMember(root, baseMember(columnChart));
    const elements = collectElements(root);
    const texts = elements.filter((el) => typeof el.text === "string");

    // Value formatted with currency
    expect(texts.some((t) => t.text.startsWith("$"))).toBe(true);
    // Trendline equation
    expect(texts.some((t) => t.text.includes("y ="))).toBe(true);
  });
});
