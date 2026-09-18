// @vitest-environment happy-dom
import { Document, Paragraph, Table, TableCell, TableRow } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import type { LaidOutTable } from "@docen/layout";
import { describe, expect, it, vi } from "vitest";

vi.mock("leafer-ui", () => {
  class StubElement {
    children: StubElement[] = [];
    x: number = 0;
    y: number = 0;
    points?: number[];
    stroke?: string;
    strokeWidth?: number;
    dashPattern?: number[];
    fill?: string;
    width?: number;
    height?: number;
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
    Group: StubElement,
    Image: make(),
    ImageManager: make(),
    Line: StubElement,
    Path: make(),
    Rect: StubElement,
    Resource: make(),
    Text: make(),
  };
});

import { drawDiagonal, paintTable } from "../../../../core/src/paint/table";
import { DocumentCommands } from "../extensions/commands";
import {
  ART_BORDER_PRESETS,
  getArtBorderSvgDataUri,
  isArtPreset,
  resolveArtPreset,
} from "./art-borders";

const { Group, Line } = (await import("leafer-ui")) as any;

const Text = TextNode.create({ name: "text", group: "inline" });

const EXTENSIONS = [Document, Paragraph, Text, Table, TableRow, TableCell, DocumentCommands];

function buildEditor(cellBorders?: Record<string, unknown>): EditorType {
  const editor = new Editor({
    element: null,
    extensions: EXTENSIONS,
    content: {
      type: "doc",
      content: [
        {
          type: "table",
          attrs: { columnWidths: [100, 100] },
          content: [
            {
              type: "tableRow",
              content: [
                {
                  type: "tableCell",
                  attrs: { borders: cellBorders ?? null },
                  content: [{ type: "paragraph", content: [{ type: "text", text: "A1" }] }],
                },
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", content: [{ type: "text", text: "B1" }] }],
                },
              ],
            },
          ],
        },
      ],
    },
  });
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
}

describe("W1.6 Art Borders", () => {
  it("defines 10 bespoke vector presets", () => {
    expect(ART_BORDER_PRESETS.length).toBe(10);
    const ids = ART_BORDER_PRESETS.map((p) => p.id);
    expect(ids).toContain("stars");
    expect(ids).toContain("hearts");
    expect(ids).toContain("apples");
    expect(ids).toContain("diamonds");
    expect(ids).toContain("decoArch");
    expect(ids).toContain("classic");
    expect(ids).toContain("zigzag");
    expect(ids).toContain("vines");
    expect(ids).toContain("dots");
    expect(ids).toContain("doubleWave");
  });

  it("resolves Word OOXML art tokens into known presets", () => {
    expect(resolveArtPreset("Stars")).toBe("stars");
    expect(resolveArtPreset("Heart")).toBe("hearts");
    expect(resolveArtPreset("Apples")).toBe("apples");
    expect(resolveArtPreset("Diamonds")).toBe("diamonds");
    expect(resolveArtPreset("ArtDeco")).toBe("decoArch");
    expect(resolveArtPreset("ZigZag")).toBe("zigzag");
    expect(resolveArtPreset("FloralVines")).toBe("vines");
    expect(resolveArtPreset("Pearls")).toBe("dots");
    expect(resolveArtPreset("DoubleWave")).toBe("doubleWave");
    expect(resolveArtPreset("NonExistentToken")).toBe("classic");
    expect(isArtPreset("apples")).toBe(true);
  });

  it("generates valid SVG data URI strings with given color", () => {
    const uri = getArtBorderSvgDataUri("hearts", "FF0000");
    expect(uri.startsWith("data:image/svg+xml;utf8,")).toBe(true);
    const decoded = decodeURIComponent(uri);
    expect(decoded).toContain("<svg");
    expect(decoded).toContain('fill="#FF0000"');
    expect(decoded).toContain("</svg>");
  });

  it("handles auto/default color fallback in getArtBorderSvgDataUri", () => {
    const uri = getArtBorderSvgDataUri("stars", "auto");
    const decoded = decodeURIComponent(uri);
    expect(decoded).toContain('fill="#2F5597"');
  });
});

describe("W1.6 Diagonal Borders Painter", () => {
  it("draws diagonal lines with drawDiagonal", () => {
    const tree = new Group();
    drawDiagonal(tree, 0, 0, 100, 50, {
      style: "dashed",
      px: 2,
      color: "FF0000",
    });

    expect(tree.children.length).toBe(1);
    const line = tree.children[0] as any;
    expect(line.x).toBe(0);
    expect(line.y).toBe(0);
    expect(line.stroke).toBe("#FF0000");
    expect(line.strokeWidth).toBe(2);
    expect(line.points).toEqual([0, 0, 100, 50]);
    expect(line.dashPattern).toEqual([4, 2]);
  });

  it("paints diagonal borders in paintTable for tl2br and tr2bl", () => {
    const tree = new Group();
    const mockTable: LaidOutTable = {
      kind: "table",
      widthPx: 100,
      heightPx: 50,
      columnWidthsPx: [100],
      rows: [
        {
          heightPx: 50,
          cells: [
            {
              colspan: 1,
              rowspan: 1,
              insets: { left: 5, top: 5, right: 5, bottom: 5 },
              innerWidthPx: 90,
              stack: [],
              borders: {
                tl2br: { style: "single", px: 1, color: "000000" },
                tr2bl: { style: "single", px: 1, color: "0000FF" },
              },
            },
          ],
        },
      ],
    };

    paintTable(tree, mockTable, 0, 0, {
      fontString: () => "12px sans-serif",
    } as any);

    const lines = (tree.children as any[]).filter((c: any) => c instanceof Line);
    expect(lines.length).toBeGreaterThanOrEqual(2);
    // Find tl2br: (0, 0) -> (100, 50)
    const tl2br = lines.find(
      (l: any) => l.x === 0 && l.y === 0 && l.points?.[2] === 100 && l.points?.[3] === 50,
    );
    expect(tl2br).toBeDefined();
    // Find tr2bl: (100, 0) -> (0, 50), points = [0, 0, -100, 50]
    const tr2bl = lines.find(
      (l: any) => l.x === 100 && l.y === 0 && l.points?.[2] === -100 && l.points?.[3] === 50,
    );
    expect(tr2bl).toBeDefined();
  });
});

describe("W1.6 Table Diagonal Borders Command", () => {
  it("toggles diagonalDown (tl2br) border on cell", () => {
    const editor = buildEditor();
    // Move selection into first cell (pos = 3)
    editor.commands.setTextSelection(3);

    // Apply diagonalDown
    editor.commands["table-borders"]("diagonalDown");

    let cellAttrs: any;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tableCell" && !cellAttrs) {
        cellAttrs = node.attrs;
      }
    });
    expect(cellAttrs?.borders?.tl2br).toBeDefined();
    expect(cellAttrs?.borders?.tl2br.style).toBe("single");

    // Toggle off
    editor.commands["table-borders"]("diagonalDown");
    cellAttrs = null;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tableCell" && !cellAttrs) {
        cellAttrs = node.attrs;
      }
    });
    expect(cellAttrs?.borders?.tl2br).toBeUndefined();
  });

  it("toggles diagonalUp (tr2bl) border on cell", () => {
    const editor = buildEditor();
    editor.commands.setTextSelection(3);

    // Apply diagonalUp
    editor.commands["table-borders"]("diagonalUp");

    let cellAttrs: any;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tableCell" && !cellAttrs) {
        cellAttrs = node.attrs;
      }
    });
    expect(cellAttrs?.borders?.tr2bl).toBeDefined();
    expect(cellAttrs?.borders?.tr2bl.style).toBe("single");

    // Toggle off
    editor.commands["table-borders"]("diagonalUp");
    cellAttrs = null;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "tableCell" && !cellAttrs) {
        cellAttrs = node.attrs;
      }
    });
    expect(cellAttrs?.borders?.tr2bl).toBeUndefined();
  });
});
