import type { ParagraphChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";
import { isPromotableVmlPict, parseColorToHex, parseWeightToEmu } from "./vml-promotion";

function roundTripChild(child: ParagraphChild) {
  const doc = {
    sections: [{ children: [{ paragraph: { children: [child] } }] }],
  };
  const json = resolveDocument(doc as any, docxExtensions);
  const compiled = compileDocument(json, docxExtensions);
  const paragraph = compiled.sections[0]!.children[0] as {
    paragraph: { children?: ParagraphChild[] };
  };
  return {
    json,
    node: (
      json.content?.[0] as
        | { content?: { type: string; attrs?: Record<string, unknown> }[] }
        | undefined
    )?.content?.[0],
    child: paragraph.paragraph.children?.[0],
  };
}

describe("VML pict non-textbox promotion to wpsShape", () => {
  describe("color and weight helpers", () => {
    it("parses colors to uppercase 6-character hex", () => {
      expect(parseColorToHex("#ff0000")).toBe("FF0000");
      expect(parseColorToHex("#f00")).toBe("FF0000");
      expect(parseColorToHex("4472c4")).toBe("4472C4");
      expect(parseColorToHex("red")).toBe("FF0000");
      expect(parseColorToHex("blue")).toBe("0000FF");
      expect(parseColorToHex("rgb(0, 128, 255)")).toBe("0080FF");
      expect(parseColorToHex(undefined)).toBeUndefined();
    });

    it("parses strokeweight to EMU", () => {
      expect(parseWeightToEmu("1pt")).toBe(12700);
      expect(parseWeightToEmu("2pt")).toBe(25400);
      expect(parseWeightToEmu(undefined)).toBe(12700);
    });
  });

  describe("isPromotableVmlPict", () => {
    it("returns true for non-textbox shapes", () => {
      expect(
        isPromotableVmlPict({
          children: [{ rect: { style: "width:100pt;height:50pt" } }],
        }),
      ).toBe(true);
      expect(
        isPromotableVmlPict({
          children: [{ oval: { style: "width:100pt;height:100pt" } }],
        }),
      ).toBe(true);
    });

    it("returns false for shapes with textbox or imagedata", () => {
      expect(
        isPromotableVmlPict({
          children: [{ shape: { textbox: { children: [] } } }],
        }),
      ).toBe(false);
      expect(
        isPromotableVmlPict({
          children: [{ shape: { imagedata: { src: "test.png" } } }],
        }),
      ).toBe(false);
      expect(isPromotableVmlPict({ children: [] })).toBe(false);
    });
  });

  describe("round-trip and layout mapping", () => {
    it("promotes VML rect to wpsShape with presetGeometry rect and roundtrips", () => {
      const vmlChild: ParagraphChild = {
        pict: {
          children: [
            {
              rect: {
                style: "position:absolute;left:20pt;top:30pt;width:120pt;height:60pt;z-index:3",
                fillcolor: "#ED7D31",
                strokecolor: "#4472C4",
                strokeweight: "1.5pt",
              },
            },
          ],
        },
      } as any;

      const { node, child } = roundTripChild(vmlChild);
      expect(node?.type).toBe("wpsShape");

      const wps = (node?.attrs as Record<string, any>)?.wpsShape;
      expect(wps?.presetGeometry).toBe("rect");
      // 120pt * 12700 = 1524000 EMU
      expect(wps?.transformation?.width).toBe(1524000);
      // 60pt * 12700 = 762000 EMU
      expect(wps?.transformation?.height).toBe(762000);
      expect(wps?.fill?.color).toBe("ED7D31");
      expect(wps?.outline?.color).toBe("4472C4");
      expect(wps?.floating?.horizontalPosition?.offset).toBe(254000); // 20pt
      expect(wps?.floating?.verticalPosition?.offset).toBe(381000); // 30pt
      expect(wps?.floating?.zIndex).toBe(3);

      // Verify lossless round-trip back to pict
      expect(child).toBeDefined();
      expect("pict" in (child as any)).toBe(true);
      const pict = (child as any).pict;
      expect(pict.children[0].rect).toBeDefined();
      expect(pict.children[0].rect.fillcolor).toBe("#ED7D31");
      expect(pict.children[0].rect.strokecolor).toBe("#4472C4");
    });

    it("promotes oval to ellipse presetGeometry", () => {
      const vmlChild: ParagraphChild = {
        pict: {
          children: [
            {
              oval: {
                style: "width:80pt;height:80pt",
                fillcolor: "green",
              },
            },
          ],
        },
      } as any;

      const { node, child } = roundTripChild(vmlChild);
      expect(node?.type).toBe("wpsShape");
      const wps = (node?.attrs as Record<string, any>)?.wpsShape;
      expect(wps?.presetGeometry).toBe("ellipse");
      expect(wps?.fill?.color).toBe("008000");

      const pict = (child as any).pict;
      expect(pict.children[0].oval).toBeDefined();
    });

    it("promotes polyline to customGeometry with lnTo commands", () => {
      const vmlChild: ParagraphChild = {
        pict: {
          children: [
            {
              polyline: {
                points: "0,0 50,100 100,50",
                style: "width:100pt;height:100pt",
                strokecolor: "black",
              },
            },
          ],
        },
      } as any;

      const { node, child } = roundTripChild(vmlChild);
      expect(node?.type).toBe("wpsShape");
      const wps = (node?.attrs as Record<string, any>)?.wpsShape;
      expect(wps?.customGeometry?.pathList).toBeDefined();
      const cmds = wps.customGeometry.pathList[0].commands;
      expect(cmds[0].command).toBe("moveTo");
      expect(cmds[1].command).toBe("lnTo");
      expect(cmds[2].command).toBe("lnTo");

      const pict = (child as any).pict;
      expect(pict.children[0].polyline).toBeDefined();
    });

    it("preserves updated geometry modifications on round-trip", () => {
      const vmlChild: ParagraphChild = {
        pict: {
          children: [
            {
              rect: {
                style: "position:absolute;left:10pt;top:10pt;width:50pt;height:50pt",
                fillcolor: "#FF0000",
              },
            },
          ],
        },
      } as any;

      const doc = {
        sections: [{ children: [{ paragraph: { children: [vmlChild] } }] }],
      };
      const json = resolveDocument(doc as any, docxExtensions);
      const wpsShapeNode = json.content?.[0]?.content?.[0] as any;
      // Mutate size (e.g. user resized on canvas to 150pt = 1905000 EMU)
      wpsShapeNode.attrs.wpsShape.transformation.width = 1905000;
      wpsShapeNode.attrs.wpsShape.fill.color = "0000FF";

      const compiled = compileDocument(json, docxExtensions);
      const compiledChild = (compiled.sections[0]!.children[0] as any).paragraph.children[0];
      const rect = compiledChild.pict.children[0].rect;

      expect(rect.style).toContain("width:150pt");
      expect(rect.fillcolor).toBe("#0000FF");
    });
  });
});
