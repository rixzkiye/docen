import { describe, expect, it } from "vitest";

import {
  ptToTwip,
  twipToPt,
  emuToTwip,
  twipToEmu,
  parseLengthToTwips,
  parseVmlShapeLayout,
  stringifyVmlShapeLayout,
  normalizeVmlShapeStyle,
  compileDocument,
  resolveDocument,
  docxExtensions,
} from "../index";

describe("W9.3 Drawing and VML Shape Layout Modeling", () => {
  describe("Unit conversion helpers", () => {
    it("converts between pt and twip", () => {
      expect(ptToTwip(10)).toBe(200);
      expect(twipToPt(200)).toBe(10);
      expect(ptToTwip(72)).toBe(1440);
    });

    it("converts between emu and twip", () => {
      expect(emuToTwip(914400)).toBe(1440);
      expect(twipToEmu(1440)).toBe(914400);
    });

    it("parses different length units to twips", () => {
      expect(parseLengthToTwips("200pt")).toBe(4000);
      expect(parseLengthToTwips("1in")).toBe(1440);
      expect(parseLengthToTwips(100, "pt")).toBe(2000);
      expect(parseLengthToTwips(100, "twip")).toBe(100);
      expect(parseLengthToTwips("914400emu")).toBe(1440);
      expect(parseLengthToTwips(undefined)).toBeUndefined();
    });
  });

  describe("VML Style Parsing and Stringification", () => {
    it("normalizes CSS-like style strings", () => {
      const style = "position:absolute; width: 200pt; height: 100pt; z-index: 5";
      const normalized = normalizeVmlShapeStyle(style);
      expect(normalized).toEqual({
        position: "absolute",
        width: "200pt",
        height: "100pt",
        "z-index": 5,
      });
    });

    it("parses raw VML style string into structured DrawingShapeLayout", () => {
      const raw =
        "position:absolute;left:50pt;top:30pt;width:200pt;height:100pt;z-index:10;rotation:45;mso-wrap-style:square";
      const layout = parseVmlShapeLayout(raw);
      expect(layout).toBeDefined();
      expect(layout?.position).toBe("absolute");
      expect(layout?.left).toBe(1000);
      expect(layout?.top).toBe(600);
      expect(layout?.width).toBe(4000);
      expect(layout?.height).toBe(2000);
      expect(layout?.zIndex).toBe(10);
      expect(layout?.rotation).toBe(45);
      expect(layout?.wrapping).toBe("square");
    });

    it("stringifies structured DrawingShapeLayout back to VML style string", () => {
      const layout = {
        position: "absolute" as const,
        left: 1000,
        top: 600,
        width: 4000,
        height: 2000,
        zIndex: 10,
        rotation: 45,
        wrapping: "square" as const,
      };
      const stringified = stringifyVmlShapeLayout(layout);
      expect(stringified).toContain("position:absolute");
      expect(stringified).toContain("left:50pt");
      expect(stringified).toContain("top:30pt");
      expect(stringified).toContain("width:200pt");
      expect(stringified).toContain("height:100pt");
      expect(stringified).toContain("z-index:10");
      expect(stringified).toContain("rotation:45");
      expect(stringified).toContain("mso-wrap-style:square");
    });
  });

  describe("Textbox structured layout integration", () => {
    it("extracts layout attribute during document resolution and preserves on compile", () => {
      const doc = {
        sections: [
          {
            children: [
              {
                textbox: {
                  style: "position:absolute;width:150pt;height:80pt",
                  children: [{ paragraph: { text: "Inside text box" } }],
                },
              },
            ],
          },
        ],
      };

      const resolved = resolveDocument(doc as any, docxExtensions);
      const boxNode = resolved.content?.[0];
      expect(boxNode?.type).toBe("textbox");
      expect(boxNode?.attrs?.layout).toBeDefined();
      expect(boxNode?.attrs?.layout.width).toBe(3000); // 150pt * 20
      expect(boxNode?.attrs?.layout.height).toBe(1600); // 80pt * 20

      const compiled = compileDocument(resolved, docxExtensions);
      const compiledBox = (compiled.sections[0]!.children[0] as any).textbox;
      expect(compiledBox.style).toBeDefined();
    });
  });
});
