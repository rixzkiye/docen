import { describe, expect, it } from "vitest";

import {
  composeMatrix,
  getActiveKit,
  IDENTITY_MATRIX,
  leaferKit,
  localMatrixOf,
  NodeBox,
  NodeEllipse,
  NodeGroup,
  nodeKit,
  NodeLine,
  NodePath,
  NodeRect,
  NodeText,
  Rect,
  withKit,
} from "./kit";

describe("PaintKit abstraction", () => {
  it("computes 2D affine matrix compositions correctly", () => {
    expect(IDENTITY_MATRIX).toEqual({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
    const m1 = localMatrixOf(10, 20, 2, 2);
    const m2 = localMatrixOf(5, 5, 1, 1);
    const composed = composeMatrix(m1, m2);
    expect(composed.e).toBe(20); // 10 + 2 * 5
    expect(composed.f).toBe(30); // 20 + 2 * 5
    expect(composed.a).toBe(2);
    expect(composed.d).toBe(2);
  });

  it("builds nodeKit element hierarchy and computes world transforms", () => {
    const root = new NodeGroup({ x: 50, y: 100 });
    const childGroup = nodeKit.createGroup({ x: 10, y: 20 });
    expect(leaferKit).toBeDefined();
    const rect = nodeKit.createRect({
      x: 5,
      y: 5,
      width: 100,
      height: 50,
      fill: "#ff0000",
      cornerRadius: 4,
    });

    childGroup.add(rect);
    root.add(childGroup);

    expect(root.children).toHaveLength(1);
    expect(childGroup.children).toHaveLength(1);

    const world = rect.worldTransform;
    expect(world.e).toBe(65); // 50 + 10 + 5
    expect(world.f).toBe(125); // 100 + 20 + 5
  });

  it("generates correct SVG path strings for nodeKit shapes", () => {
    const rect = new NodeRect({ width: 200, height: 100 });
    expect(rect.getPathString()).toBe("M 0 0 L 200 0 L 200 100 L 0 100 Z");

    const roundedRect = new NodeRect({ width: 100, height: 100, cornerRadius: 10 });
    expect(roundedRect.getPathString()).toContain("Q");

    const line = new NodeLine({ points: [10, 20, 100, 200] });
    expect(line.getPathString()).toBe("M 10 20 L 100 200");

    const path = new NodePath({ path: "M 0 0 C 10 10 20 20 30 30 Z" });
    expect(path.getPathString()).toBe("M 0 0 C 10 10 20 20 30 30 Z");

    const ellipse = new NodeEllipse({ width: 50, height: 30 });
    expect(ellipse.getPathString()).toContain("A 25 15");

    const box = new NodeBox({ width: 300, height: 400 });
    expect(box.getPathString()).toBe("M 0 0 L 300 0 L 300 400 L 0 400 Z");

    const text = new NodeText({ text: "Hello World", fontSize: 16, width: 120 });
    expect(text.textDrawData.rows[0]?.text).toBe("Hello World");
    // Leafer's baseline: ((lineHeight + 0.7 × fontSize) / 2). The painter pins
    // lineHeight to the size, so the row baseline is 0.85 × size — the value
    // the browser scene export produces (was 0.8 × size, a Node-only drift).
    expect(text.textDrawData.rows[0]?.y).toBe(16 * 0.85);
    const taller = new NodeText({ text: "Hello", fontSize: 10, lineHeight: 20 });
    expect(taller.textDrawData.rows[0]?.y).toBe((20 + 0.7 * 10) / 2);
  });

  it("distributes a justified Text's slack like Leafer's CharLayout", () => {
    // Word mode (Latin): the slack lands on the Leafer word boundaries; a
    // line-final whitespace run stays untrimmed/unstretched (Leafer's
    // trimRight) and the last word rides the interval edge.
    const word = new NodeText({
      text: "ab cd ef ",
      textAlign: "both-justify",
      width: 100,
      textNaturalWidth: 80,
      fontSize: 10,
    });
    const wordRow = word.textDrawData.rows[0]!;
    expect(wordRow.extras).toHaveLength(9);
    expect(wordRow.extras!.at(-1)).toBe(0);
    expect(wordRow.extras!.reduce((sum, extra) => sum + extra, 0)).toBeCloseTo(20, 6);

    // Letter mode (CJK / squeezed): one uniform share per code point, so the
    // last glyph's prefix shift fills the interval.
    const letter = new NodeText({
      text: "天地",
      textAlign: "both-letter",
      width: 100,
      textNaturalWidth: 80,
      fontSize: 10,
    });
    const letterRow = letter.textDrawData.rows[0]!;
    expect(letterRow.extras).toEqual([20, 20]);

    // Undefined natural width or a left-aligned Text keeps the natural rows
    // every other node-kit consumer relies on.
    expect(new NodeText({ text: "ab", width: 100 }).textDrawData.rows[0]!.extras).toBeUndefined();
    expect(
      new NodeText({ text: "ab", textAlign: "left", width: 100, textNaturalWidth: 80 }).textDrawData
        .rows[0]!.extras,
    ).toBeUndefined();
  });

  it("supports dynamic withKit context switching", () => {
    expect(getActiveKit()).toBe(leaferKit);

    const dummyLeaferKit = {
      createGroup: (p: any) => ({ tag: "CustomGroup", ...p }),
      createBox: (p: any) => ({ tag: "CustomBox", ...p }),
      createRect: (p: any) => ({ tag: "CustomRect", ...p }),
      createLine: (p: any) => ({ tag: "CustomLine", ...p }),
      createText: (p: any) => ({ tag: "CustomText", ...p }),
      createPath: (p: any) => ({ tag: "CustomPath", ...p }),
      createEllipse: (p: any) => ({ tag: "CustomEllipse", ...p }),
      createImage: (p: any) => ({ tag: "CustomImage", ...p }),
    };

    withKit(dummyLeaferKit, () => {
      expect(getActiveKit()).toBe(dummyLeaferKit);
      const r = new Rect({ width: 50, height: 50 });
      expect((r as any).tag).toBe("CustomRect");
    });

    expect(getActiveKit()).toBe(leaferKit);
  });
});
