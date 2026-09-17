// Shaped-path parity: when the canvas measures through the exact advances the
// shaper produces (a canvas stub delegating to the same FontRef), the two
// measurers must agree on wrapping — line count and every line's break index —
// for Latin, CJK and mixed runs. An independent oracle in the sense that the
// canvas side goes through pretext's own packer, and any divergence in the
// shaped breaker's segment provider shows up as a different break.

// @vitest-environment node
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createFontRefSync, initShapingWasm, type FontRef } from "@docen/shaping";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { layoutBlock } from "../block/block";
import { browserFontMetrics, clearRegisteredFontRatios } from "../font";
import type { LayoutBlock, LayoutParagraph, LayoutTextStyle } from "../layout-doc";
import { TextMeasurer } from "./measure";
import { ShapedMeasurer } from "./shaped-measurer";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const LATIN_FONT = path.resolve(
  __dirname,
  "../../../shaping/test/fixtures/fonts/OpenSans-Regular.ttf",
);
const CJK_FONT = "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc";
const hasCjk = fs.existsSync(CJK_FONT);

const LATIN_FAMILY = "Parity Open Sans";
const CJK_FAMILY = "Parity Noto CJK";
const WIDTH = 300;

/** The same faces the shaped measurer uses, keyed by lowercased family. */
const faces = new Map<string, FontRef>();

/** Parse `16px "Family"` out of pretext's CSS shorthand. */
function fontOf(shorthand: string): { sizePx: number; family: string } {
  const sizePx = Number(/(\d+(?:\.\d+)?)px/.exec(shorthand)?.[1] ?? 16);
  const family = /"([^"]+)"/.exec(shorthand)?.[1] ?? "";
  return { sizePx, family };
}

/**
 * A canvas metered in the shaper's own advances: pretext's canvas path and
 * the shaped path then receive identical metrics, so equal breaks are a
 * property of the pipeline, not of the font.
 */
function installShapedCanvas(): void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- duck-typed
  (globalThis as any).OffscreenCanvas = class {
    getContext(): unknown {
      let font = "16px serif";
      return {
        set font(v: string) {
          font = v;
        },
        get font(): string {
          return font;
        },
        measureText(text: string): { width: number } {
          const { sizePx, family } = fontOf(font);
          const face = faces.get(family.toLowerCase());
          if (!face || !text) return { width: text.length * (sizePx / 2) };
          return {
            width: (face.shape(text).totalAdvance * sizePx) / (face.metrics.unitsPerEm || 1000),
          };
        },
      };
    }
  };
}

const paragraph = (text: string, family: string): LayoutParagraph => {
  const style: LayoutTextStyle = { family, sizePx: 16 };
  return {
    kind: "paragraph",
    defaultTextStyle: style,
    inline: [{ kind: "text", text, style }],
  };
};

const linesOf = (laid: unknown): string[] =>
  (laid as { lines: { items: { text?: string }[] }[] }).lines.map((l) =>
    l.items.map((i) => i.text ?? "").join(""),
  );

const CASES: [name: string, text: string, family: string][] = [
  [
    "latin",
    "The quick brown fox jumps over the lazy dog before typography breaks it. ".repeat(3),
    LATIN_FAMILY,
  ],
  [
    "cjk",
    "排版引擎在换行时会按照中日韩的禁则规则处理标点符号，句号不能出现在行首。".repeat(4),
    CJK_FAMILY,
  ],
  [
    "mixed",
    "混排段落 Latin words wrap at spaces while 中文可以在任意字之间断开，标点遵守禁则。 ".repeat(4),
    CJK_FAMILY,
  ],
];

describe.skipIf(!hasCjk)("shaped vs canvas wrap parity (same advances)", () => {
  let canvas: TextMeasurer;
  let shaped: ShapedMeasurer;

  beforeAll(async () => {
    installShapedCanvas();
    await initShapingWasm();
    const latinRef = createFontRefSync(new Uint8Array(fs.readFileSync(LATIN_FONT)));
    const cjkRef = createFontRefSync(new Uint8Array(fs.readFileSync(CJK_FONT)), { index: 2 });
    shaped = new ShapedMeasurer(browserFontMetrics, { enabled: true });
    shaped.registerFont(LATIN_FAMILY, latinRef);
    shaped.registerFont(CJK_FAMILY, cjkRef);
    faces.set(LATIN_FAMILY.toLowerCase(), latinRef);
    faces.set(CJK_FAMILY.toLowerCase(), cjkRef);
    canvas = new TextMeasurer(browserFontMetrics);
  });

  afterAll(() => {
    clearRegisteredFontRatios();
    for (const face of faces.values()) face.dispose();
  });

  for (const [name, text, family] of CASES) {
    it(`wraps the ${name} sample identically`, () => {
      const canvasLines = linesOf(
        layoutBlock(paragraph(text, family) as LayoutBlock, WIDTH, undefined, canvas),
      );
      const shapedLines = linesOf(
        layoutBlock(paragraph(text, family) as LayoutBlock, WIDTH, undefined, shaped),
      );
      expect(shapedLines).toEqual(canvasLines);
      // The sample must actually wrap (a single line would make parity trivial).
      expect(canvasLines.length).toBeGreaterThan(2);
    });
  }
});
