// Regenerates test/fixtures/golden-records.json with the harfbuzzjs oracle
// (the same HarfBuzz shaping engine family as the hb-shape CLI). Run from the
// repository root:
//
//   node packages/shaping/test/generate-goldens.mjs
//
// The generated file's `meta` records the generator and versions so the
// golden.spec.ts records are traceable to a reproducible oracle.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import * as hb from "harfbuzzjs";

const dir = path.dirname(fileURLToPath(import.meta.url));
const fontsDir = path.join(dir, "fixtures/fonts");
const harfbuzzjsVersion = JSON.parse(
  fs.readFileSync(new URL("../node_modules/harfbuzzjs/package.json", import.meta.url), "utf8"),
).version;

const openSans = fs.readFileSync(path.join(fontsDir, "OpenSans-Regular.ttf"));
const arabic = fs.readFileSync(path.join(fontsDir, "NotoNaskhArabic-Regular.ttf"));

function shape(fontBytes, text, options = {}) {
  const blob = new hb.Blob(fontBytes);
  const face = new hb.Face(blob, 0);
  const font = new hb.Font(face);
  const buffer = new hb.Buffer();
  buffer.addText(text);
  if (options.direction === "rtl") buffer.setDirection(hb.Direction.RTL);
  buffer.guessSegmentProperties();
  hb.shape(font, buffer, []);
  const glyphs = buffer.getGlyphInfosAndPositions().map((g) => ({
    glyphId: g.codepoint,
    cluster: g.cluster,
    xAdvance: g.xAdvance ?? 0,
    yAdvance: g.yAdvance ?? 0,
    xOffset: g.xOffset ?? 0,
    yOffset: g.yOffset ?? 0,
  }));
  font.destroy?.();
  face.destroy?.();
  blob.destroy?.();
  return glyphs;
}

const records = {
  meta: {
    generator: "packages/shaping/test/generate-goldens.mjs",
    oracle: "harfbuzzjs (HarfBuzz)",
    harfbuzzjs: harfbuzzjsVersion,
    fonts: {
      latin: "OpenSans-Regular.ttf",
      arabic: "NotoNaskhArabic-Regular.ttf",
    },
    texts: {
      latinAlphabet: "The quick brown fox jumps over the lazy dog.",
      kerningPairs: "AV To Wa LT",
      ligatures: "office flight waffle",
      arabicGreeting: "مرحبا",
    },
  },
  latinAlphabet: shape(openSans, "The quick brown fox jumps over the lazy dog."),
  kerningPairs: shape(openSans, "AV To Wa LT"),
  ligatures: shape(openSans, "office flight waffle"),
  arabicGreeting: shape(arabic, "مرحبا", { direction: "rtl" }),
};

const outPath = path.join(dir, "fixtures/golden-records.json");
fs.writeFileSync(outPath, `${JSON.stringify(records, null, 2)}\n`);
console.log(`wrote ${outPath}`);
