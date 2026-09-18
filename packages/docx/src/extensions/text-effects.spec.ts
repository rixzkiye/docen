import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, type JSONContent } from "../index";
import {
  applyTextEffect,
  getTextEffectChunk,
  parseTextEffects,
  serializeTextEffect,
  setTextEffectChunk,
  splitW14Elements,
  type TextEffects,
} from "./text-effects";

// The w14 text effects a Word 2010+ document stores in w:rPr — office-open
// preserves them verbatim as RunOptions.w14RawXml; this module is the
// structured parse/serialize layer the gallery, the projection, and the
// painter share.
const WORD_EFFECTS = [
  '<w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr">',
  '<w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill>',
  '<w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>',
  '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="FFFFFF"/></w14:solidFill></w14:textFill>',
  '<w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000"',
  ' w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">',
  '<w14:srgbClr w14:val="000000"><w14:alpha w14:val="38000"/></w14:srgbClr></w14:shadow>',
  '<w14:glow w14:rad="63500"><w14:srgbClr w14:val="FFC000">',
  '<w14:alpha w14:val="60000"/></w14:srgbClr></w14:glow>',
  '<w14:reflection w14:blurRad="6350" w14:stA="52000" w14:stPos="0" w14:endA="300"',
  ' w14:endPos="55000" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000"',
  ' w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>',
  '<w14:props3d w14:prstMaterial="plastic">',
  '<w14:bevelT w14:w="63500" w14:h="50800" w14:prst="cross"/>',
  '<w14:bevelB w14:w="12700" w14:h="12700" w14:prst="softRound"/></w14:props3d>',
  '<w14:scene3d><w14:camera w14:prst="perspectiveFront"/>',
  '<w14:lightRig w14:rig="threePt" w14:dir="t">',
  '<w14:rot w14:lat="1200000" w14:lon="600000" w14:rev="5400000"/>',
  "</w14:lightRig></w14:scene3d>",
].join("");

describe("splitW14Elements", () => {
  it("returns each top-level w14 chunk in document order", () => {
    const chunks = splitW14Elements(WORD_EFFECTS);
    expect(chunks.map((chunk) => chunk.name)).toEqual([
      "textOutline",
      "textFill",
      "shadow",
      "glow",
      "reflection",
      "props3d",
      "scene3d",
    ]);
    // Self-closing chunks are captured whole; paired chunks include content.
    expect(chunks[4]!.xml.endsWith("/>")).toBe(true);
    expect(chunks[2]!.xml).toContain("<w14:alpha");
  });
});

describe("parseTextEffects", () => {
  const effects = parseTextEffects(WORD_EFFECTS);

  it("parses the outline, keeping the color and width", () => {
    expect(effects.outline).toEqual({ color: "FF0000", widthPx: 1, dash: "solid" });
  });

  it("parses the shadow direction/distance/alpha", () => {
    expect(effects.shadow).toMatchObject({ color: "000000", opacity: 0.38, distPx: 2, dirDeg: 45 });
    expect(effects.shadow!.blurPx).toBeCloseTo(4, 5);
  });

  it("parses the glow radius and alpha", () => {
    expect(effects.glow).toMatchObject({ color: "FFC000", opacity: 0.6 });
    expect(effects.glow!.radiusPx).toBeCloseTo(63500 / 9525, 5);
  });

  it("parses the reflection opacities", () => {
    expect(effects.reflection).toMatchObject({ startOpacity: 0.52, endOpacity: 0.003, distPx: 0 });
    expect(effects.reflection!.blurPx).toBeCloseTo(6350 / 9525, 5);
  });

  it("parses both bevels", () => {
    expect(effects.bevel?.top).toMatchObject({ preset: "cross" });
    expect(effects.bevel?.top?.widthPx).toBeCloseTo(63500 / 9525, 5);
    expect(effects.bevel?.bottom).toMatchObject({ preset: "softRound" });
  });

  it("parses the 3-D rotation sphere and camera", () => {
    expect(effects.rotation).toEqual({ x: 20, y: 10, z: 90, camera: "perspectiveFront" });
  });

  it("returns an empty set for absent or unrelated XML", () => {
    expect(parseTextEffects(null)).toEqual({});
    expect(parseTextEffects('<w14:ligatures w14:val="all"/>')).toEqual({});
  });
});

describe("serializeTextEffect", () => {
  const effects = parseTextEffects(WORD_EFFECTS);

  it("round-trips every parsed effect through serialize → parse", () => {
    const kinds = ["outline", "shadow", "glow", "reflection", "bevel", "rotation"] as const;
    for (const kind of kinds) {
      const chunk = serializeTextEffect(kind, effects);
      expect(chunk, kind).toBeTruthy();
      const reparsed = parseTextEffects(chunk);
      expect(reparsed[kind], kind).toBeTruthy();
      // Widths/distances survive the EMU round-trip (within rounding).
      if (kind === "outline") expect(reparsed.outline!.color).toBe("FF0000");
      if (kind === "shadow") expect(reparsed.shadow!.dirDeg).toBe(45);
      if (kind === "glow") expect(reparsed.glow!.color).toBe("FFC000");
      if (kind === "reflection") expect(reparsed.reflection!.startOpacity).toBe(0.52);
      if (kind === "bevel") expect(reparsed.bevel!.top!.preset).toBe("cross");
      if (kind === "rotation")
        expect(reparsed.rotation).toEqual({
          x: 20,
          y: 10,
          z: 90,
          camera: "perspectiveFront",
        });
    }
  });

  it("produces Word-shaped XML (valid w14 element names and attributes)", () => {
    const outline = serializeTextEffect("outline", effects)!;
    expect(outline).toContain('<w14:textOutline w14:w="9525"');
    expect(outline).toContain('<w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill>');
    expect(outline).toContain('<w14:prstDash w14:val="solid"/><w14:round/>');
    const rotation = serializeTextEffect("rotation", effects)!;
    expect(rotation).toContain('<w14:camera w14:prst="perspectiveFront"/>');
    expect(rotation).toContain('<w14:rot w14:lat="1200000" w14:lon="600000" w14:rev="5400000"/>');
  });
});

describe("applyTextEffect preservation", () => {
  it("swaps only the addressed chunk and keeps the rest byte-for-byte", () => {
    const before = WORD_EFFECTS;
    const withShadow = applyTextEffect(before, "shadow", {
      shadow: { color: "FF0000", opacity: 1, blurPx: 0, distPx: 3, dirDeg: 90 },
    })!;
    // The untouched chunks survive verbatim.
    expect(withShadow).toContain(getTextEffectChunk(before, "outline")!);
    expect(withShadow).toContain(getTextEffectChunk(before, "fill")!);
    expect(withShadow).toContain(getTextEffectChunk(before, "rotation")!);
    expect(getTextEffectChunk(withShadow, "fill")).toBe(getTextEffectChunk(before, "fill"));
    // Exactly one shadow remains, the new one.
    expect(splitW14Elements(withShadow).filter((c) => c.name === "shadow")).toHaveLength(1);
    expect(parseTextEffects(withShadow).shadow).toMatchObject({ color: "FF0000", dirDeg: 90 });
  });

  it("clears one effect without touching the others", () => {
    const cleared = applyTextEffect(WORD_EFFECTS, "glow", null)!;
    expect(splitW14Elements(cleared).some((c) => c.name === "glow")).toBe(false);
    expect(getTextEffectChunk(cleared, "reflection")).toBe(
      getTextEffectChunk(WORD_EFFECTS, "reflection"),
    );
  });

  it("adding to an empty carrier starts fresh; clearing the last effect drops the XML", () => {
    const added = setTextEffectChunk(
      null,
      "outline",
      serializeTextEffect("outline", {
        outline: { widthPx: 1, color: "000000" },
      }),
    );
    expect(splitW14Elements(added)).toHaveLength(1);
    expect(applyTextEffect(added, "outline", null)).toBeNull();
  });
});

describe("DOCX round-trip", () => {
  const runDoc = (w14RawXml: string): JSONContent => ({
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "text",
            text: "Effect",
            marks: [{ type: "textStyle", attrs: { w14RawXml } }],
          },
        ],
      },
    ],
  });

  const markXmlOf = (json: JSONContent): string | undefined => {
    const text = (json.content?.[0]?.content ?? [])[0] as JSONContent;
    const mark = (text.marks ?? []).find((m) => m.type === "textStyle");
    return (mark?.attrs as { w14RawXml?: string } | undefined)?.w14RawXml;
  };

  it("preserves the raw w14 XML through generate → parse byte-for-byte", () => {
    const doc = runDoc(WORD_EFFECTS);
    const parsed = parseDOCXSync(generateDOCXSync(doc) as Uint8Array);
    expect(markXmlOf(parsed)).toBe(WORD_EFFECTS);
    // ...and the structured parse still reads the same effects.
    const before = parseTextEffects(WORD_EFFECTS);
    const after = parseTextEffects(markXmlOf(parsed));
    expect(after).toEqual(before);
  });

  it("preserves an editor-applied effect end-to-end", () => {
    const applied: TextEffects = {
      outline: { widthPx: 1.33, color: "2E75B5" },
      bevel: { top: { widthPx: 5.3, heightPx: 4.2, preset: "circle" } },
    };
    let xml = applyTextEffect(null, "outline", applied);
    xml = applyTextEffect(xml, "bevel", applied);
    const parsed = parseDOCXSync(generateDOCXSync(runDoc(xml!)) as Uint8Array);
    const reparsed = parseTextEffects(markXmlOf(parsed));
    expect(reparsed.outline?.color).toBe("2E75B5");
    expect(reparsed.bevel?.top?.preset).toBe("circle");
    expect(reparsed.bevel?.top?.widthPx).toBeCloseTo(5.3, 3);
    expect(reparsed.bevel?.top?.heightPx).toBeCloseTo(4.2, 3);
  });
});
