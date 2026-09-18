import type { DocumentOptions, SectionChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { projectDocumentOptions } from "../project";
import { runStyleOf } from "./styles";

// Word 2010+ text effects (w14) ride the run's raw XML; the projection must
// surface them as the painter's structured LayoutTextStyle fields.
const WORD_EFFECTS = [
  '<w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr">',
  '<w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill>',
  '<w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>',
  '<w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000"',
  ' w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">',
  '<w14:srgbClr w14:val="000000"><w14:alpha w14:val="38000"/></w14:srgbClr></w14:shadow>',
  '<w14:glow w14:rad="63500"><w14:srgbClr w14:val="FFC000">',
  '<w14:alpha w14:val="60000"/></w14:srgbClr></w14:glow>',
  '<w14:reflection w14:blurRad="6350" w14:stA="52000" w14:stPos="0" w14:endA="300"',
  ' w14:endPos="55000" w14:dist="1000" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000"',
  ' w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>',
  '<w14:props3d w14:prstMaterial="plastic">',
  '<w14:bevelT w14:w="63500" w14:h="50800" w14:prst="cross"/>',
  '<w14:bevelB w14:w="12700" w14:h="12700" w14:prst="softRound"/></w14:props3d>',
  '<w14:scene3d><w14:camera w14:prst="perspectiveFront"/>',
  '<w14:lightRig w14:rig="threePt" w14:dir="t">',
  '<w14:rot w14:lat="1200000" w14:lon="600000" w14:rev="5400000"/>',
  "</w14:lightRig></w14:scene3d>",
].join("");

describe("runStyleOf w14 text effects", () => {
  const style = runStyleOf({ w14RawXml: WORD_EFFECTS });

  it("resolves the outline with its width in px", () => {
    expect(style.outline).toEqual({ color: "FF0000", widthPx: 1 });
  });

  it("resolves the shadow into a px offset and rgba ink", () => {
    const shadow = style.shadow as { x?: number; y?: number; blur?: number; color?: string };
    expect(shadow).toMatchObject({ blur: 4, color: "rgba(0,0,0,0.38)" });
    expect(shadow.x).toBeCloseTo(Math.SQRT2, 5);
    expect(shadow.y).toBeCloseTo(Math.SQRT2, 5);
  });

  it("resolves the glow radius and alpha", () => {
    expect(style.glow?.radiusPx).toBeCloseTo(63500 / 9525, 5);
    expect(style.glow?.color).toBe("rgba(255,192,0,0.6)");
  });

  it("resolves the reflection", () => {
    expect(style.reflection).toMatchObject({
      blur: 6350 / 9525,
      distancePx: 1000 / 9525,
      opacity: 0.52,
    });
  });

  it("resolves both bevels and the 3-D rotation", () => {
    expect(style.bevel?.top?.preset).toBe("cross");
    expect(style.bevel?.bottom?.preset).toBe("softRound");
    expect(style.rotation3d).toEqual({ x: 20, y: 10, z: 90 });
  });

  it("leaves a run without w14 effects untouched", () => {
    const plain = runStyleOf({ size: 24 });
    expect(plain.outline).toBeUndefined();
    expect(plain.shadow).toBeUndefined();
    expect(plain.glow).toBeUndefined();
    expect(plain.reflection).toBeUndefined();
    expect(plain.bevel).toBeUndefined();
    expect(plain.rotation3d).toBeUndefined();
  });
});

describe("projectDocumentOptions text effects", () => {
  const children: SectionChild[] = [
    {
      paragraph: {
        children: [{ text: "Glow", w14RawXml: WORD_EFFECTS }],
      },
    },
  ];
  const doc: DocumentOptions = { sections: [{ children }] };

  it("carries the resolved effects onto the laid text run", () => {
    const blocks = projectDocumentOptions(doc).sections[0]!.blocks;
    const block = blocks[0];
    if (block?.kind !== "paragraph") throw new Error("expected paragraph");
    const atom = block.inline[0];
    if (atom?.kind !== "text") throw new Error("expected text atom");
    expect((atom.style.outline as { color?: string } | undefined)?.color).toBe("FF0000");
    expect(atom.style.glow?.color).toBe("rgba(255,192,0,0.6)");
    expect(atom.style.bevel?.top?.preset).toBe("cross");
    expect(atom.style.rotation3d?.z).toBe(90);
  });
});
