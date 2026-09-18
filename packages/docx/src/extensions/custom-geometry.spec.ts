import { generateDocumentSync, parseDocumentSync } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";

describe("Custom Geometry (a:custGeom) round-trip", () => {
  it("preserves customGeometry path commands across resolve and compile", () => {
    const customGeometry = {
      pathList: [
        {
          w: 2000,
          h: 1000,
          commands: [
            { command: "moveTo" as const, point: { x: "0", y: "0" } },
            { command: "lnTo" as const, point: { x: "2000", y: "0" } },
            {
              command: "quadBezTo" as const,
              points: [
                { x: "1500", y: "500" },
                { x: "1000", y: "1000" },
              ],
            },
            {
              command: "cubicBezTo" as const,
              points: [
                { x: "800", y: "900" },
                { x: "400", y: "700" },
                { x: "0", y: "500" },
              ],
            },
            { command: "close" as const },
          ],
        },
      ],
    };

    const wpsShape = {
      customGeometry,
      transformation: { width: 1828800, height: 914400 },
      fill: { type: "solid" as const, color: "ED7D31" },
      outline: { color: "4472C4", width: 12700 },
    };

    const doc = {
      sections: [{ children: [{ paragraph: { children: [{ wpsShape: wpsShape as any }] } }] }],
    };

    const resolved = resolveDocument(doc as any, docxExtensions);
    const node = resolved.content?.[0]?.content?.[0];
    expect(node?.type).toBe("wpsShape");
    expect((node?.attrs as any)?.wpsShape?.customGeometry).toEqual(customGeometry);

    const compiled = compileDocument(resolved, docxExtensions);
    const compiledShape = (compiled.sections[0]!.children[0] as any).paragraph.children[0].wpsShape;
    expect(compiledShape.customGeometry).toEqual(customGeometry);
  });

  it("surfaces through full OPC zip package round-trip", () => {
    const customGeometry = {
      pathList: [
        {
          w: 1000,
          h: 1000,
          commands: [
            { command: "moveTo" as const, point: { x: "0", y: "0" } },
            { command: "lnTo" as const, point: { x: "1000", y: "0" } },
            { command: "lnTo" as const, point: { x: "500", y: "1000" } },
            { command: "close" as const },
          ],
        },
      ],
    };

    const doc = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    wpsShape: {
                      customGeometry,
                      transformation: { width: 1270000, height: 1270000 },
                      fill: { type: "solid" as const, color: "4472C4" },
                    } as any,
                  },
                ],
              },
            },
          ],
        },
      ],
    };

    const binary = generateDocumentSync(doc as any);
    const parsed = parseDocumentSync(new Uint8Array(binary as Buffer));
    const json = resolveDocument(parsed, docxExtensions);
    const node = json.content?.[0]?.content?.[0];
    expect(node?.type).toBe("wpsShape");
    expect((node?.attrs as any)?.wpsShape?.customGeometry).toBeDefined();
  });
});
