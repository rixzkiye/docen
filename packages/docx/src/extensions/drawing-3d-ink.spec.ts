import type { DocumentOptions, ParagraphChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";
import {
  DELEGATION_NOTICE,
  INK_2010_URI,
  INK_MAIN_URI,
  INKML_MIME,
  MODEL3D_URI,
  is3DModelChild,
  is3DModelXml,
  isInkChild,
  isInkXml,
  parse3DModelDocx,
  parseInkDocx,
  render3DModelDocx,
  renderInkDocx,
  serialize3DModelXml,
  serializeInkXml,
  updateCamera,
  updateDocPr,
  updateExtent,
} from "./drawing-3d-ink";

describe("W6.6 Ink & 3D Models", () => {
  describe("DELEGATION_NOTICE", () => {
    it("matches the exact required mandate string", () => {
      expect(DELEGATION_NOTICE).toBe(
        "3D mesh vertex manipulation and ink vector splitting are delegated to external 3D/ink tools; viewing, layout, sizing, and alt text are editable natively.",
      );
    });
  });

  describe("Detection Helpers", () => {
    it("detects 3D Model XML correctly", () => {
      expect(
        is3DModelXml(
          `<wp:inline><a:graphicData uri="${MODEL3D_URI}"><am3d:model3D/></a:graphicData></wp:inline>`,
        ),
      ).toBe(true);
      expect(is3DModelXml(`<w16:contentPart r:id="rId5" contentType="model/gltf-binary"/>`)).toBe(
        true,
      );
      expect(
        is3DModelXml(
          `<w:drawing><wp:anchor><a:graphicData><am3d:model3D/></a:graphicData></wp:anchor></w:drawing>`,
        ),
      ).toBe(true);
      expect(
        is3DModelXml(
          `<w:drawing><wp:inline><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></wp:inline></w:drawing>`,
        ),
      ).toBe(false);
    });

    it("detects Ink XML correctly", () => {
      expect(
        isInkXml(
          `<wp:inline><a:graphicData uri="${INK_2010_URI}"><msink:ink/></a:graphicData></wp:inline>`,
        ),
      ).toBe(true);
      expect(
        isInkXml(
          `<wp:anchor><a:graphicData uri="${INK_MAIN_URI}"><ink:ink/></a:graphicData></wp:anchor>`,
        ),
      ).toBe(true);
      expect(isInkXml(`<w16:contentPart r:id="rId6" contentType="${INKML_MIME}"/>`)).toBe(true);
      expect(isInkXml(`<w:r><w:t>Hello</w:t></w:r>`)).toBe(false);
    });

    it("detects 3D and Ink paragraph children", () => {
      expect(
        is3DModelChild({ rawXml: `<am3d:model3D xmlns:am3d="${MODEL3D_URI}"/>` } as ParagraphChild),
      ).toBe(true);
      expect(
        isInkChild({ rawXml: `<msink:ink xmlns:msink="${INK_2010_URI}"/>` } as ParagraphChild),
      ).toBe(true);
      expect(is3DModelChild({ rawXml: `<w:drawing/>` } as ParagraphChild)).toBe(false);
      expect(isInkChild({ rawXml: `<w:drawing/>` } as ParagraphChild)).toBe(false);
    });
  });

  describe("XML Updaters", () => {
    it("updates extent cx and cy in XML", () => {
      const xml = `<wp:inline><wp:extent cx="1000" cy="2000"/></wp:inline>`;
      const updated = updateExtent(xml, 3000, 4000);
      expect(updated).toContain('cx="3000"');
      expect(updated).toContain('cy="4000"');
    });

    it("updates docPr title and descr in XML", () => {
      const xml = `<wp:inline><wp:docPr id="1" name="Model 1" descr="Old description"/><wp:extent cx="100" cy="100"/></wp:inline>`;
      const updated = updateDocPr(xml, { title: "New Title", descr: "Updated alt text" });
      expect(updated).toContain('title="New Title"');
      expect(updated).toContain('descr="Updated alt text"');
      expect(updated).toContain('name="Model 1"');
    });

    it("updates camera rotation in 3D model XML", () => {
      const xml = `<am3d:model3D><am3d:camera><am3d:rot lat="100" lon="200" rev="300"/></am3d:camera></am3d:model3D>`;
      const updated = updateCamera(xml, { lat: 500, lon: 600, rev: 700 });
      expect(updated).toContain('lat="500"');
      expect(updated).toContain('lon="600"');
      expect(updated).toContain('rev="700"');
    });
  });

  describe("3D Model Parsing & Serialization", () => {
    const raw3D = `<w:drawing>
      <wp:inline>
        <wp:extent cx="1905000" cy="1905000"/>
        <wp:docPr id="10" name="3D Cube" title="Astronaut" descr="Detailed 3D model of astronaut"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="${MODEL3D_URI}">
            <am3d:model3D xmlns:am3d="${MODEL3D_URI}" r:id="rId10">
              <am3d:camera fov="4500000">
                <am3d:rot lat="1200000" lon="2400000" rev="0"/>
              </am3d:camera>
            </am3d:model3D>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>`;

    it("parses 3D model properties accurately", () => {
      const parsed = parse3DModelDocx(raw3D);
      expect(parsed).toBeDefined();
      expect(parsed.type).toBe("model3d");
      const attrs = parsed.attrs as Record<string, any>;
      expect(attrs?.cx).toBe(1905000);
      expect(attrs?.cy).toBe(1905000);
      expect(attrs?.width).toBe(200); // 1905000 / 9525
      expect(attrs?.height).toBe(200);
      expect(attrs?.title).toBe("Astronaut");
      expect(attrs?.descr).toBe("Detailed 3D model of astronaut");
      expect(attrs?.name).toBe("3D Cube");
      expect(attrs?.camera?.fov).toBe(4500000);
      expect(attrs?.camera?.rot?.lat).toBe(1200000);
      expect(attrs?.camera?.rot?.lon).toBe(2400000);
      expect(attrs?.rawXml).toBe(raw3D);
    });

    it("serializes 3D model to lossless DrawingML XML", () => {
      const serialized = serialize3DModelXml({
        cx: 952500,
        cy: 952500,
        width: 100,
        height: 100,
        title: "Satellite",
        descr: "Satellite alt text",
        name: "Model3D 1",
        rId: "rId2",
        camera: { rot: { lat: 1000, lon: 2000, rev: 0 } },
      });
      expect(serialized).toContain(MODEL3D_URI);
      expect(serialized).toContain('title="Satellite"');
      expect(serialized).toContain('descr="Satellite alt text"');
      expect(serialized).toContain('cx="952500"');
      expect(serialized).toContain('cy="952500"');
    });

    it("renders 3D model child preserving rawXml with updated extents and docPr", () => {
      const child = render3DModelDocx({
        rawXml: raw3D,
        cx: 2857500,
        cy: 2857500,
        width: 300,
        height: 300,
        title: "Updated Title",
        descr: "Updated Description",
      });
      const xml = (child as { rawXml: string }).rawXml;
      expect(xml).toContain('cx="2857500"');
      expect(xml).toContain('cy="2857500"');
      expect(xml).toContain('title="Updated Title"');
      expect(xml).toContain('descr="Updated Description"');
    });
  });

  describe("Ink Parsing & Serialization", () => {
    const rawInk = `<w:drawing>
      <wp:inline>
        <wp:extent cx="1428750" cy="952500"/>
        <wp:docPr id="11" name="Ink Drawing 1" title="Signature" descr="Handwritten cursive signature"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="${INK_2010_URI}">
            <msink:ink xmlns:msink="${INK_2010_URI}">
              <msink:trace>M 10 10 L 20 20</msink:trace>
            </msink:ink>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>`;

    it("parses Ink properties accurately", () => {
      const parsed = parseInkDocx(rawInk);
      expect(parsed).toBeDefined();
      expect(parsed.type).toBe("ink");
      const attrs = parsed.attrs as Record<string, any>;
      expect(attrs?.cx).toBe(1428750);
      expect(attrs?.cy).toBe(952500);
      expect(attrs?.width).toBe(150); // 1428750 / 9525
      expect(attrs?.height).toBe(100); // 952500 / 9525
      expect(attrs?.title).toBe("Signature");
      expect(attrs?.descr).toBe("Handwritten cursive signature");
      expect(attrs?.rawXml).toBe(rawInk);
    });

    it("serializes Ink to lossless DrawingML XML", () => {
      const serialized = serializeInkXml({
        cx: 952500,
        cy: 952500,
        width: 100,
        height: 100,
        title: "Sketch",
        descr: "Hand-drawn diagram sketch",
        name: "Ink 1",
        traces: ["M 0 0 L 100 100"],
      });
      expect(serialized).toContain(INK_2010_URI);
      expect(serialized).toContain('title="Sketch"');
      expect(serialized).toContain('descr="Hand-drawn diagram sketch"');
      expect(serialized).toContain('cx="952500"');
    });

    it("renders Ink child preserving rawXml with updated extents and docPr", () => {
      const child = renderInkDocx({
        rawXml: rawInk,
        cx: 1905000,
        cy: 1905000,
        width: 200,
        height: 200,
        title: "New Signature",
        descr: "Updated Ink Description",
      });
      const xml = (child as { rawXml: string }).rawXml;
      expect(xml).toContain('cx="1905000"');
      expect(xml).toContain('cy="1905000"');
      expect(xml).toContain('title="New Signature"');
      expect(xml).toContain('descr="Updated Ink Description"');
    });
  });

  describe("Document Round-Trip Integration", () => {
    it("round-trips 3D model through resolveDocument and compileDocument", () => {
      const raw3D = `<w:drawing>
        <wp:inline>
          <wp:extent cx="1905000" cy="1905000"/>
          <wp:docPr id="50" name="3D Cube" title="Cube Model" descr="Cube alt text"/>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="${MODEL3D_URI}">
              <am3d:model3D xmlns:am3d="${MODEL3D_URI}" r:id="rId50"/>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>`;

      const doc: DocumentOptions = {
        sections: [
          { children: [{ paragraph: { children: [{ rawXml: raw3D } as ParagraphChild] } }] },
        ],
      };

      const resolved = resolveDocument(doc, docxExtensions);
      const node = (
        resolved.content?.[0] as { content?: { type: string; attrs?: Record<string, unknown> }[] }
      )?.content?.find((n) => n.type === "model3d");

      expect(node).toBeDefined();
      expect(node?.attrs?.title).toBe("Cube Model");
      expect(node?.attrs?.descr).toBe("Cube alt text");
      expect(node?.attrs?.cx).toBe(1905000);

      const compiled = compileDocument(resolved, docxExtensions);
      const p = compiled.sections[0]!.children[0] as { paragraph: { children?: ParagraphChild[] } };
      const outChild = p.paragraph.children?.[0] as { rawXml?: string };

      expect(outChild).toBeDefined();
      expect(outChild.rawXml).toContain(MODEL3D_URI);
      expect(outChild.rawXml).toContain('title="Cube Model"');
    });

    it("round-trips Ink through resolveDocument and compileDocument", () => {
      const rawInk = `<w:drawing>
        <wp:inline>
          <wp:extent cx="1428750" cy="952500"/>
          <wp:docPr id="51" name="Ink 51" title="Formula Ink" descr="Math formula handwritten"/>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="${INK_2010_URI}">
              <msink:ink xmlns:msink="${INK_2010_URI}"/>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>`;

      const doc: DocumentOptions = {
        sections: [
          { children: [{ paragraph: { children: [{ rawXml: rawInk } as ParagraphChild] } }] },
        ],
      };

      const resolved = resolveDocument(doc, docxExtensions);
      const node = (
        resolved.content?.[0] as { content?: { type: string; attrs?: Record<string, unknown> }[] }
      )?.content?.find((n) => n.type === "ink");

      expect(node).toBeDefined();
      expect(node?.attrs?.title).toBe("Formula Ink");
      expect(node?.attrs?.descr).toBe("Math formula handwritten");
      expect(node?.attrs?.cx).toBe(1428750);

      const compiled = compileDocument(resolved, docxExtensions);
      const p = compiled.sections[0]!.children[0] as { paragraph: { children?: ParagraphChild[] } };
      const outChild = p.paragraph.children?.[0] as { rawXml?: string };

      expect(outChild).toBeDefined();
      expect(outChild.rawXml).toContain(INK_2010_URI);
      expect(outChild.rawXml).toContain('title="Formula Ink"');
    });
  });
});
