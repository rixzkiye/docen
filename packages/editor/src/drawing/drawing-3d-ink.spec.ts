// @vitest-environment happy-dom
import { DELEGATION_NOTICE } from "@docen/docx";
import { NodeSelection } from "@tiptap/pm/state";
import { describe, expect, it, vi } from "vitest";

import { DocenAltTextPane } from "../ui/components/workspace/alt-text-pane";
import { drawingNodePos, drawingSelectionKind } from "./docx";
import { resizeBox, type Box } from "./geometry";

describe("Lane W6.6: Ink & 3D Model Editor Support", () => {
  describe("Drawing Selection & Position Detection", () => {
    it("identifies model3d node selection", () => {
      const mockState = {
        selection: Object.create(NodeSelection.prototype, {
          node: {
            value: {
              type: { name: "model3d" },
              attrs: {
                title: "Satellite",
                descr: "3D model of weather satellite",
                cx: 1905000,
                cy: 1905000,
              },
            },
          },
        }),
      } as any;

      expect(drawingSelectionKind(mockState)).toBe("model3d");
    });

    it("identifies ink node selection", () => {
      const mockState = {
        selection: Object.create(NodeSelection.prototype, {
          node: {
            value: {
              type: { name: "ink" },
              attrs: { title: "Signature", descr: "Handwritten ink", cx: 1428750, cy: 952500 },
            },
          },
        }),
      } as any;

      expect(drawingSelectionKind(mockState)).toBe("ink");
    });

    it("resolves drawingNodePos for floating and inline model3d / ink", () => {
      const children = [
        { type: { name: "model3d" }, attrs: { floating: { x: 10, y: 10 } }, nodeSize: 1 },
        { type: { name: "model3d" }, attrs: {}, nodeSize: 1 },
        { type: { name: "ink" }, attrs: { floating: { x: 20, y: 20 } }, nodeSize: 1 },
        { type: { name: "ink" }, attrs: {}, nodeSize: 1 },
      ];

      const mockDoc = {
        nodeAt: (pos: number) => {
          if (pos === 0) {
            return {
              childCount: 4,
              child: (i: number) => children[i],
              forEach: (fn: (child: any, offset: number, index: number) => void) => {
                let offset = 0;
                children.forEach((c, idx) => {
                  fn(c, offset, idx);
                  offset += c.nodeSize;
                });
              },
            };
          }
          return null;
        },
      };

      const mockHost = {
        doc: mockDoc as any,
        selection: {} as any,
        posOfPara: (_para: unknown) => 1,
      };

      // Floating model3d is index 0 of "drawing"
      const floatingModelPos = drawingNodePos(mockHost, {}, 0, "drawing");
      expect(floatingModelPos).toBe(1);

      // Inline model3d is index 0 of "inline"
      const inlineModelPos = drawingNodePos(mockHost, {}, 0, "inline");
      expect(inlineModelPos).toBe(2);

      // Floating ink is index 1 of "drawing"
      const floatingInkPos = drawingNodePos(mockHost, {}, 1, "drawing");
      expect(floatingInkPos).toBe(3);

      // Inline ink is index 1 of "inline"
      const inlineInkPos = drawingNodePos(mockHost, {}, 1, "inline");
      expect(inlineInkPos).toBe(4);
    });
  });

  describe("Aspect Ratio Preservation (Shift-lock) on Resize", () => {
    it("locks aspect ratio when lockAspectRatio is active", () => {
      const initialBox: Box = { x: 100, y: 100, width: 200, height: 100 }; // 2:1 ratio

      // Resize from bottom-right (se) with lock = true (Shift key pressed)
      const resized = resizeBox(initialBox, "se", 50, 80, 24, true);

      const newWidth = resized.width;
      const newHeight = resized.height;

      // Aspect ratio (width / height) should be preserved at 2:1
      expect(newWidth / newHeight).toBeCloseTo(2.0, 4);
    });
  });

  describe("Alt Text Task Pane & Delegation Notice", () => {
    it("displays exact delegation notice for 3D models and ink", () => {
      const pane = new DocenAltTextPane();
      document.body.appendChild(pane);

      // Set target to 3D model
      pane.setTarget({
        kind: "model3d",
        title: "Rover 3D",
        descr: "Mars rover 3D wireframe mesh",
      });

      expect(pane.showNotice).toBe(true);
      expect(pane.title).toBe("Rover 3D");
      expect(pane.description).toBe("Mars rover 3D wireframe mesh");
      expect(DELEGATION_NOTICE).toBe(
        "3D mesh vertex manipulation and ink vector splitting are delegated to external 3D/ink tools; viewing, layout, sizing, and alt text are editable natively.",
      );

      // Set target to ink
      pane.setTarget({
        kind: "ink",
        title: "Pen Annotation",
        descr: "Handwritten note with ink strokes",
      });

      expect(pane.showNotice).toBe(true);
      expect(pane.title).toBe("Pen Annotation");
      expect(pane.description).toBe("Handwritten note with ink strokes");

      // Set target to image (non 3D/ink)
      pane.setTarget({
        kind: "image",
        title: "Photo",
        descr: "Landscape photography",
      });

      expect(pane.showNotice).toBe(false);

      pane.remove();
    });

    it("emits alt-text:change and alt-text:apply events on input and submit", () => {
      const pane = new DocenAltTextPane();
      document.body.appendChild(pane);

      pane.setTarget({
        kind: "model3d",
        title: "Old Title",
        descr: "Old Descr",
      });

      const changeSpy = vi.fn();
      const applySpy = vi.fn();
      pane.addEventListener("alt-text:change", changeSpy);
      pane.addEventListener("alt-text:apply", applySpy);

      // Simulate title input change
      pane.handleTitleInput({ target: { value: "Updated 3D Model" } } as any);
      expect(changeSpy).toHaveBeenCalledWith(
        expect.objectContaining({
          detail: expect.objectContaining({
            title: "Updated 3D Model",
            kind: "model3d",
          }),
        }),
      );

      // Simulate descr input change
      pane.handleDescrInput({ target: { value: "Updated 3D description" } } as any);
      expect(changeSpy).toHaveBeenCalledWith(
        expect.objectContaining({
          detail: expect.objectContaining({
            descr: "Updated 3D description",
            kind: "model3d",
          }),
        }),
      );

      // Simulate apply
      pane.apply();
      expect(applySpy).toHaveBeenCalledWith(
        expect.objectContaining({
          detail: expect.objectContaining({
            title: "Updated 3D Model",
            descr: "Updated 3D description",
            kind: "model3d",
          }),
        }),
      );

      pane.remove();
    });
  });
});
