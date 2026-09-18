// @vitest-environment node
import type { LayoutDrawingMember } from "@docen/layout";
import { describe, expect, it, vi } from "vitest";

vi.mock("leafer-ui", () => {
  class StubElement {
    children: StubElement[] = [];
    constructor(attrs: Record<string, unknown> = {}) {
      Object.assign(this, attrs);
    }
    add(...items: StubElement[]): void {
      this.children.push(...items);
    }
  }
  const make = () => class extends StubElement {};
  return {
    Box: make(),
    Ellipse: make(),
    Group: make(),
    Image: make(),
    ImageManager: make(),
    Line: make(),
    Path: make(),
    Rect: make(),
    Resource: make(),
    Text: make(),
  };
});

import { Group } from "leafer-ui";

import { paint3DModelMember, paintInkMember } from "./3d-ink";

/** Every painted element, depth first. */
const flatten = (group: Group): Record<string, unknown>[] => {
  const out: Record<string, unknown>[] = [];
  const walk = (node: { children?: unknown[] }): void => {
    for (const child of node.children ?? []) {
      out.push(child as Record<string, unknown>);
      walk(child as { children?: unknown[] });
    }
  };
  walk(group);
  return out;
};

describe("3D Model & Ink Paint Rendering", () => {
  it("paints 3D Model member wireframe, badge and alt text accurately", () => {
    const tree = new Group();
    const member: LayoutDrawingMember = {
      kind: "model3d",
      model3d: {},
      x: 50,
      y: 50,
      width: 150,
      height: 150,
      title: "3D Satellite",
      descr: "Satellite isometric model",
      rotation: 20,
      camera: { rot: { lat: 30, lon: 45, rev: 0 } },
    };

    paint3DModelMember(tree, member);

    expect(tree.children.length).toBeGreaterThan(0);
    const elements = flatten(tree);
    // Find badge text and title text
    const texts = elements.filter((e) => "text" in e).map((e) => (e as { text: string }).text);
    expect(texts).toContain("3D MODEL");
    expect(texts).toContain("3D Satellite");
  });

  it("paints Ink member cursive stroke paths and pen nib accurately", () => {
    const tree = new Group();
    const member: LayoutDrawingMember = {
      kind: "ink",
      ink: {},
      x: 20,
      y: 20,
      width: 140,
      height: 90,
      title: "Signature Ink",
      descr: "Handwritten cursive ink strokes",
    };

    paintInkMember(tree, member);

    expect(tree.children.length).toBeGreaterThan(0);
    const elements = flatten(tree);
    const texts = elements.filter((e) => "text" in e).map((e) => (e as { text: string }).text);
    expect(texts).toContain("INK");
    expect(texts).toContain("Signature Ink");
  });
});
