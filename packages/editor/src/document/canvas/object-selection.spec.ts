// @vitest-environment happy-dom
// Select Objects mode: selection, Ctrl toggling, marquee, delete, and the
// Esc/empty-click stepdown — driven against a fake hit surface and a real
// editor (the delete transaction's observable result is the surviving text).
import { docxExtensions } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it, vi } from "vitest";

import type { DrawingHit } from "../../drawing";
import { ObjectSelectionMode, type ObjectSelectionHost } from "./object-selection";

const rect = (left: number, top: number, width = 600, height = 800): DOMRect =>
  ({ left, top, width, height, right: left + width, bottom: top + height }) as DOMRect;

const pageEl = (r: DOMRect): HTMLElement =>
  ({ getBoundingClientRect: () => r }) as unknown as HTMLElement;

interface Box extends DrawingHit {
  para: string;
}

const box = (
  para: string,
  page: number,
  x: number,
  y: number,
  w: number,
  h: number,
  kind: "drawing" | "inline" = "drawing",
): Box => ({
  para,
  page,
  index: 0,
  kind,
  x,
  y,
  width: w,
  height: h,
});

/** A fake hit surface with two pages; para ids double as PM positions. */
function makeHost(
  boxes: Box[],
  positions: Record<string, number>,
  editor?: Editor,
): ObjectSelectionHost & { overlayRect: DOMRect } {
  const overlayRect = rect(100, 50);
  const pages = new Map<number, HTMLElement>([
    [0, pageEl(rect(100, 50))],
    [1, pageEl(rect(740, 50))],
  ]);
  return {
    overlayRect,
    drawingAt: (page, lx, ly) =>
      boxes.find(
        (b) =>
          b.page === page && lx >= b.x && lx <= b.x + b.width && ly >= b.y && ly <= b.y + b.height,
      ) ?? null,
    allDrawingBoxes: () => boxes,
    resolveBox: (hit) =>
      boxes.find(
        (b) =>
          b.para === hit.para &&
          b.index === hit.index &&
          b.kind === hit.kind &&
          b.page === hit.page,
      ) ?? null,
    nodePosOf: (hit) => positions[hit.para as string] ?? null,
    pageHost: (page) => pages.get(page) ?? null,
    overlayHost: () => ({ getBoundingClientRect: () => overlayRect }) as unknown as HTMLElement,
    scale: () => 1,
    editor: () =>
      editor ??
      ({
        state: { doc: { nodeAt: () => null }, tr: { delete: vi.fn() } },
        view: { dispatch: vi.fn() },
      } as never),
  };
}

const press = (over: Partial<Parameters<ObjectSelectionMode["press"]>[0]> = {}) => ({
  page: 0,
  lx: 10,
  ly: 10,
  clientX: 110,
  clientY: 60,
  ctrlKey: false,
  metaKey: false,
  ...over,
});

describe("ObjectSelectionMode", () => {
  const A = box("A", 0, 10, 10, 100, 50);
  const B = box("B", 0, 200, 10, 100, 50);
  const C = box("C", 1, 10, 10, 100, 50);
  const inline = box("I", 0, 10, 100, 100, 20, "inline");

  it("selects the clicked object and switches selection on a plain click", () => {
    const mode = new ObjectSelectionMode(makeHost([A, B], { A: 0, B: 3 }));
    mode.setActive(true);
    expect(mode.press(press({ lx: 20, ly: 20, clientX: 120, clientY: 70 }))).toBe(true);
    expect(mode.count).toBe(1);
    // A frame is painted for the selection.
    expect(mode.el.querySelectorAll(".docen-object-selection-frame")).toHaveLength(1);

    mode.press(press({ lx: 210, ly: 20, clientX: 310, clientY: 70 }));
    expect(mode.count).toBe(1);
    expect(mode.el.querySelectorAll(".docen-object-selection-frame")).toHaveLength(1);
  });

  it("toggles membership with Ctrl+click", () => {
    const mode = new ObjectSelectionMode(makeHost([A, B], { A: 0, B: 3 }));
    mode.setActive(true);
    mode.press(press({ lx: 20, ly: 20 }));
    mode.press(press({ lx: 210, ly: 20, clientX: 310, ctrlKey: true }));
    expect(mode.count).toBe(2);
    mode.press(press({ lx: 210, ly: 20, clientX: 310, ctrlKey: true }));
    expect(mode.count).toBe(1);
  });

  it("marquees floating objects, skipping inline art", () => {
    const mode = new ObjectSelectionMode(makeHost([A, B, C, inline], { A: 0, B: 3 }));
    mode.setActive(true);
    mode.press(press({ lx: 500, ly: 500, clientX: 110, clientY: 60 }));
    mode.move(320, 90);
    expect(mode.dragging).toBe(true);
    expect(mode.el.querySelector(".docen-object-selection-marquee")).not.toBeNull();
    expect(mode.release()).toBe(true);
    // A and B intersect the band; C is on page 1 and the inline image is
    // excluded.
    expect(mode.count).toBe(2);
    expect(mode.el.querySelectorAll(".docen-object-selection-frame")).toHaveLength(2);
  });

  it("clears the selection on an empty click and exits when nothing is selected", () => {
    const mode = new ObjectSelectionMode(makeHost([A], { A: 0 }));
    mode.setActive(true);
    // Empty click with no selection leaves the mode.
    mode.press(press({ lx: 500, ly: 500, clientX: 600, clientY: 550 }));
    mode.release();
    expect(mode.active).toBe(false);

    mode.setActive(true);
    mode.press(press({ lx: 20, ly: 20 }));
    mode.press(press({ lx: 500, ly: 500, clientX: 600, clientY: 550 }));
    mode.release();
    expect(mode.active).toBe(true);
    expect(mode.count).toBe(0);
  });

  it("Esc clears the selection first, then leaves the mode", () => {
    const mode = new ObjectSelectionMode(makeHost([A], { A: 0 }));
    mode.setActive(true);
    mode.press(press({ lx: 20, ly: 20 }));
    expect(mode.escape()).toBe(true);
    expect(mode.active).toBe(true);
    expect(mode.count).toBe(0);
    expect(mode.escape()).toBe(true);
    expect(mode.active).toBe(false);
  });

  it("deletes the selected objects in one transaction", () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          { type: "paragraph", content: [{ type: "text", text: "A" }] },
          { type: "paragraph", content: [{ type: "text", text: "B" }] },
          { type: "paragraph", content: [{ type: "text", text: "C" }] },
        ],
      },
    });
    const mode = new ObjectSelectionMode(makeHost([A, C], { A: 0, C: 6 }, editor));
    mode.setActive(true);
    mode.press(press({ lx: 20, ly: 20 }));
    mode.press(press({ lx: 20, ly: 20, page: 1, clientX: 760, ctrlKey: true }));
    expect(mode.deleteSelection()).toBe(true);
    expect(editor.state.doc.textContent).toBe("B");
    expect(mode.count).toBe(0);
  });

  it("drops selections whose boxes no longer paint after a refresh", () => {
    const boxes = [A];
    const host = makeHost(boxes, { A: 0 });
    const mode = new ObjectSelectionMode(host);
    mode.setActive(true);
    mode.press(press({ lx: 20, ly: 20 }));
    boxes.length = 0;
    mode.refresh();
    expect(mode.count).toBe(0);
  });

  it("turning the mode off clears the selection", () => {
    const mode = new ObjectSelectionMode(makeHost([A], { A: 0 }));
    mode.setActive(true);
    mode.press(press({ lx: 20, ly: 20 }));
    mode.setActive(false);
    expect(mode.count).toBe(0);
    expect(mode.el.style.display).toBe("none");
  });

  it("ignores presses while inactive", () => {
    const mode = new ObjectSelectionMode(makeHost([A], { A: 0 }));
    expect(mode.press(press())).toBe(false);
    expect(mode.count).toBe(0);
  });
});
