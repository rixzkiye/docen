// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest";

import type { Box } from "./geometry";
import { NodeEditOverlay } from "./node-edit-overlay";

describe("NodeEditOverlay", () => {
  const sampleCustomGeometry = {
    pathList: [
      {
        w: 100,
        h: 100,
        commands: [
          { command: "moveTo", point: { x: "0", y: "0" } },
          { command: "lnTo", point: { x: "100", y: "0" } },
          { command: "lnTo", point: { x: "50", y: "100" } },
          { command: "close" },
        ],
      },
    ],
  };

  const box: Box = { x: 50, y: 50, width: 200, height: 200 };

  it("renders overlay and vertex handles when shown", () => {
    const applySpy = vi.fn();
    const exitSpy = vi.fn();
    const overlay = new NodeEditOverlay({
      scale: () => 1,
      applyCustomGeometry: applySpy,
      onExit: exitSpy,
    });
    document.body.appendChild(overlay.el);

    expect(overlay.active).toBe(false);
    expect(overlay.el.style.display).toBe("none");

    overlay.show(box, 0, sampleCustomGeometry);

    expect(overlay.active).toBe(true);
    expect(overlay.el.style.display).toBe("block");

    const handles = overlay.el.querySelectorAll("[data-vertex-handle]");
    expect(handles.length).toBe(3); // (0,0), (100,0), (50,100)

    const path = overlay.el.querySelector("path");
    expect(path).not.toBeNull();
    expect(path?.getAttribute("d")).toBe("M 0 0 L 200 0 L 100 200 Z");

    overlay.hide();
    expect(overlay.active).toBe(false);
    expect(exitSpy).toHaveBeenCalled();
    overlay.el.remove();
  });

  it("handles bezier curve vertices properly", () => {
    const bezierCg = {
      pathList: [
        {
          w: 100,
          h: 100,
          commands: [
            { command: "moveTo", point: { x: "0", y: "0" } },
            {
              command: "cubicBezTo",
              points: [
                { x: "25", y: "50" },
                { x: "75", y: "50" },
                { x: "100", y: "0" },
              ],
            },
            { command: "close" },
          ],
        },
      ],
    };

    const overlay = new NodeEditOverlay({
      scale: () => 1,
      applyCustomGeometry: vi.fn(),
    });
    document.body.appendChild(overlay.el);

    overlay.show(box, 0, bezierCg);

    // 1 start point + 3 cubic control/end points = 4 handles
    const handles = overlay.el.querySelectorAll("[data-vertex-handle]");
    expect(handles.length).toBe(4);

    overlay.hide();
    overlay.el.remove();
  });

  it("commits on Enter key and cancels on Escape key", () => {
    const applySpy = vi.fn();
    const exitSpy = vi.fn();
    const overlay = new NodeEditOverlay({
      scale: () => 1,
      applyCustomGeometry: applySpy,
      onExit: exitSpy,
    });
    document.body.appendChild(overlay.el);

    // Test Enter key commit
    overlay.show(box, 0, sampleCustomGeometry);
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter" }));
    expect(applySpy).toHaveBeenCalledWith(expect.objectContaining({ pathList: expect.any(Array) }));
    expect(overlay.active).toBe(false);

    // Test Escape key cancel
    applySpy.mockClear();
    overlay.show(box, 0, sampleCustomGeometry);
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(applySpy).not.toHaveBeenCalled();
    expect(overlay.active).toBe(false);

    overlay.el.remove();
  });

  it("commits when clicking outside the overlay", () => {
    const applySpy = vi.fn();
    const overlay = new NodeEditOverlay({
      scale: () => 1,
      applyCustomGeometry: applySpy,
    });
    document.body.appendChild(overlay.el);

    overlay.show(box, 0, sampleCustomGeometry);
    expect(overlay.active).toBe(true);

    // Click outside
    const outsideEl = document.createElement("div");
    document.body.appendChild(outsideEl);
    outsideEl.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true }));

    expect(applySpy).toHaveBeenCalled();
    expect(overlay.active).toBe(false);

    outsideEl.remove();
    overlay.el.remove();
  });

  it("allows dragging a vertex handle to mutate custom geometry", () => {
    const applySpy = vi.fn();
    const overlay = new NodeEditOverlay({
      scale: () => 1,
      applyCustomGeometry: applySpy,
    });
    document.body.appendChild(overlay.el);

    overlay.show(box, 0, sampleCustomGeometry);

    const handles = overlay.el.querySelectorAll<HTMLDivElement>("[data-vertex-handle]");
    expect(handles.length).toBe(3);

    // Handle 0 is at (0, 0). Drag it by (+20px, +40px)
    // Box width=200, path width=100 -> sx = 200/100 = 2
    // dx = 20px / 2 = 10 path units; dy = 40px / 2 = 20 path units
    const handle0 = handles[0]!;
    handle0.dispatchEvent(
      new PointerEvent("pointerdown", {
        bubbles: true,
        clientX: 0,
        clientY: 0,
      }),
    );

    window.dispatchEvent(
      new PointerEvent("pointermove", {
        bubbles: true,
        clientX: 20,
        clientY: 40,
      }),
    );

    window.dispatchEvent(
      new PointerEvent("pointerup", {
        bubbles: true,
        clientX: 20,
        clientY: 40,
      }),
    );

    expect(applySpy).toHaveBeenCalled();
    const lastCall = applySpy.mock.calls[applySpy.mock.calls.length - 1]![0];
    expect(lastCall.pathList[0].commands[0].point.x).toBe("10");
    expect(lastCall.pathList[0].commands[0].point.y).toBe("20");

    overlay.hide();
    overlay.el.remove();
  });
});
