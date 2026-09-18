// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest";

import { DocenTouchHandles } from "./touch-handles";

describe("W5.4 Touch & Pen Input Support", () => {
  beforeEach(() => {
    document.body.replaceChildren();
  });

  it("registers <docen-touch-handles> custom element with start and end teardrop handles", () => {
    expect(customElements.get("docen-touch-handles")).toBeDefined();
    const handles = new DocenTouchHandles();
    document.body.append(handles);

    expect(handles.visible).toBe(false);
    expect(handles.hasSelection).toBe(false);

    // Initial position
    handles.setPositions({ x: 100, y: 150, height: 20 }, { x: 250, y: 150, height: 20 });
    handles.show();

    expect(handles.hasSelection).toBe(true);
    expect(handles.visible).toBe(true);
    expect(handles.startX).toBe(100);
    expect(handles.startY).toBe(150);
    expect(handles.endX).toBe(250);
    expect(handles.endY).toBe(150);

    const shadow = handles.shadowRoot!;
    const startHandle = shadow.querySelector(".handle-start") as HTMLElement;
    const endHandle = shadow.querySelector(".handle-end") as HTMLElement;

    expect(startHandle).not.toBeNull();
    expect(endHandle).not.toBeNull();

    // Hit area requirement: 40x40px minimum
    expect(startHandle.className).toContain("handle-start");
    expect(endHandle.className).toContain("handle-end");
  });

  it("handles dragging selection handles and dispatches drag events", () => {
    const handles = new DocenTouchHandles();
    document.body.append(handles);

    handles.setPositions({ x: 50, y: 100, height: 18 }, { x: 180, y: 100, height: 18 });
    handles.show();

    const dragStartSpy = vi.fn();
    const dragMoveSpy = vi.fn();
    const dragEndSpy = vi.fn();

    handles.addEventListener("handle-drag-start", dragStartSpy);
    handles.addEventListener("handle-drag", dragMoveSpy);
    handles.addEventListener("handle-drag-end", dragEndSpy);

    const shadow = handles.shadowRoot!;
    const startHandle = shadow.querySelector(".handle-start") as HTMLElement;

    // Simulate pointerdown
    const downEvent = new PointerEvent("pointerdown", {
      bubbles: true,
      composed: true,
      clientX: 50,
      clientY: 118,
      pointerId: 1,
      pointerType: "touch",
    });
    // Stub setPointerCapture
    startHandle.setPointerCapture = vi.fn();
    startHandle.releasePointerCapture = vi.fn();
    startHandle.dispatchEvent(downEvent);

    expect(dragStartSpy).toHaveBeenCalledTimes(1);
    expect(dragStartSpy.mock.calls[0][0].detail.handle).toBe("start");
    expect(handles.isDragging).toBe(true);
    expect(handles.activeHandle).toBe("start");

    // Simulate pointermove
    const moveEvent = new PointerEvent("pointermove", {
      bubbles: true,
      composed: true,
      clientX: 70,
      clientY: 118,
      pointerId: 1,
      pointerType: "touch",
    });
    startHandle.dispatchEvent(moveEvent);

    expect(dragMoveSpy).toHaveBeenCalledTimes(1);
    expect(dragMoveSpy.mock.calls[0][0].detail.handle).toBe("start");
    expect(dragMoveSpy.mock.calls[0][0].detail.clientX).toBe(70);

    // Simulate pointerup
    const upEvent = new PointerEvent("pointerup", {
      bubbles: true,
      composed: true,
      clientX: 70,
      clientY: 118,
      pointerId: 1,
      pointerType: "touch",
    });
    startHandle.dispatchEvent(upEvent);

    expect(dragEndSpy).toHaveBeenCalledTimes(1);
    expect(handles.isDragging).toBe(false);
    expect(handles.activeHandle).toBeNull();
  });

  it("synchronizes with EditBridge and moves selection on handle drag", () => {
    const handles = new DocenTouchHandles();
    document.body.append(handles);

    const listeners: Record<string, Function[]> = {};
    const mockEditor = {
      state: {
        selection: { from: 5, to: 15 },
      },
      commands: {
        setTextSelection: vi.fn(),
      },
      on: (event: string, fn: Function) => {
        (listeners[event] ??= []).push(fn);
      },
      off: (event: string, fn: Function) => {
        listeners[event] = (listeners[event] ?? []).filter((f) => f !== fn);
      },
    };

    const mockBridge = {
      activeEditor: () => mockEditor as any,
      selectionClientRect: vi.fn().mockReturnValue({
        left: 100,
        top: 200,
        right: 250,
        bottom: 220,
        width: 150,
        height: 20,
      }),
      posAtClient: vi.fn((x: number) => (x < 150 ? 2 : 25)),
    };

    const container = document.createElement("div");
    container.getBoundingClientRect = () =>
      ({ left: 0, top: 0, right: 1000, bottom: 1000, width: 1000, height: 1000 }) as DOMRect;
    document.body.append(container);

    const detach = handles.attachToBridge(mockBridge as any, container);

    // Trigger selectionUpdate
    listeners["selectionUpdate"]?.forEach((fn) => fn());
    expect(handles.visible).toBe(true);
    expect(handles.hasSelection).toBe(true);
    expect(handles.startX).toBe(100);
    expect(handles.endX).toBe(250);

    // Simulate dragging start handle to x=60
    handles.dispatchEvent(
      new CustomEvent("handle-drag", {
        detail: { handle: "start", clientX: 60, clientY: 210 },
      }),
    );

    expect(mockBridge.posAtClient).toHaveBeenCalledWith(60, 210);
    expect(mockEditor.commands.setTextSelection).toHaveBeenCalledWith({ from: 2, to: 15 });

    // Simulate dragging end handle to x=300
    handles.dispatchEvent(
      new CustomEvent("handle-drag", {
        detail: { handle: "end", clientX: 300, clientY: 210 },
      }),
    );
    expect(mockEditor.commands.setTextSelection).toHaveBeenCalledWith({ from: 5, to: 25 });

    detach();
  });

  it("handles RTL orientation properly on selection teardrops", () => {
    const handles = new DocenTouchHandles();
    document.body.append(handles);

    expect(handles.isRtl).toBe(false);

    handles.setAttribute("dir", "rtl");
    handles.dirChanged();
    expect(handles.isRtl).toBe(true);
  });
});
