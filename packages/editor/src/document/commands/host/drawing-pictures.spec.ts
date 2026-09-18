// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest";

// leafer's module graph touches these globals at import time (the domain
// imports pinImage from @docen/core) — stub them BEFORE the dynamic import.
(globalThis as any).CanvasRenderingContext2D ??= class {};
(globalThis as any).Path2D ??= class {};
(globalThis as any).ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
};
const fakeCtx = () =>
  new Proxy({ measureText: (s: string) => ({ width: s.length * 8 }) } as Record<string, unknown>, {
    get: (t, p: string) => (p in t ? t[p] : () => {}),
    set: () => true,
  });
const canvasProto = (globalThis as any).HTMLCanvasElement?.prototype;
if (canvasProto) canvasProto.getContext = () => fakeCtx();

const { DrawingPicturesHostCommands } = await import("./drawing-pictures");
type DrawingPicturesHostView = import("./drawing-pictures").DrawingPicturesHostView;

/** A fake host view — only the online-picture surface is exercised. */
function fakeHost(): { host: DrawingPicturesHostView; opens: () => number } {
  let opens = 0;
  const host = {
    editor: () => ({}) as never,
    activeEditor: () => ({}) as never,
    element: () => ({}) as HTMLElement,
    showCompressPictures: vi.fn(),
    armTransparentPick: vi.fn(),
    drawingMulti: () => null,
    pickImage: vi.fn(),
    openOnlinePictures: () => {
      opens++;
    },
    pickPicture: vi.fn(),
    focusBridge: vi.fn(),
    drawingState: () => null,
    enterCropMode: vi.fn(),
    insertShapeAt: vi.fn(),
    armShapeDrawer: vi.fn(),
    insertWordArt: vi.fn(),
  } as unknown as DrawingPicturesHostView;
  return { host, opens: () => opens };
}

describe("DrawingPicturesHostCommands online-picture", () => {
  it("declares online-picture as a wired editor event", () => {
    const { host } = fakeHost();
    expect(new DrawingPicturesHostCommands(host).editor).toContain("online-picture");
  });

  it("opens the address dialog for the ribbon event", () => {
    const { host, opens } = fakeHost();
    expect(new DrawingPicturesHostCommands(host).run("online-picture")).toBe(true);
    expect(opens()).toBe(1);
  });
});
