// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest";

import DocenLayoutOptionsFlyout from "./layout-options-flyout";

describe("DocenLayoutOptionsFlyout", () => {
  it("initializes with default values", () => {
    const flyout = new DocenLayoutOptionsFlyout();
    expect(flyout.wrapMode).toBe("inline");
    expect(flyout.positionMode).toBe("moveWithText");
    expect(flyout.isFloating).toBe(false);
    expect(flyout.isOpen).toBe(false);
  });

  it("toggles open and closed", () => {
    const flyout = new DocenLayoutOptionsFlyout();
    const fakeEvent = {
      stopPropagation: vi.fn(),
      preventDefault: vi.fn(),
    } as unknown as MouseEvent;

    flyout.toggleOpen(fakeEvent);
    expect(flyout.isOpen).toBe(true);

    flyout.toggleOpen(fakeEvent);
    expect(flyout.isOpen).toBe(false);
  });

  it("emits select-wrap event when selecting a wrap mode", () => {
    const flyout = new DocenLayoutOptionsFlyout();
    const spy = vi.fn();
    flyout.addEventListener("select-wrap", spy);

    flyout.selectWrap("square");
    expect(flyout.wrapMode).toBe("square");
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.mock.calls[0][0].detail).toEqual({ wrap: "square" });
  });

  it("emits select-position event when selecting position mode", () => {
    const flyout = new DocenLayoutOptionsFlyout();
    const spy = vi.fn();
    flyout.addEventListener("select-position", spy);

    flyout.selectPosition("fixPosition");
    expect(flyout.positionMode).toBe("fixPosition");
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.mock.calls[0][0].detail).toEqual({ mode: "fixPosition" });
  });

  it("correctly identifies selected wrap mode including top-bottom aliases", () => {
    const flyout = new DocenLayoutOptionsFlyout();
    flyout.wrapMode = "square";
    expect(flyout.isWrapSelected("square")).toBe(true);
    expect(flyout.isWrapSelected("inline")).toBe(false);

    flyout.wrapMode = "topBottom";
    expect(flyout.isWrapSelected("top-bottom")).toBe(true);

    flyout.wrapMode = "top-bottom";
    expect(flyout.isWrapSelected("top-bottom")).toBe(true);
  });

  it("emits open-dialog on openMoreDialog()", () => {
    const flyout = new DocenLayoutOptionsFlyout();
    flyout.isOpen = true;
    const spy = vi.fn();
    flyout.addEventListener("open-dialog", spy);

    flyout.openMoreDialog();
    expect(flyout.isOpen).toBe(false);
    expect(spy).toHaveBeenCalledTimes(1);
  });
});
