import { describe, expect, it, vi } from "vitest";

import { DesignHostCommands, type DesignHostView } from "./design";

/** A fake host view capturing the design-command writes. */
function fakeHost(): {
  host: DesignHostView;
  setDefaults: () => number;
  restoreStyles: () => number;
} {
  let defaults = 0;
  let restores = 0;
  const host: DesignHostView = {
    setPageColor: vi.fn(),
    setParagraphSpacing: vi.fn(),
    openWatermarkDialog: vi.fn(),
    setWatermark: vi.fn(),
    openFillEffectsDialog: vi.fn(),
    setAsDefault: () => {
      defaults++;
    },
    restoreStylesSnapshot: () => {
      restores++;
    },
  };
  return { host, setDefaults: () => defaults, restoreStyles: () => restores };
}

describe("DesignHostCommands set-default", () => {
  it("declares set-default as a wired editor event", () => {
    const { host } = fakeHost();
    expect(new DesignHostCommands(host).editor).toContain("set-default");
  });

  it("persists the document defaults and consumes the event", () => {
    const { host, setDefaults } = fakeHost();
    const commands = new DesignHostCommands(host);
    expect(commands.run("set-default")).toBe(true);
    expect(setDefaults()).toBe(1);
  });

  it("still routes the style-set default entry to the styles snapshot", () => {
    const { host, restoreStyles } = fakeHost();
    const commands = new DesignHostCommands(host);
    expect(commands.run("style-set", "default")).toBe(true);
    expect(restoreStyles()).toBe(1);
    // A real preset falls through to the wired Tiptap command.
    expect(commands.run("style-set", "modern")).toBe(false);
  });
});
