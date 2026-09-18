import { describe, expect, it, vi } from "vitest";

import { ClipboardFormatHostCommands, type ClipboardFormatHostView } from "./clipboard-format";

/** A fake host view — the Select Objects surface is what these tests exercise. */
function fakeHost(): { host: ClipboardFormatHostView; toggles: () => number } {
  let toggles = 0;
  const host = {
    editor: () => ({}) as never,
    activeEditor: () => ({}) as never,
    element: () => ({}) as HTMLElement,
    copySelection: vi.fn(),
    paste: vi.fn(),
    togglePane: vi.fn(),
    showTaskpane: vi.fn(),
    renderStylesPane: vi.fn(),
    toggleMarkdownInput: vi.fn(),
    syncFormatButtons: vi.fn(),
    insertLink: vi.fn(),
    hrefAtCaret: () => null,
    jumpToBookmark: vi.fn(),
    select: vi.fn(),
    toggleObjectSelect: () => {
      toggles++;
    },
    toggleFormatPainter: vi.fn(),
  } as ClipboardFormatHostView;
  return { host, toggles: () => toggles };
}

describe("ClipboardFormatHostCommands select-objects", () => {
  it("declares select-objects as a wired editor event", () => {
    const { host } = fakeHost();
    expect(new ClipboardFormatHostCommands(host).editor).toContain("select-objects");
  });

  it("toggles the object-selection mode", () => {
    const { host, toggles } = fakeHost();
    const commands = new ClipboardFormatHostCommands(host);
    expect(commands.run("select-objects")).toBe(true);
    expect(commands.run("select-objects")).toBe(true);
    expect(toggles()).toBe(2);
  });
});
