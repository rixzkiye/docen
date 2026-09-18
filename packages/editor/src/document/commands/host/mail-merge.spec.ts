import { describe, expect, it, vi } from "vitest";

import { MailMergeHostCommands, type MailMergeHostView } from "./mail-merge";

/** A fake host view capturing the merge-state writes. */
function fakeHost(): { host: MailMergeHostView; highlightToggles: () => number } {
  let toggles = 0;
  const host: MailMergeHostView = {
    element: () => ({ shadowRoot: null }) as unknown as HTMLElement,
    recipients: () => null,
    insertAddressBlock: vi.fn(),
    insertGreetingLine: vi.fn(),
    togglePreview: vi.fn(),
    firstRecord: vi.fn(),
    lastRecord: vi.fn(),
    setMergeType: vi.fn(),
    toggleHighlightMergeFields: () => {
      toggles++;
    },
    finishMerge: vi.fn(),
  };
  return { host, highlightToggles: () => toggles };
}

describe("MailMergeHostCommands highlight-merge", () => {
  it("declares highlight-merge as a wired editor event", () => {
    const { host } = fakeHost();
    expect(new MailMergeHostCommands(host).editor).toContain("highlight-merge");
  });

  it("flips the host's Highlight Merge Fields flag and consumes the event", () => {
    const { host, highlightToggles } = fakeHost();
    const commands = new MailMergeHostCommands(host);
    expect(commands.run("highlight-merge")).toBe(true);
    expect(commands.run("highlight-merge")).toBe(true);
    expect(highlightToggles()).toBe(2);
  });

  it("leaves unknown events to the wired dispatch", () => {
    const { host, highlightToggles } = fakeHost();
    const commands = new MailMergeHostCommands(host);
    expect(commands.run("not-a-merge-command")).toBe(false);
    expect(highlightToggles()).toBe(0);
  });
});
