import { describe, expect, it, vi } from "vitest";

import { RevisionsHostCommands, type RevisionsHostView } from "./revisions";

/** A fake host view capturing the display-state writes. */
function fakeHost(): {
  host: RevisionsHostView;
  balloons: string[];
  renders: () => number;
  syncs: () => number;
} {
  let rendered = 0;
  let synced = 0;
  const balloons: string[] = [];
  const host: RevisionsHostView = {
    editor: () => ({}) as never,
    togglePane: vi.fn(),
    setMarkupView: vi.fn(),
    getMarkupAuthors: () => null,
    setMarkupAuthors: vi.fn(),
    setMarkupColors: vi.fn(),
    setBalloons: (mode) => balloons.push(mode),
    renderDoc: () => {
      rendered++;
    },
    syncMarkupMenus: () => {
      synced++;
    },
    getJSON: () => ({ type: "doc", content: [] }),
  };
  return { host, balloons, renders: () => rendered, syncs: () => synced };
}

describe("RevisionsHostCommands show-markup", () => {
  it("stores the balloons mode and re-projects the document", () => {
    const { host, balloons, renders, syncs } = fakeHost();
    const commands = new RevisionsHostCommands(host);
    expect(commands.run("show-markup", "comments")).toBe(true);
    expect(balloons).toEqual(["comments"]);
    expect(renders()).toBe(1);
    expect(syncs()).toBe(1);
  });

  it("accepts every documented mode and ignores anything else", () => {
    const { host, balloons } = fakeHost();
    const commands = new RevisionsHostCommands(host);
    for (const mode of ["all", "revisions", "none"]) {
      expect(commands.run("show-markup", mode)).toBe(true);
    }
    expect(balloons).toEqual(["all", "revisions", "none"]);
    commands.run("show-markup", "balloons-please");
    commands.run("show-markup");
    expect(balloons).toEqual(["all", "revisions", "none"]);
  });

  it("declares show-markup as a wired editor event", () => {
    const { host } = fakeHost();
    expect(new RevisionsHostCommands(host).editor).toContain("show-markup");
  });
});
