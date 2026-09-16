// @vitest-environment node
// Quick Parts domain orchestration — the ribbon→dialog→commit flow with a fake
// host view (the real one is the DOM-bound <docen-document> element).
import { describe, expect, it } from "vitest";

import type { BuildingBlock } from "../../building-blocks";
import {
  BuildingBlocksHostCommands,
  type BuildingBlocksHostView,
  type QuickPartCapture,
} from "./building-blocks";

const block = (id: string, name: string): BuildingBlock => ({
  id,
  name,
  gallery: "custQuickParts",
  category: "General",
  description: "",
  savedAt: 0,
  insertMode: "content",
  content: { openStart: 0, openEnd: 0, content: [{ type: "paragraph", text: name }] },
});

interface FakeHost {
  view: BuildingBlocksHostView;
  blocks: BuildingBlock[];
  inserted: string[];
  focuses: number;
  seed: unknown;
  capture: QuickPartCapture | null;
  editable: boolean;
}

const makeHost = (blocks: BuildingBlock[] = []): FakeHost => {
  const dialogs: Record<string, unknown> = {};
  const host: FakeHost = {
    blocks,
    inserted: [],
    focuses: 0,
    seed: undefined,
    capture: null,
    editable: true,
    view: {
      element: () =>
        ({
          shadowRoot: { querySelector: (tag: string) => dialogs[tag] ?? null },
        }) as unknown as HTMLElement,
      editable: () => host.editable,
      blocks: () => host.blocks,
      setBlocks: (next) => {
        host.blocks = [...next];
      },
      insertBlock: (id) => {
        host.inserted.push(id);
      },
      selectionSlice: () => host.capture,
      focusBridge: () => {
        host.focuses++;
      },
    },
  };
  dialogs["docen-quick-part-dialog"] = {
    show: (seed: unknown) => {
      host.seed = seed;
    },
  };
  dialogs["docen-building-blocks-dialog"] = {
    show: (seed: unknown) => {
      host.seed = seed;
    },
  };
  return host;
};

const capture: QuickPartCapture = {
  slice: { openStart: 0, openEnd: 0, content: [{ type: "paragraph", text: "Hello" }] },
  preview: "Hello",
  suggestedName: "Hello",
};

describe("BuildingBlocksHostCommands", () => {
  it("opens the save dialog seeded with the selection and current names", () => {
    const host = makeHost([block("a", "Existing")]);
    host.capture = capture;
    const domain = new BuildingBlocksHostCommands(host.view);
    expect(domain.run("save-quick-part")).toBe(true);
    expect(host.seed).toMatchObject({
      preview: "Hello",
      suggestedName: "Hello",
      existingNames: ["Existing"],
    });
  });

  it("does not open the save dialog without a selection or in Viewing mode", () => {
    const host = makeHost();
    const domain = new BuildingBlocksHostCommands(host.view);
    expect(domain.run("save-quick-part")).toBe(true);
    expect(host.seed).toBeUndefined();

    host.capture = capture;
    host.editable = false;
    expect(domain.run("save-quick-part")).toBe(true);
    expect(host.seed).toBeUndefined();
  });

  it("commits a save once, appending the captured content", () => {
    const host = makeHost();
    const domain = new BuildingBlocksHostCommands(host.view);
    host.capture = capture;
    domain.run("save-quick-part");
    domain.commitSave({
      name: "My Block",
      gallery: "custAutoTxt",
      category: "Letters",
      description: "d",
    });
    expect(host.blocks).toHaveLength(1);
    expect(host.blocks[0]).toMatchObject({
      name: "My Block",
      gallery: "custAutoTxt",
      category: "Letters",
      description: "d",
      content: capture.slice,
    });
    expect(host.focuses).toBe(1);
    // A second commit without a fresh capture is a no-op (the snapshot is
    // consumed by the first OK).
    domain.commitSave({
      name: "Again",
      gallery: "custQuickParts",
      category: "General",
      description: "",
    });
    expect(host.blocks).toHaveLength(1);
  });

  it("rejects a duplicate save name defensively", () => {
    const host = makeHost([block("a", "Taken")]);
    const domain = new BuildingBlocksHostCommands(host.view);
    host.capture = capture;
    domain.run("save-quick-part");
    domain.commitSave({
      name: "taken",
      gallery: "custQuickParts",
      category: "General",
      description: "",
    });
    expect(host.blocks).toHaveLength(1);
  });

  it("renames and deletes blocks, rejecting collisions", () => {
    const host = makeHost([block("a", "One"), block("b", "Two")]);
    const domain = new BuildingBlocksHostCommands(host.view);
    domain.renameBlock("a", "First");
    expect(host.blocks.map((b) => b.name)).toEqual(["First", "Two"]);

    domain.renameBlock("a", "two"); // collision with "Two" (case-insensitive)
    expect(host.blocks.map((b) => b.name)).toEqual(["First", "Two"]);

    domain.renameBlock("a", ""); // empty
    expect(host.blocks[0]!.name).toBe("First");

    domain.deleteBlock("a");
    expect(host.blocks.map((b) => b.id)).toEqual(["b"]);
    domain.deleteBlock("missing");
    expect(host.blocks).toHaveLength(1);
  });

  it("opens the organizer (also without an editor) and inserts by id", () => {
    const host = makeHost([block("a", "One")]);
    host.editable = false;
    const domain = new BuildingBlocksHostCommands(host.view);
    expect(domain.run("building-blocks-organizer")).toBe(true);
    expect(host.seed).toMatchObject({ editable: false });

    expect(domain.run("quick-parts", "a")).toBe(true);
    expect(host.inserted).toEqual(["a"]);
    expect(domain.run("quick-parts")).toBe(true);
    expect(host.inserted).toEqual(["a"]);
    expect(domain.run("unknown")).toBe(false);
  });
});
