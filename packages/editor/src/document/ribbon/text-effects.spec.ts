// @vitest-environment happy-dom
// The Text Effects gallery wiring: every menu token resolves to a known
// preset (or the documented :0 clear / :options dialog entries), and both
// ribbon surfaces (Home → Font, Design → Document Formatting) carry it.
import { describe, expect, it } from "vitest";

import { presetXml, TEXT_EFFECT_PRESETS } from "../text-effects";
import { designTab } from "./design";
import { homeTab } from "./home";
import { textEffectsItems } from "./text-effects";

interface MenuItem {
  text: string;
  event?: string;
  value?: string;
  children?: MenuItem[];
}

const walk = (items: MenuItem[]): MenuItem[] =>
  items.flatMap((item) => [item, ...(item.children ? walk(item.children) : [])]);

describe("textEffectsItems", () => {
  const items = JSON.parse(textEffectsItems()) as MenuItem[];

  it("offers the six Word effect families plus the options entry", () => {
    expect(items.map((item) => item.text)).toEqual([
      "ribbon.opt.te-outline",
      "ribbon.opt.te-shadow",
      "ribbon.opt.te-reflection",
      "ribbon.opt.te-glow",
      "ribbon.opt.te-bevel",
      "ribbon.opt.te-rotation",
      "ribbon.opt.te-options",
    ]);
    // Every family is a nested submenu with its presets.
    for (const family of items.slice(0, 6)) {
      expect(family.children?.length).toBeGreaterThanOrEqual(4);
    }
  });

  it("maps every token to a known preset, a :0 clear, or the options dialog", () => {
    for (const item of walk(items)) {
      if (!item.value) continue;
      expect(item.event).toBe("text-effects");
      if (item.value === "options") continue;
      if (item.value.endsWith(":options")) continue;
      if (item.value.endsWith(":0")) continue;
      expect(presetXml(item.value), item.value).toBeTruthy();
      expect(TEXT_EFFECT_PRESETS[item.value], item.value).toBeDefined();
    }
  });
});

describe("ribbon placement", () => {
  const controlEvents = (tab: ReturnType<typeof homeTab>): Array<string | undefined> => {
    const walkControls = (controls: readonly { event?: string; controls?: unknown }[]): string[] =>
      controls.flatMap((control) => [
        ...(control.event ? [control.event] : []),
        ...(Array.isArray(control.controls)
          ? walkControls(control.controls as { event?: string; controls?: unknown }[])
          : []),
      ]);
    return tab.groups.flatMap((group) => walkControls(group.controls as never));
  };

  it("sits in the Home Font group with a gallery menu", () => {
    expect(controlEvents(homeTab())).toContain("text-effects");
  });

  it("sits in the Design Document Formatting group", () => {
    expect(controlEvents(designTab())).toContain("effects");
  });
});
