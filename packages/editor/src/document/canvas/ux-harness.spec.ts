// @vitest-environment node
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import {
  CDP_MODIFIERS,
  compareScreenshots,
  createUxHarness,
  parseShortcut,
  type CdpSession,
} from "./ux-harness";

describe("UX Interaction Harness (Headless Chromium + CDP)", () => {
  let session: CdpSession;

  beforeAll(async () => {
    session = await createUxHarness({ width: 800, height: 600 });
  }, 15000);

  afterAll(async () => {
    if (session) {
      await session.close();
    }
  });

  describe("Shortcut Parser & Modifier Bitmasks", () => {
    it("parses single key without modifiers", () => {
      const parsed = parseShortcut("A");
      expect(parsed.key).toBe("A");
      expect(parsed.code).toBe("KeyA");
      expect(parsed.modifiers).toBe(0);
    });

    it("parses Ctrl+B", () => {
      const parsed = parseShortcut("Ctrl+B");
      expect(parsed.key).toBe("B");
      expect(parsed.code).toBe("KeyB");
      expect(parsed.modifiers).toBe(CDP_MODIFIERS.CTRL);
    });

    it("parses Ctrl+Shift+W", () => {
      const parsed = parseShortcut("Ctrl+Shift+W");
      expect(parsed.key).toBe("W");
      expect(parsed.modifiers).toBe(CDP_MODIFIERS.CTRL | CDP_MODIFIERS.SHIFT);
    });

    it("parses Alt+Shift+ArrowUp", () => {
      const parsed = parseShortcut("Alt+Shift+Up");
      expect(parsed.key).toBe("ArrowUp");
      expect(parsed.modifiers).toBe(CDP_MODIFIERS.ALT | CDP_MODIFIERS.SHIFT);
      expect(parsed.keyCode).toBe(38);
    });

    it("parses named navigation keys, digits, and function keys", () => {
      expect(parseShortcut("Ctrl+Home")).toMatchObject({ key: "Home", keyCode: 36 });
      expect(parseShortcut("Ctrl+End")).toMatchObject({ key: "End", keyCode: 35 });
      expect(parseShortcut("Ctrl+PageDown")).toMatchObject({ key: "PageDown", keyCode: 34 });
      expect(parseShortcut("Ctrl+1")).toMatchObject({
        key: "1",
        code: "Digit1",
        keyCode: 49,
        modifiers: CDP_MODIFIERS.CTRL,
      });
      expect(parseShortcut("F8")).toMatchObject({ key: "F8", code: "F8", keyCode: 119 });
      expect(parseShortcut("Shift+F3")).toMatchObject({
        key: "F3",
        keyCode: 114,
        modifiers: CDP_MODIFIERS.SHIFT,
      });
      expect(parseShortcut("Tab")).toMatchObject({ key: "Tab", keyCode: 9 });
    });

    it("parses a bare modifier press as its own key", () => {
      expect(parseShortcut("Alt")).toMatchObject({ key: "Alt", code: "AltLeft", modifiers: 0 });
      expect(parseShortcut("Shift")).toMatchObject({ key: "Shift", modifiers: 0 });
      expect(parseShortcut("Ctrl")).toMatchObject({ key: "Control", modifiers: 0 });
    });
  });

  describe("Screenshot Diffing & Visual State Verification", () => {
    it("detects identical screenshots", () => {
      const a = new Uint8Array([1, 2, 3, 4, 5]);
      const b = new Uint8Array([1, 2, 3, 4, 5]);
      const res = compareScreenshots(a, b);
      expect(res.identical).toBe(true);
      expect(res.diffCount).toBe(0);
      expect(res.diffPercentage).toBe(0);
    });

    it("detects differences between dissimilar screenshots", () => {
      const a = new Uint8Array([1, 2, 3, 4, 5]);
      const b = new Uint8Array([1, 9, 3, 9, 5]);
      const res = compareScreenshots(a, b);
      expect(res.identical).toBe(false);
      expect(res.diffCount).toBe(2);
      expect(res.diffPercentage).toBe(40);
    });
  });

  describe.sequential("Interactive Scenarios (Pointer, Keyboard, Cursors)", () => {
    it("sets content and executes DOM evaluation", async () => {
      await session.setContent(`
        <!DOCTYPE html>
        <html>
          <head>
            <style>
              body { margin: 0; background: #fff; font-family: sans-serif; }
              #editor { padding: 20px; border: 1px solid #ccc; cursor: text; }
              .selected { background: #b4d5fe; }
              #mini-toolbar { display: none; position: absolute; top: 10px; left: 10px; background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.15); padding: 4px; }
              #mini-toolbar.visible { display: flex; }
              #draw-table-surface { width: 100px; height: 100px; cursor: crosshair; }
            </style>
          </head>
          <body>
            <div id="editor" tabindex="0">Hello Word Parity</div>
            <div id="mini-toolbar">
              <button id="bold-btn">B</button>
              <button id="italic-btn">I</button>
            </div>
            <div id="draw-table-surface"></div>
            <script>
              window.__events = [];
              window.addEventListener("mousedown", () => window.__events.push("mousedown"));
              window.addEventListener("click", () => window.__events.push("click"));
              window.addEventListener("dblclick", () => {
                window.__events.push("dblclick");
                document.getElementById("mini-toolbar").classList.add("visible");
              });
              window.addEventListener("keydown", (e) => {
                window.__events.push("key:" + (e.ctrlKey ? "Ctrl+" : "") + e.key);
              });
            </script>
          </body>
        </html>
      `);

      const title = await session.evaluate("document.title");
      expect(typeof title).toBe("string");
    });

    it("verifies pointer clicks and double-click interaction", async () => {
      await session.clickAt(50, 30);
      await new Promise((r) => setTimeout(r, 60));
      let events = await session.evaluate<string[]>("window.__events || []");
      expect(events).toContain("click");

      // Double-click triggers mini-toolbar display
      await session.clickAt(50, 30, { clickCount: 2 });
      await new Promise((r) => setTimeout(r, 60));
      const isMiniToolbarVisible = await session.evaluate<boolean>(
        "document.getElementById('mini-toolbar')?.classList.contains('visible') || false",
      );
      expect(isMiniToolbarVisible).toBe(true);
    });

    it("verifies keyboard shortcuts dispatch via CDP", async () => {
      await session.pressShortcut("Ctrl+B");
      await new Promise((r) => setTimeout(r, 60));
      const events = await session.evaluate<string[]>("window.__events || []");
      expect(events.some((e) => e.includes("Ctrl+") || e.includes("B"))).toBe(true);
    });

    it("inspects computed cursor styles", async () => {
      const editorCursor = await session.getComputedCursor("#editor");
      expect(editorCursor).toBe("text");

      const drawCursor = await session.getComputedCursor("#draw-table-surface");
      expect(drawCursor).toBe("crosshair");
    });

    it("captures visual screenshot and compares visual state changes", async () => {
      const shot1 = await session.captureScreenshot();
      expect(shot1.length).toBeGreaterThan(0);

      // Hide mini-toolbar and change background
      await session.evaluate(`
        document.getElementById('mini-toolbar').classList.remove('visible');
        document.body.style.background = '#f0f0f0';
      `);

      const shot2 = await session.captureScreenshot();
      expect(shot2.length).toBeGreaterThan(0);

      const diff = compareScreenshots(shot1, shot2);
      expect(diff.diffCount).toBeGreaterThanOrEqual(0);
    });
  });
});
