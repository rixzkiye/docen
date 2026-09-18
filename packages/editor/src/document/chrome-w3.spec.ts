// @vitest-environment happy-dom
import { Document, Paragraph } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { DocenRibbon } from "../ui/components/ribbon";
import { appendMenuItems } from "../ui/components/ribbon/command-helpers";
import { DocenKeyTips } from "../ui/components/ribbon/key-tips";
import { DocenMiniToolbar } from "../ui/components/workspace/mini-toolbar";
import {
  DocenStatusBar,
  STATUS_BAR_STORAGE_KEY,
  type StatusBarWidgetsConfig,
} from "../ui/components/workspace/status-bar";
import { DocumentCommands, WIRED_DISPATCH } from "./extensions/commands";

const Text = TextNode.create({ name: "text", group: "inline" });

const buildEditor = (content = "Hello Docen Chrome W3"): EditorType => {
  const editor = new Editor({
    element: null,
    extensions: [Document, Paragraph, Text, DocumentCommands, UndoRedo],
    content: {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: content }],
        },
      ],
    },
  });
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
};

describe("Lane W3 Chrome Pack (W3.1 - W3.5)", () => {
  // ── W3.1: Mini Toolbar ──────────────────────────────────────────────────
  describe("W3.1: Floating Mini Toolbar (<docen-mini-toolbar>)", () => {
    let toolbar: DocenMiniToolbar;

    beforeEach(() => {
      toolbar = new DocenMiniToolbar();
      document.body.appendChild(toolbar);
    });

    afterEach(() => {
      toolbar.remove();
    });

    it("initializes and renders all 11 formatting controls", () => {
      const shadow = toolbar.shadowRoot;
      expect(shadow).not.toBeNull();

      // Font Family & Size selects
      const fontSelect = shadow?.querySelector(".font-select") as HTMLSelectElement;
      const sizeSelect = shadow?.querySelector(".size-select") as HTMLSelectElement;
      expect(fontSelect).not.toBeNull();
      expect(sizeSelect).not.toBeNull();
      expect(fontSelect.options.length).toBeGreaterThan(5);
      expect(sizeSelect.options.length).toBeGreaterThan(10);

      // Buttons: grow, shrink, bold, italic, underline, textColor, highlight, bullet, format-painter
      const buttons = shadow?.querySelectorAll("button[data-cmd]");
      expect(buttons?.length).toBe(9);

      const cmds = Array.from(buttons ?? []).map((b) => (b as HTMLElement).dataset.cmd);
      expect(cmds).toContain("grow-font");
      expect(cmds).toContain("shrink-font");
      expect(cmds).toContain("bold");
      expect(cmds).toContain("italic");
      expect(cmds).toContain("underline");
      expect(cmds).toContain("font-color");
      expect(cmds).toContain("highlight");
      expect(cmds).toContain("bullet-list");
      expect(cmds).toContain("format-painter");
    });

    it("positions near selection bounds and sets data-open on showNear", () => {
      const rect = { left: 100, top: 150, right: 250, bottom: 170 };
      toolbar.showNear(rect);

      expect(toolbar.hasAttribute("data-open")).toBe(true);
      expect(toolbar.isOpen).toBe(true);
      expect(toolbar.style.left).toBeTruthy();
      expect(toolbar.style.top).toBeTruthy();
    });

    it("hides and removes data-open on hide()", () => {
      toolbar.showNear({ left: 100, top: 150, right: 200, bottom: 170 });
      expect(toolbar.isOpen).toBe(true);

      toolbar.hide();
      expect(toolbar.isOpen).toBe(false);
      expect(toolbar.style.opacity).toBe("0");
    });

    it("updates formatting indicators accurately", () => {
      toolbar.updateFormatting({
        bold: true,
        italic: false,
        underline: true,
        fontName: "Arial",
        fontSize: "16",
        fontColor: "#ff0000",
        highlightColor: "#00ff00",
      });

      const shadow = toolbar.shadowRoot;
      const boldBtn = shadow?.querySelector('button[data-cmd="bold"]');
      const italicBtn = shadow?.querySelector('button[data-cmd="italic"]');
      const underlineBtn = shadow?.querySelector('button[data-cmd="underline"]');

      expect(boldBtn?.getAttribute("aria-pressed")).toBe("true");
      expect(italicBtn?.getAttribute("aria-pressed")).toBe("false");
      expect(underlineBtn?.getAttribute("aria-pressed")).toBe("true");

      const fontSelect = shadow?.querySelector(".font-select") as HTMLSelectElement;
      const sizeSelect = shadow?.querySelector(".size-select") as HTMLSelectElement;
      expect(fontSelect.value).toBe("Arial");
      expect(sizeSelect.value).toBe("16");
    });

    it("dispatches command events when controls are clicked or selected", () => {
      const dispatched: Array<{ event: string; value?: string }> = [];
      toolbar.addEventListener("command", ((e: CustomEvent) => {
        dispatched.push({ event: e.detail.event, value: e.detail.value });
      }) as EventListener);

      const shadow = toolbar.shadowRoot;
      const boldBtn = shadow?.querySelector('button[data-cmd="bold"]') as HTMLButtonElement;
      boldBtn?.click();

      expect(dispatched.length).toBe(1);
      expect(dispatched[0].event).toBe("bold");

      const painterBtn = shadow?.querySelector(
        'button[data-cmd="format-painter"]',
      ) as HTMLButtonElement;
      painterBtn?.click();
      expect(dispatched[1].event).toBe("copy-format");

      const fontSelect = shadow?.querySelector(".font-select") as HTMLSelectElement;
      fontSelect.value = "Georgia";
      fontSelect.dispatchEvent(new Event("change"));
      expect(dispatched[2].event).toBe("font-name");
      expect(dispatched[2].value).toBe("Georgia");
    });

    it("fades out based on distance from pointer", () => {
      toolbar.showNear({ left: 100, top: 100, right: 200, bottom: 120 });
      toolbar.getBoundingClientRect = () => ({
        left: 100,
        top: 60,
        right: 440,
        bottom: 94,
        width: 340,
        height: 34,
        x: 100,
        y: 60,
        toJSON: () => {},
      });

      // Pointer close to toolbar (< 40px)
      document.dispatchEvent(new PointerEvent("pointermove", { clientX: 110, clientY: 70 }));
      expect(toolbar.style.opacity).toBe("1");

      // Pointer moderately away (between 40px and 160px)
      document.dispatchEvent(new PointerEvent("pointermove", { clientX: 110, clientY: 180 }));
      expect(Number(toolbar.style.opacity)).toBeLessThan(1);
      expect(Number(toolbar.style.opacity)).toBeGreaterThan(0);

      // Pointer very far away (> 160px) triggers hide
      document.dispatchEvent(new PointerEvent("pointermove", { clientX: 500, clientY: 500 }));
      expect(toolbar.isOpen).toBe(false);
    });
  });

  // ── W3.2: Context Menus ─────────────────────────────────────────────────
  describe("W3.2: Context Menus per context", () => {
    it("renders text context menu with full Word suite including nested submenus", () => {
      const container = document.createElement("div");
      const textMenuItems = [
        { text: "Cut", event: "cut" },
        { text: "Copy", event: "copy" },
        { text: "Paste", event: "paste" },
        { text: "Keep Text Only", event: "paste", value: "keep-text-only" },
        { text: "-" },
        { text: "Font...", event: "font-dialog" },
        { text: "Paragraph...", event: "paragraph-dialog" },
        {
          text: "Styles",
          items: [
            { text: "Styles Pane...", event: "styles-pane" },
            { text: "-" },
            { text: "Normal", event: "style", value: "Normal" },
            { text: "Heading 1", event: "style", value: "Heading 1" },
          ],
        },
        {
          text: "Bullets & Numbering",
          items: [
            { text: "Bullet List", event: "bullet-list" },
            { text: "Numbered List", event: "ordered-list" },
            { text: "Multilevel List", event: "multilevel-list" },
          ],
        },
        { text: "-" },
        { text: "Translate", event: "translate" },
        { text: "Link…", event: "link" },
      ];

      const clicks: any[] = [];
      appendMenuItems(container, textMenuItems, (item) => clicks.push(item));

      // Font & Paragraph items present
      const fontItem = Array.from(container.children).find((c) =>
        c.textContent?.includes("Font..."),
      );
      const paraItem = Array.from(container.children).find((c) =>
        c.textContent?.includes("Paragraph..."),
      );
      expect(fontItem).toBeDefined();
      expect(paraItem).toBeDefined();

      // Styles submenu has nested menu list
      const stylesItem = Array.from(container.children).find((c) =>
        c.textContent?.includes("Styles"),
      ) as HTMLElement;
      expect(stylesItem.hasAttribute("data-has-submenu")).toBe(true);
      const stylesSub = stylesItem.querySelector("fluent-menu-list");
      expect(stylesSub).not.toBeNull();
      expect(stylesSub?.children.length).toBe(4);

      // Bullets & Numbering submenu
      const bulletsItem = Array.from(container.children).find((c) =>
        c.textContent?.includes("Bullets"),
      ) as HTMLElement;
      expect(bulletsItem.hasAttribute("data-has-submenu")).toBe(true);
      const bulletsSub = bulletsItem.querySelector("fluent-menu-list");
      expect(bulletsSub?.children.length).toBe(3);

      // Translate item
      const translateItem = Array.from(container.children).find((c) =>
        c.textContent?.includes("Translate"),
      ) as HTMLElement;
      expect(translateItem).toBeDefined();
      translateItem.click();
      expect(clicks.length).toBe(1);
      expect(clicks[0].event).toBe("translate");
    });

    it("renders spelling error context menu with suggestions at top and single-click pick", () => {
      const container = document.createElement("div");
      const spellingMenuItems = [
        { text: "spelling", event: "spell-pick", value: "spelling" },
        { text: "spieling", event: "spell-pick", value: "spieling" },
        { text: "-" },
        { text: "Ignore Once", event: "spell-ignore-once" },
        { text: "Ignore All", event: "spell-ignore-all" },
        { text: "Add to Dictionary", event: "spell-add" },
        { text: "Spelling...", event: "spell-check" },
      ];

      const picked: any[] = [];
      appendMenuItems(container, spellingMenuItems, (item) => picked.push(item));

      const firstItem = container.children[0] as HTMLElement;
      expect(firstItem.textContent).toBe("spelling");

      // Single click on spelling candidate triggers replacement
      firstItem.click();
      expect(picked.length).toBe(1);
      expect(picked[0].event).toBe("spell-pick");
      expect(picked[0].value).toBe("spelling");

      // Spelling... pane shortcut present
      const paneShortcut = Array.from(container.children).find((c) =>
        c.textContent?.includes("Spelling..."),
      ) as HTMLElement;
      expect(paneShortcut).toBeDefined();
      paneShortcut.click();
      expect(picked[1].event).toBe("spell-check");
    });

    it("renders shape/image context menu with Wrap Text submenu and Drawing properties", () => {
      const container = document.createElement("div");
      const drawingMenuItems = [
        { text: "Cut", event: "cut" },
        { text: "Copy", event: "copy" },
        { text: "Paste", event: "paste" },
        { text: "-" },
        {
          text: "Wrap Text",
          items: [
            { text: "In Line with Text", event: "wrap", value: "inline" },
            { text: "Square", event: "wrap", value: "square" },
            { text: "Tight", event: "wrap", value: "tight" },
            { text: "Behind Text", event: "wrap", value: "behind" },
            { text: "In Front of Text", event: "wrap", value: "front" },
            { text: "Top and Bottom", event: "wrap", value: "top-bottom" },
          ],
        },
        { text: "-" },
        { text: "Bring to Front", event: "bring-to-front" },
        { text: "Send to Back", event: "send-to-back" },
        { text: "Size and Position...", event: "drawing-properties" },
        { text: "Crop", event: "drawing-crop" },
      ];

      const clicks: any[] = [];
      appendMenuItems(container, drawingMenuItems, (item) => clicks.push(item));

      const wrapItem = Array.from(container.children).find((c) =>
        c.textContent?.includes("Wrap Text"),
      ) as HTMLElement;
      expect(wrapItem).toBeDefined();
      expect(wrapItem.hasAttribute("data-has-submenu")).toBe(true);

      const wrapSub = wrapItem.querySelector("fluent-menu-list");
      expect(wrapSub?.children.length).toBe(6);

      const squareWrap = wrapSub?.children[1] as HTMLElement;
      squareWrap.click();
      expect(clicks.length).toBe(1);
      expect(clicks[0].event).toBe("wrap");
      expect(clicks[0].value).toBe("square");
    });

    it("registers all required context commands in WIRED_DISPATCH", () => {
      expect(WIRED_DISPATCH.has("font-dialog")).toBe(true);
      expect(WIRED_DISPATCH.has("paragraph-dialog-apply")).toBe(true);
      expect(WIRED_DISPATCH.has("bullet-list")).toBe(true);
      expect(WIRED_DISPATCH.has("ordered-list")).toBe(true);
      expect(WIRED_DISPATCH.has("multilevel-list")).toBe(true);
      expect(WIRED_DISPATCH.has("table-properties-apply")).toBe(true);
      expect(WIRED_DISPATCH.has("wrap")).toBe(true);
      expect(WIRED_DISPATCH.has("translate")).toBe(true);
      expect(WIRED_DISPATCH.has("toggle-ribbon-minimized")).toBe(true);
    });
  });

  // ── W3.3: Ribbon Key Tips ───────────────────────────────────────────────
  describe("W3.3: Ribbon Key Tips overlay (<docen-key-tips>)", () => {
    let keyTips: DocenKeyTips;

    beforeEach(() => {
      keyTips = new DocenKeyTips();
      document.body.appendChild(keyTips);
    });

    afterEach(() => {
      keyTips.remove();
    });

    it("toggles active state and renders container", () => {
      expect(keyTips.isActive).toBe(false);
      keyTips.activate();
      expect(keyTips.isActive).toBe(true);
      expect(keyTips.hasAttribute("data-active")).toBe(true);

      keyTips.dismiss();
      expect(keyTips.isActive).toBe(false);
      expect(keyTips.hasAttribute("data-active")).toBe(false);
    });

    it("renders badges for ribbon tabs on activate", () => {
      // Create mock ribbon tabs
      const tabHome = document.createElement("docen-ribbon-tab");
      tabHome.id = "home";
      tabHome.textContent = "Home";
      tabHome.getBoundingClientRect = () => ({
        left: 20,
        top: 10,
        width: 60,
        height: 26,
        right: 80,
        bottom: 36,
        x: 20,
        y: 10,
        toJSON: () => {},
      });

      const tabInsert = document.createElement("docen-ribbon-tab");
      tabInsert.id = "insert";
      tabInsert.textContent = "Insert";
      tabInsert.getBoundingClientRect = () => ({
        left: 80,
        top: 10,
        width: 60,
        height: 26,
        right: 140,
        bottom: 36,
        x: 80,
        y: 10,
        toJSON: () => {},
      });

      document.body.append(tabHome, tabInsert);

      keyTips.activate();

      const badges = keyTips.badges;
      expect(badges.length).toBe(2);
      expect(badges[0].key).toBe("H");
      expect(badges[1].key).toBe("I");

      tabHome.remove();
      tabInsert.remove();
      keyTips.dismiss();
    });

    it("handles Alt key press to toggle overlay", () => {
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "Alt" }));
      expect(keyTips.isActive).toBe(true);

      window.dispatchEvent(new KeyboardEvent("keydown", { key: "Alt" }));
      expect(keyTips.isActive).toBe(false);
    });

    it("handles Escape key to dismiss overlay", () => {
      keyTips.activate();
      expect(keyTips.isActive).toBe(true);

      window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
      expect(keyTips.isActive).toBe(false);
    });
  });

  // ── W3.4: Ribbon Collapse & Status Bar Customizer ────────────────────────
  describe("W3.4: Ribbon Collapse (Ctrl+F1) & Status Bar Customizer", () => {
    describe("Ribbon Collapse", () => {
      let ribbon: DocenRibbon;

      beforeEach(() => {
        ribbon = new DocenRibbon();
        document.body.appendChild(ribbon);
      });

      afterEach(() => {
        ribbon.remove();
      });

      it("starts with always mode and toggles to tabs-only via toggleMinimized()", () => {
        expect(ribbon.currentMode).toBe("always");

        ribbon.toggleMinimized();
        expect(ribbon.currentMode).toBe("tabs-only");
        expect(ribbon.getAttribute("data-ribbon-mode")).toBe("tabs-only");

        ribbon.toggleMinimized();
        expect(ribbon.currentMode).toBe("always");
        expect(ribbon.hasAttribute("data-ribbon-mode")).toBe(false);
      });
    });

    describe("Status Bar Customization Menu", () => {
      let statusBar: DocenStatusBar;

      beforeEach(() => {
        localStorage.clear();
        statusBar = new DocenStatusBar();
        document.body.appendChild(statusBar);
      });

      afterEach(() => {
        statusBar.remove();
        document.body.querySelectorAll("docen-status-bar").forEach((el) => el.remove());
        localStorage.clear();
      });

      it("loads default widgets configuration and exposes 6 customization items", () => {
        const config = statusBar.widgetsConfig;
        expect(config.pageNumber).toBe(true);
        expect(config.wordCount).toBe(true);
        expect(config.language).toBe(true);
        expect(config.extendSelection).toBe(true);
        expect(config.capsLock).toBe(true);
        expect(config.zoom).toBe(true);
      });

      it("persists widget toggle changes to localStorage", () => {
        statusBar.setWidgetVisibility("capsLock", false);
        expect(statusBar.widgetsConfig.capsLock).toBe(false);

        const stored = JSON.parse(
          localStorage.getItem(STATUS_BAR_STORAGE_KEY) || "{}",
        ) as StatusBarWidgetsConfig;
        expect(stored.capsLock).toBe(false);
        expect(stored.pageNumber).toBe(true);

        statusBar.setWidgetVisibility("pageNumber", false);
        const stored2 = JSON.parse(
          localStorage.getItem(STATUS_BAR_STORAGE_KEY) || "{}",
        ) as StatusBarWidgetsConfig;
        expect(stored2.pageNumber).toBe(false);
      });

      it("shows status-bar context menu on right-click with 6 customizable widgets", () => {
        statusBar.dispatchEvent(
          new MouseEvent("contextmenu", { clientX: 200, clientY: 500, bubbles: true }),
        );

        const shadow = statusBar.shadowRoot;
        const popover = shadow?.querySelector(".status-context-menu") as HTMLElement;
        expect(popover).not.toBeNull();
        expect(popover.style.display).toBe("block");

        const items = popover.querySelectorAll(".status-menu-item");
        expect(items.length).toBe(6);

        // Click a widget row to toggle visibility
        const zoomItem = Array.from(items).find(
          (el) => (el as HTMLElement).dataset.widget === "zoom",
        ) as HTMLElement;
        expect(zoomItem).toBeDefined();

        zoomItem.click();
        expect(statusBar.widgetsConfig.zoom).toBe(false);
      });

      it("detects Caps Lock and updates caps lock pill", () => {
        statusBar.setWidgetVisibility("capsLock", true);
        const shadow = statusBar.shadowRoot;
        const capsPill = shadow?.querySelector(".caps-lock") as HTMLElement;
        expect(capsPill).not.toBeNull();
        expect(capsPill.style.display).toBe("none");

        // Simulate modifier key press with CapsLock on
        const evOn = new KeyboardEvent("keydown");
        Object.defineProperty(evOn, "getModifierState", {
          value: (key: string) => key === "CapsLock",
        });
        window.dispatchEvent(evOn);
        expect(capsPill.style.display).toBe("inline-flex");

        // Simulate modifier key press with CapsLock off
        const evOff = new KeyboardEvent("keydown");
        Object.defineProperty(evOff, "getModifierState", {
          value: () => false,
        });
        window.dispatchEvent(evOff);
        expect(capsPill.style.display).toBe("none");
      });
    });
  });

  // ── W3.5: Live Preview Hover ────────────────────────────────────────────
  describe("W3.5: Live Preview Hover in Galleries", () => {
    it("dispatches item-preview on pointerenter and item-preview-end on pointerleave", () => {
      const container = document.createElement("div");
      const items = [
        { text: "Heading 1", event: "style", value: "Heading 1" },
        { text: "Heading 2", event: "style", value: "Heading 2" },
      ];

      const events: string[] = [];
      container.addEventListener("item-preview", ((e: CustomEvent) => {
        events.push(`preview:${e.detail.value}`);
      }) as EventListener);
      container.addEventListener("item-preview-end", (() => {
        events.push("preview-end");
      }) as EventListener);

      appendMenuItems(container, items, () => {});

      const item0 = container.children[0] as HTMLElement;
      item0.dispatchEvent(new PointerEvent("pointerenter", { bubbles: true }));
      expect(events).toContain("preview:Heading 1");

      item0.dispatchEvent(new PointerEvent("pointerleave", { bubbles: true }));
      expect(events).toContain("preview-end");
    });

    it("applies temporary preview transaction with addToHistory: false without modifying undo stack", () => {
      const editor = buildEditor("Test content for live preview");
      const initialHistoryLength = (editor.state as any).history$?.done?.eventCount ?? 0;

      // Apply a preview transaction with addToHistory: false
      const tr = editor.state.tr.insertText(" [Preview]", 12);
      tr.setMeta("addToHistory", false);
      editor.view.dispatch(tr);

      expect(editor.state.doc.textContent).toContain("[Preview]");
      // History should not record preview transaction
      const previewHistoryLength = (editor.state as any).history$?.done?.eventCount ?? 0;
      expect(previewHistoryLength).toBe(initialHistoryLength);

      // Revert preview transaction with addToHistory: false
      const revertTr = editor.state.tr.delete(12, 12 + " [Preview]".length);
      revertTr.setMeta("addToHistory", false);
      editor.view.dispatch(revertTr);

      expect(editor.state.doc.textContent).toBe("Test content for live preview");
      const revertedHistoryLength = (editor.state as any).history$?.done?.eventCount ?? 0;
      expect(revertedHistoryLength).toBe(initialHistoryLength);
    });
  });
});
