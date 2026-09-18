// @vitest-environment happy-dom
import { beforeEach, describe, expect, it } from "vitest";

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

const { DocenEditor } = await import("./index");

import DocenDialog from "../ui/components/workspace/dialog";
import { DocenNavPane } from "../ui/components/workspace/nav-pane";
import { DocenRuler } from "../ui/components/workspace/ruler";
import { DocenStatusBar } from "../ui/components/workspace/status-bar";
import {
  getUiDirection,
  notifyLocaleChange,
  resolveDir,
  setUiDirection,
} from "../ui/i18n/localize";
import {
  NavigationViewsHostCommands,
  type NavigationViewsHostView,
} from "./commands/host/navigation-views";

describe("W5.4 Full RTL UI Mirroring", () => {
  beforeEach(() => {
    setUiDirection("auto");
    document.documentElement.removeAttribute("dir");
    document.documentElement.removeAttribute("lang");
    document.body.replaceChildren();
  });

  it("resolves UI direction correctly based on language, explicit dir, and setUiDirection", () => {
    // Default LTR
    expect(getUiDirection()).toBe("auto");
    expect(resolveDir()).toBe("ltr");

    // Explicit setUiDirection
    setUiDirection("rtl");
    expect(getUiDirection()).toBe("rtl");
    expect(resolveDir()).toBe("rtl");

    setUiDirection("ltr");
    expect(getUiDirection()).toBe("ltr");
    expect(resolveDir()).toBe("ltr");

    // Reset to auto
    setUiDirection("auto");

    // Language detection: Arabic / Hebrew
    document.documentElement.setAttribute("lang", "ar");
    notifyLocaleChange();
    expect(resolveDir()).toBe("rtl");

    document.documentElement.setAttribute("lang", "he");
    notifyLocaleChange();
    expect(resolveDir()).toBe("rtl");

    document.documentElement.setAttribute("lang", "en");
    notifyLocaleChange();
    expect(resolveDir()).toBe("ltr");

    // Explicit dir on element overrides lang
    const container = document.createElement("div");
    container.setAttribute("dir", "rtl");
    document.body.append(container);
    expect(resolveDir(container)).toBe("rtl");

    container.setAttribute("dir", "ltr");
    expect(resolveDir(container)).toBe("ltr");
  });

  it("detects dir attribute across shadow DOM host boundaries", () => {
    const host = document.createElement("div");
    host.setAttribute("dir", "rtl");
    document.body.append(host);

    const shadow = host.attachShadow({ mode: "open" });
    const innerChild = document.createElement("span");
    shadow.append(innerChild);

    expect(resolveDir(innerChild)).toBe("rtl");
  });

  it("mirrors DocenStatusBar in RTL mode", () => {
    const status = new DocenStatusBar();
    document.body.append(status);

    expect(status.isRtl).toBe(false);

    status.setAttribute("dir", "rtl");
    expect(status.isRtl).toBe(true);

    // Dynamic RTL via setUiDirection
    setUiDirection("rtl");
    status.dirChanged();
    expect(status.isRtl).toBe(true);
    expect(status.getAttribute("dir")).toBe("rtl");
    setUiDirection("auto");
  });

  it("mirrors DocenRuler coordinate math and ticks in RTL mode", () => {
    const ruler = new DocenRuler();
    ruler.pageWidthPx = 800;
    ruler.marginLeftPx = 100;
    ruler.marginRightPx = 100;
    ruler.scale = 1;
    document.body.append(ruler);

    expect(ruler.isRtl).toBe(false);

    // In LTR: zero point is at marginLeftPx (100)
    expect(ruler.marginLeftPx).toBe(100);

    // Switch to RTL
    ruler.setAttribute("dir", "rtl");
    expect(ruler.isRtl).toBe(true);

    ruler.marginLeftPx = 120;
    ruler.marginRightPx = 120;
    expect(ruler.isRtl).toBe(true);
  });

  it("mirrors DocenNavPane tree item padding and search box in RTL mode", () => {
    setUiDirection("auto");
    const pane = new DocenNavPane();
    pane.setAttribute("dir", "ltr");
    document.body.append(pane);

    pane.headings = [
      { id: "h1", title: "Heading 1", level: 1, pos: 1, nodeIndex: 0 },
      { id: "h2", title: "Subheading", level: 2, pos: 10, nodeIndex: 1 },
    ];
    pane.visibleHeadings = pane.headings;
    pane.renderHeadingsTree();

    const tree = pane.treeContainer!;
    const items = tree.querySelectorAll<HTMLElement>(".tree-item");
    expect(items.length).toBe(2);
    // In LTR: paddingLeft is applied
    expect(items[1].style.paddingLeft).toBe("24px"); // (2-1)*16 + 8 = 24px

    // Switch to RTL
    pane.setAttribute("dir", "rtl");
    pane.dirChanged();

    const rtlItems = tree.querySelectorAll<HTMLElement>(".tree-item");
    expect(rtlItems[1].style.paddingRight).toBe("24px");
    expect(rtlItems[1].style.paddingLeft).toBe("8px");
  });

  it("mirrors DocenDialog header and actions in RTL mode", () => {
    const dialog = new DocenDialog();
    dialog.heading = "RTL Dialog";
    document.body.append(dialog);

    expect(dialog.isRtl).toBe(false);

    dialog.setAttribute("dir", "rtl");
    dialog.dirChanged();
    expect(dialog.isRtl).toBe(true);
    expect(dialog.getAttribute("dir")).toBe("rtl");
  });

  it("executes set-ui-direction command via NavigationViewsHostCommands", () => {
    let directed: string | undefined;
    const hostView: NavigationViewsHostView = {
      editor: () => null,
      togglePane: () => {},
      goToPage: () => {},
      openSearch: () => {},
      openFindReplace: () => {},
      zoom: () => 100,
      setZoom: () => {},
      showZoomDialog: () => {},
      zoomPreset: () => {},
      docProtected: () => false,
      syncEditModeMenu: () => {},
      setShowMarks: () => {},
      getShowMarks: () => false,
      showRuler: () => false,
      setShowRuler: () => {},
      showGridlines: () => false,
      setShowGridlines: () => {},
      setView: () => {},
      setUiDirection: (dir) => {
        directed = dir;
        setUiDirection(dir);
      },
    };

    const navCommands = new NavigationViewsHostCommands(hostView);
    expect(navCommands.chrome).toContain("set-ui-direction");

    const ran = navCommands.run("set-ui-direction", "rtl");
    expect(ran).toBe(true);
    expect(directed).toBe("rtl");
    expect(getUiDirection()).toBe("rtl");

    navCommands.run("set-ui-direction", "ltr");
    expect(directed).toBe("ltr");
    expect(getUiDirection()).toBe("ltr");
  });

  it("provides <docen-editor> custom element with touchMode forwarding", () => {
    expect(customElements.get("docen-editor")).toBeDefined();
    const editor = new DocenEditor();
    expect(editor).toBeInstanceOf(DocenEditor);

    expect(editor.touchMode).toBe(false);
    editor.touchMode = true;
    expect(editor.touchMode).toBe(true);

    editor.setAttribute("dir", "rtl");
    expect(editor.isRtl).toBe(true);
  });
});
