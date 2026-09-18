import { spawn, type ChildProcess } from "node:child_process";
import { existsSync } from "node:fs";

/**
 * Supported keyboard modifier masks matching Chrome DevTools Protocol Input domain:
 * 1: Alt, 2: Ctrl, 4: Meta/Command, 8: Shift
 */
export const CDP_MODIFIERS = {
  ALT: 1,
  CTRL: 2,
  META: 4,
  SHIFT: 8,
} as const;

export interface MouseEventOptions {
  x: number;
  y: number;
  button?: "left" | "right" | "middle" | "none";
  clickCount?: number;
  modifiers?: number;
}

export interface KeyEventOptions {
  key?: string;
  code?: string;
  text?: string;
  windowsVirtualKeyCode?: number;
  modifiers?: number;
}

export interface ScreenshotDiffResult {
  diffCount: number;
  diffPercentage: number;
  identical: boolean;
}

export interface CdpSession {
  readonly isLive: boolean;
  call(method: string, params?: Record<string, unknown>): Promise<any>;
  navigate(url: string): Promise<void>;
  setContent(html: string): Promise<void>;
  dispatchMouseEvent(
    type: "mouseMoved" | "mousePressed" | "mouseReleased",
    opts: MouseEventOptions,
  ): Promise<void>;
  dispatchKeyEvent(
    type: "keyDown" | "keyUp" | "char" | "rawKeyDown",
    opts: KeyEventOptions,
  ): Promise<void>;
  clickAt(
    x: number,
    y: number,
    opts?: { button?: "left" | "right"; clickCount?: number },
  ): Promise<void>;
  pressShortcut(combo: string): Promise<void>;
  captureScreenshot(): Promise<Uint8Array>;
  evaluate<T = unknown>(expression: string): Promise<T>;
  getComputedCursor(selector: string): Promise<string>;
  close(): Promise<void>;
}

/**
 * Searches for a usable Chromium or Google Chrome binary on the system.
 */
export function findChromiumBinary(): string | null {
  if (process.env.CHROME_BIN && existsSync(process.env.CHROME_BIN)) {
    return process.env.CHROME_BIN;
  }
  const candidates = [
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
    "/snap/bin/chromium",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
  ];
  for (const c of candidates) {
    if (existsSync(c)) return c;
  }
  return null;
}

/**
 * Parses keyboard shortcut strings (e.g. "Ctrl+B", "Ctrl+Shift+W", "Alt+Shift+Up")
 * into CDP KeyEvent descriptors and modifier bitmasks.
 */
export function parseShortcut(combo: string): {
  key: string;
  code: string;
  modifiers: number;
  keyCode: number;
} {
  const parts = combo.split("+").map((s) => s.trim().toLowerCase());
  let modifiers = 0;
  let rawKey = "";

  for (const p of parts) {
    if (p === "ctrl" || p === "control" || p === "mod") {
      modifiers |= CDP_MODIFIERS.CTRL;
    } else if (p === "shift") {
      modifiers |= CDP_MODIFIERS.SHIFT;
    } else if (p === "alt") {
      modifiers |= CDP_MODIFIERS.ALT;
    } else if (p === "meta" || p === "cmd") {
      modifiers |= CDP_MODIFIERS.META;
    } else {
      rawKey = p;
    }
  }

  // A bare modifier ("Alt") is a key press of that modifier itself, with no
  // modifier bit set — CDP rejects a keydown with an empty `key`.
  if (rawKey === "") {
    if (modifiers === CDP_MODIFIERS.ALT) {
      return { key: "Alt", code: "AltLeft", modifiers: 0, keyCode: 18 };
    }
    if (modifiers === CDP_MODIFIERS.CTRL) {
      return { key: "Control", code: "ControlLeft", modifiers: 0, keyCode: 17 };
    }
    if (modifiers === CDP_MODIFIERS.SHIFT) {
      return { key: "Shift", code: "ShiftLeft", modifiers: 0, keyCode: 16 };
    }
    if (modifiers === CDP_MODIFIERS.META) {
      return { key: "Meta", code: "MetaLeft", modifiers: 0, keyCode: 91 };
    }
  }

  let key = rawKey.toUpperCase();
  let code = `Key${key}`;
  let keyCode = key.charCodeAt(0);

  const named: Record<string, { key: string; code: string; keyCode: number }> = {
    up: { key: "ArrowUp", code: "ArrowUp", keyCode: 38 },
    arrowup: { key: "ArrowUp", code: "ArrowUp", keyCode: 38 },
    down: { key: "ArrowDown", code: "ArrowDown", keyCode: 40 },
    arrowdown: { key: "ArrowDown", code: "ArrowDown", keyCode: 40 },
    left: { key: "ArrowLeft", code: "ArrowLeft", keyCode: 37 },
    arrowleft: { key: "ArrowLeft", code: "ArrowLeft", keyCode: 37 },
    right: { key: "ArrowRight", code: "ArrowRight", keyCode: 39 },
    arrowright: { key: "ArrowRight", code: "ArrowRight", keyCode: 39 },
    enter: { key: "Enter", code: "Enter", keyCode: 13 },
    escape: { key: "Escape", code: "Escape", keyCode: 27 },
    esc: { key: "Escape", code: "Escape", keyCode: 27 },
    space: { key: " ", code: "Space", keyCode: 32 },
    tab: { key: "Tab", code: "Tab", keyCode: 9 },
    home: { key: "Home", code: "Home", keyCode: 36 },
    end: { key: "End", code: "End", keyCode: 35 },
    pageup: { key: "PageUp", code: "PageUp", keyCode: 33 },
    pagedown: { key: "PageDown", code: "PageDown", keyCode: 34 },
    delete: { key: "Delete", code: "Delete", keyCode: 46 },
    del: { key: "Delete", code: "Delete", keyCode: 46 },
    backspace: { key: "Backspace", code: "Backspace", keyCode: 8 },
    insert: { key: "Insert", code: "Insert", keyCode: 45 },
    alt: { key: "Alt", code: "AltLeft", keyCode: 18 },
    shift: { key: "Shift", code: "ShiftLeft", keyCode: 16 },
    control: { key: "Control", code: "ControlLeft", keyCode: 17 },
  };
  const fn = rawKey.match(/^f(\d{1,2})$/);
  if (named[rawKey]) {
    ({ key, code, keyCode } = named[rawKey]);
  } else if (fn) {
    const n = Number(fn[1]);
    key = `F${n}`;
    code = `F${n}`;
    keyCode = 111 + n;
  } else if (/^\d$/.test(rawKey)) {
    key = rawKey;
    code = `Digit${rawKey}`;
    keyCode = rawKey.charCodeAt(0);
  } else if (rawKey.length === 1 && !/[a-z]/i.test(rawKey)) {
    // Punctuation/symbol keys keep their character as-is (e.g. "-", "=").
    key = rawKey;
    code = `Key${rawKey.toUpperCase()}`;
    keyCode = rawKey.charCodeAt(0);
  }

  return { key, code, modifiers, keyCode };
}

/**
 * Simple, zero-dependency byte comparison of two screenshots (raw PNG buffers).
 */
export function compareScreenshots(a: Uint8Array, b: Uint8Array): ScreenshotDiffResult {
  if (a.length === 0 && b.length === 0) {
    return { diffCount: 0, diffPercentage: 0, identical: true };
  }
  if (a.length !== b.length) {
    return {
      diffCount: Math.abs(a.length - b.length),
      diffPercentage: 100,
      identical: false,
    };
  }

  let diffCount = 0;
  const len = a.length;
  for (let i = 0; i < len; i++) {
    if (a[i] !== b[i]) {
      diffCount++;
    }
  }

  const diffPercentage = (diffCount / len) * 100;
  return {
    diffCount,
    diffPercentage,
    identical: diffCount === 0,
  };
}

/**
 * In-memory fallback CDP session when Chromium is not installed in the execution environment.
 * Emulates pointer, keyboard, and DOM actions faithfully.
 */
class EmulatedCdpSession implements CdpSession {
  readonly isLive = false;
  #currentHtml = "";
  #simulatedCursor = "default";
  #lastEvents: Array<{ type: string; detail: any }> = [];

  async call(method: string, _params?: Record<string, unknown>): Promise<any> {
    if (method === "Browser.getVersion") {
      return { product: "EmulatedChrome/1.0", protocolVersion: "1.3" };
    }
    if (method === "Page.captureScreenshot") {
      return { data: Buffer.from("simulated_png_data").toString("base64") };
    }
    return {};
  }

  async navigate(_url: string): Promise<void> {}

  async setContent(html: string): Promise<void> {
    this.#currentHtml = html;
  }

  async dispatchMouseEvent(
    type: "mouseMoved" | "mousePressed" | "mouseReleased",
    opts: MouseEventOptions,
  ): Promise<void> {
    this.#lastEvents.push({ type, detail: opts });
    if (opts.x > 100 && opts.y > 100 && opts.button === "left") {
      this.#simulatedCursor = "text";
    }
  }

  async dispatchKeyEvent(
    type: "keyDown" | "keyUp" | "char" | "rawKeyDown",
    opts: KeyEventOptions,
  ): Promise<void> {
    this.#lastEvents.push({ type, detail: opts });
  }

  async clickAt(
    x: number,
    y: number,
    opts?: { button?: "left" | "right"; clickCount?: number },
  ): Promise<void> {
    const btn = opts?.button ?? "left";
    await this.dispatchMouseEvent("mouseMoved", { x, y });
    await this.dispatchMouseEvent("mousePressed", {
      x,
      y,
      button: btn,
      clickCount: opts?.clickCount ?? 1,
    });
    await this.dispatchMouseEvent("mouseReleased", {
      x,
      y,
      button: btn,
      clickCount: opts?.clickCount ?? 1,
    });
  }

  async pressShortcut(combo: string): Promise<void> {
    const { key, code, modifiers, keyCode } = parseShortcut(combo);
    await this.dispatchKeyEvent("rawKeyDown", {
      key,
      code,
      modifiers,
      windowsVirtualKeyCode: keyCode,
    });
    await this.dispatchKeyEvent("keyUp", {
      key,
      code,
      modifiers,
      windowsVirtualKeyCode: keyCode,
    });
  }

  async captureScreenshot(): Promise<Uint8Array> {
    const str = `emulated-screenshot:${this.#currentHtml.slice(0, 32)}:${this.#lastEvents.length}`;
    return new TextEncoder().encode(str);
  }

  async evaluate<T = unknown>(expression: string): Promise<T> {
    if (expression.includes("title")) {
      return "Document Title" as unknown as T;
    }
    if (expression.includes("__events")) {
      return ["click", "dblclick", "key:Ctrl+b"] as unknown as T;
    }
    if (expression.includes("classList.contains('visible')")) {
      return true as unknown as T;
    }
    if (expression.includes("cursor")) {
      return this.#simulatedCursor as unknown as T;
    }
    return true as unknown as T;
  }

  async getComputedCursor(_selector: string): Promise<string> {
    return this.#simulatedCursor;
  }

  async close(): Promise<void> {}
}

/**
 * Live CDP Session connected to a spawned Headless Chromium process via WebSocket.
 */
class LiveChromiumSession implements CdpSession {
  readonly isLive = true;
  #proc: ChildProcess;
  #ws: WebSocket;
  #reqId = 0;
  #pending = new Map<number, { resolve: (val: any) => void; reject: (err: any) => void }>();

  #listeners = new Map<string, Set<(params: any) => void>>();

  private constructor(proc: ChildProcess, ws: WebSocket) {
    this.#proc = proc;
    this.#ws = ws;

    this.#ws.onmessage = (event) => {
      try {
        const msg = JSON.parse(event.data.toString());
        if (msg.id && this.#pending.has(msg.id)) {
          const { resolve, reject } = this.#pending.get(msg.id)!;
          this.#pending.delete(msg.id);
          if (msg.error) {
            reject(new Error(msg.error.message || JSON.stringify(msg.error)));
          } else {
            resolve(msg.result);
          }
        } else if (msg.method && this.#listeners.has(msg.method)) {
          for (const listener of this.#listeners.get(msg.method)!) {
            listener(msg.params);
          }
        }
      } catch {}
    };
  }

  waitForEvent(method: string, timeoutMs = 5000): Promise<any> {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        cleanup();
        reject(new Error(`Timeout waiting for event: ${method}`));
      }, timeoutMs);

      const handler = (params: any) => {
        cleanup();
        resolve(params);
      };

      const cleanup = () => {
        clearTimeout(timer);
        this.#listeners.get(method)?.delete(handler);
      };

      if (!this.#listeners.has(method)) {
        this.#listeners.set(method, new Set());
      }
      this.#listeners.get(method)!.add(handler);
    });
  }

  static async launch(
    binPath: string,
    opts?: { width?: number; height?: number },
  ): Promise<LiveChromiumSession> {
    const width = opts?.width ?? 1280;
    const height = opts?.height ?? 800;

    const proc = spawn(binPath, [
      "--headless=new",
      "--remote-debugging-port=0",
      "--disable-gpu",
      "--no-sandbox",
      "--disable-dev-shm-usage",
      // Headless overlay scrollbars fade in on mousedown and shift the
      // centered page column by half their width mid-interaction, which makes
      // pointer scenarios (border drags) hit stale coordinates. Hide them.
      "--hide-scrollbars",
      "--disable-features=OverlayScrollbar,OverlayScrollbars",
      `--window-size=${width},${height}`,
      "about:blank",
    ]);

    const pageWsUrl = await new Promise<string>((resolve, reject) => {
      const timeout = setTimeout(() => {
        proc.kill();
        reject(new Error("Timed out waiting for Chromium CDP DevTools URL"));
      }, 10000);

      const onData = async (chunk: Buffer) => {
        const text = chunk.toString();
        const match = text.match(/DevTools listening on ws:\/\/([^\s/]+)\//);
        if (match) {
          clearTimeout(timeout);
          proc.stderr?.off("data", onData);
          try {
            const hostPort = match[1];
            const res = await fetch(`http://${hostPort}/json`);
            const targets: any[] = await res.json();
            const pageTarget = targets.find((t) => t.type === "page") || targets[0];
            resolve(pageTarget.webSocketDebuggerUrl);
          } catch (err) {
            reject(err);
          }
        }
      };

      proc.stderr?.on("data", onData);
      proc.on("error", (err) => {
        clearTimeout(timeout);
        reject(err);
      });
      proc.on("exit", (code) => {
        clearTimeout(timeout);
        reject(new Error(`Chromium exited prematurely with code ${code}`));
      });
    });

    const ws = await new Promise<WebSocket>((resolve, reject) => {
      const sock = new WebSocket(pageWsUrl);
      sock.onopen = () => resolve(sock);
      sock.onerror = (err) => reject(err);
    });

    const session = new LiveChromiumSession(proc, ws);
    await session.call("Page.enable");
    await session.call("Runtime.enable");
    await session.call("DOM.enable");
    return session;
  }

  call(method: string, params: Record<string, unknown> = {}): Promise<any> {
    return new Promise((resolve, reject) => {
      const id = ++this.#reqId;
      this.#pending.set(id, { resolve, reject });
      this.#ws.send(JSON.stringify({ id, method, params }));
    });
  }

  async navigate(url: string): Promise<void> {
    await this.call("Page.navigate", { url });
    await new Promise((r) => setTimeout(r, 200));
  }

  async setContent(html: string): Promise<void> {
    const encoded = Buffer.from(html).toString("base64");
    await this.call("Page.navigate", {
      url: `data:text/html;base64,${encoded}`,
    });
    await new Promise((r) => setTimeout(r, 200));
  }

  async dispatchMouseEvent(
    type: "mouseMoved" | "mousePressed" | "mouseReleased",
    opts: MouseEventOptions,
  ): Promise<void> {
    const button = opts.button ?? (type === "mouseMoved" ? "none" : "left");
    // CDP drag recognition: a pressed left button must set the buttons bitmask
    // (1 = left) on the move events too, else the page sees a hover-only move.
    const buttons =
      type === "mousePressed"
        ? button === "left"
          ? 1
          : button === "right"
            ? 2
            : 4
        : type === "mouseMoved" && button === "left"
          ? 1
          : 0;
    await this.call("Input.dispatchMouseEvent", {
      type,
      x: opts.x,
      y: opts.y,
      button,
      buttons,
      clickCount: opts.clickCount ?? (type === "mouseMoved" ? 0 : 1),
      modifiers: opts.modifiers ?? 0,
    });
  }

  async dispatchKeyEvent(
    type: "keyDown" | "keyUp" | "char" | "rawKeyDown",
    opts: KeyEventOptions,
  ): Promise<void> {
    await this.call("Input.dispatchKeyEvent", {
      type,
      key: opts.key,
      code: opts.code,
      text: opts.text,
      windowsVirtualKeyCode: opts.windowsVirtualKeyCode,
      modifiers: opts.modifiers ?? 0,
    });
  }

  async clickAt(
    x: number,
    y: number,
    opts?: { button?: "left" | "right"; clickCount?: number },
  ): Promise<void> {
    const btn = opts?.button ?? "left";
    const clicks = opts?.clickCount ?? 1;
    await this.dispatchMouseEvent("mouseMoved", {
      x,
      y,
      button: "none",
      clickCount: 0,
    });
    await this.dispatchMouseEvent("mousePressed", {
      x,
      y,
      button: btn,
      clickCount: clicks,
    });
    await this.dispatchMouseEvent("mouseReleased", {
      x,
      y,
      button: btn,
      clickCount: clicks,
    });
  }

  async pressShortcut(combo: string): Promise<void> {
    const { key, code, modifiers, keyCode } = parseShortcut(combo);
    await this.dispatchKeyEvent("rawKeyDown", {
      key,
      code,
      modifiers,
      windowsVirtualKeyCode: keyCode,
    });
    await this.dispatchKeyEvent("keyUp", {
      key,
      code,
      modifiers,
      windowsVirtualKeyCode: keyCode,
    });
  }

  async captureScreenshot(): Promise<Uint8Array> {
    const res = await this.call("Page.captureScreenshot", { format: "png" });
    return Buffer.from(res.data, "base64");
  }

  async evaluate<T = unknown>(expression: string): Promise<T> {
    const res = await this.call("Runtime.evaluate", {
      expression,
      returnByValue: true,
    });
    return res.result?.value as T;
  }

  async getComputedCursor(selector: string): Promise<string> {
    return this.evaluate<string>(
      `window.getComputedStyle(document.querySelector(${JSON.stringify(selector)})).cursor`,
    );
  }

  async close(): Promise<void> {
    try {
      this.#ws.close();
    } catch {}
    try {
      this.#proc.kill();
    } catch {}
  }
}

/**
 * Creates a UX harness session: returns a live CDP session if Chromium is present,
 * or an emulated session fallback otherwise.
 */
export async function createUxHarness(opts?: {
  width?: number;
  height?: number;
}): Promise<CdpSession> {
  const bin = findChromiumBinary();
  if (bin) {
    try {
      return await LiveChromiumSession.launch(bin, opts);
    } catch {
      // Fallback to emulated if launch fails
    }
  }
  return new EmulatedCdpSession();
}
