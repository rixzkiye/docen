import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { FONT_NAMES, FONT_SIZES_PT } from "../../../document/font-lists";
import { observeLang, t } from "../../i18n/localize";
// Side-effect import: registers <docen-ribbon-combobox> so the mini toolbar's
// typeable font/size controls work even where the ribbon bundle is absent.
import "../ribbon/ribbon-combobox";
import { renderIcon } from "../ribbon/command-helpers";

/** Word's point-size range (the Font dialog's accepted span). */
const MIN_SIZE_PT = 1;
const MAX_SIZE_PT = 1638;

const RECENT_FONTS_KEY = "docen.recent-fonts";
const RECENT_FONTS_MAX = 5;

/** Recently applied font families, most-recent first (Word's "Recently Used
 *  Fonts" group). Persisted so the list survives reloads; storage failures
 *  (private mode) degrade to an in-memory list. */
function recentFonts(): string[] {
  try {
    const raw = localStorage.getItem(RECENT_FONTS_KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : [];
    if (Array.isArray(parsed)) return parsed.filter((f): f is string => typeof f === "string");
  } catch {
    // ignore malformed/blocked storage
  }
  return [];
}

function rememberFont(font: string): void {
  const list = [font, ...recentFonts().filter((f) => f !== font)].slice(0, RECENT_FONTS_MAX);
  try {
    localStorage.setItem(RECENT_FONTS_KEY, JSON.stringify(list));
  } catch {
    // ignore blocked storage
  }
}

const styles = css`
  :host {
    position: fixed;
    z-index: 10000;
    display: none;
    align-items: center;
    gap: 2px;
    padding: 3px 6px;
    background: var(--docen-color-bg, #ffffff);
    border: 1px solid var(--docen-color-stroke-1, #d1d1d1);
    border-radius: 4px;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.16);
    font-family: "Segoe UI", system-ui, sans-serif;
    font-size: 12px;
    color: var(--docen-color-text-1, #242424);
    pointer-events: auto;
    opacity: 0;
    transition: opacity 0.15s ease-out;
    user-select: none;
    -webkit-user-select: none;
  }

  :host([data-open]) {
    display: flex;
  }

  .btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 24px;
    height: 24px;
    padding: 0;
    margin: 0;
    border: 1px solid transparent;
    border-radius: 3px;
    background: transparent;
    color: inherit;
    cursor: pointer;
  }

  .btn:hover {
    background: var(--docen-color-subtle-background-hover, #f0f0f0);
    border-color: var(--docen-color-stroke-2, #d1d1d1);
  }

  .btn[aria-pressed="true"],
  .btn.active {
    background: var(--docen-color-subtle-background-selected, #e5e5e5);
    color: var(--docen-color-accent, #0f6cbd);
    border-color: var(--docen-color-accent, #0f6cbd);
  }

  .btn[aria-disabled="true"] {
    opacity: 0.4;
    cursor: default;
  }

  .btn svg {
    display: block;
    width: 16px;
    height: 16px;
    fill: currentColor;
  }

  .sep {
    width: 1px;
    height: 16px;
    margin: 0 2px;
    background: var(--docen-color-divider, #e1e1e1);
  }

  .color-wrap {
    position: relative;
    display: inline-flex;
    align-items: center;
  }

  .color-input {
    position: absolute;
    inset: 0;
    opacity: 0;
    width: 100%;
    height: 100%;
    cursor: pointer;
    border: none;
    padding: 0;
  }

  .color-indicator {
    position: absolute;
    bottom: 2px;
    left: 4px;
    right: 4px;
    height: 3px;
    border-radius: 1px;
    background: #000000;
  }

  .highlight-indicator {
    background: #ffff00;
  }
`;

const template = html<DocenMiniToolbar>`
  <docen-ribbon-combobox
    class="font-combo"
    ${ref("fontCombo")}
    event="font-name"
    style="width:132px;flex:none"
  ></docen-ribbon-combobox>

  <docen-ribbon-combobox
    class="size-combo"
    ${ref("sizeCombo")}
    event="font-size"
    size="short"
    style="width:56px;flex:none"
  ></docen-ribbon-combobox>

  <button type="button" class="btn" ${ref("growBtn")} data-cmd="grow-font">
    <svg viewBox="0 0 16 16" width="16" height="16" fill="currentColor">
      <path
        d="M7.5 2.5l-4 9.5h1.7l.9-2.2h3.8l.9 2.2h1.7l-4-9.5h-1zm-.1 2.3l1.4 3.6H6l1.4-3.6zM13 3l-2.5 3h5L13 3z"
      />
    </svg>
  </button>

  <button type="button" class="btn" ${ref("shrinkBtn")} data-cmd="shrink-font">
    <svg viewBox="0 0 16 16" width="16" height="16" fill="currentColor">
      <path
        d="M7 4.5l-3.5 8h1.5l.8-1.9h3.4l.8 1.9h1.5l-3.5-8h-1zm-.1 2l1.2 3.1H5.7l1.2-3.1zM12.5 9l-2-2.5h4L12.5 9z"
      />
    </svg>
  </button>

  <span class="sep"></span>

  <button type="button" class="btn" ${ref("boldBtn")} data-cmd="bold">
    <span ${ref("boldIcon")}></span>
  </button>

  <button type="button" class="btn" ${ref("italicBtn")} data-cmd="italic">
    <span ${ref("italicIcon")}></span>
  </button>

  <button type="button" class="btn" ${ref("underlineBtn")} data-cmd="underline">
    <span ${ref("underlineIcon")}></span>
  </button>

  <span class="sep"></span>

  <div class="color-wrap">
    <button type="button" class="btn" ${ref("textColorBtn")} data-cmd="font-color">
      <span ${ref("fontColorIcon")}></span>
      <span class="color-indicator" ${ref("textColorBar")}></span>
    </button>
    <input type="color" class="color-input" ${ref("textColorInput")} value="#000000" />
  </div>

  <div class="color-wrap">
    <button type="button" class="btn" ${ref("highlightBtn")} data-cmd="highlight">
      <span ${ref("highlightIcon")}></span>
      <span class="color-indicator highlight-indicator" ${ref("highlightColorBar")}></span>
    </button>
    <input type="color" class="color-input" ${ref("highlightInput")} value="#ffff00" />
  </div>

  <span class="sep"></span>

  <button type="button" class="btn" ${ref("bulletBtn")} data-cmd="bullet-list">
    <span ${ref("bulletIcon")}></span>
  </button>

  <button type="button" class="btn" ${ref("formatPainterBtn")} data-cmd="format-painter">
    <span ${ref("painterIcon")}></span>
  </button>
`;

/**
 * `<docen-mini-toolbar>` — Word-style floating mini toolbar that appears near
 * a text selection, fading in smoothly and fading out as the pointer moves
 * away. Font family and size are typeable comboboxes sharing the ribbon's
 * font catalog (plus theme fonts and recently used families); every other
 * control acts on the selection, and formatting state follows the caret.
 */
@customElement({ name: "docen-mini-toolbar", template, styles })
export class DocenMiniToolbar extends FASTElement {
  @observable fontCombo?: HTMLElement;
  @observable sizeCombo?: HTMLElement;
  @observable growBtn?: HTMLButtonElement;
  @observable shrinkBtn?: HTMLButtonElement;
  @observable boldBtn?: HTMLButtonElement;
  @observable italicBtn?: HTMLButtonElement;
  @observable underlineBtn?: HTMLButtonElement;
  @observable textColorBtn?: HTMLButtonElement;
  @observable textColorInput?: HTMLInputElement;
  @observable textColorBar?: HTMLElement;
  @observable highlightBtn?: HTMLButtonElement;
  @observable highlightInput?: HTMLInputElement;
  @observable highlightColorBar?: HTMLElement;
  @observable bulletBtn?: HTMLButtonElement;
  @observable formatPainterBtn?: HTMLButtonElement;

  @observable boldIcon?: HTMLElement;
  @observable italicIcon?: HTMLElement;
  @observable underlineIcon?: HTMLElement;
  @observable fontColorIcon?: HTMLElement;
  @observable highlightIcon?: HTMLElement;
  @observable bulletIcon?: HTMLElement;
  @observable painterIcon?: HTMLElement;

  #isOpen = false;
  #targetRect: { left: number; top: number; right: number; bottom: number } | null = null;
  #fadeTimeout = 0;
  #unobserveLang?: () => void;
  /** The last font/size pushed into the comboboxes — avoids re-seeding the
   *  option lists while the user is picking in them. */
  #lastFont = "";
  #lastSize = "";

  get isOpen(): boolean {
    return this.#isOpen;
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#renderIcons();
    this.#bindEvents();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
    this.setAttribute("role", "toolbar");
  }

  disconnectedCallback(): void {
    this.#unbindEvents();
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    if (this.#fadeTimeout) clearTimeout(this.#fadeTimeout);
    super.disconnectedCallback();
  }

  #fontItems(current?: string): string {
    const seen = new Set<string>();
    const names: string[] = [];
    for (const name of [
      ...(current ? [current] : []),
      ...recentFonts(),
      ...this.#themeFonts(),
      ...FONT_NAMES,
    ]) {
      if (name && !seen.has(name)) {
        seen.add(name);
        names.push(name);
      }
    }
    return JSON.stringify(names.map((text) => ({ text })));
  }

  /** The document theme's heading/body families (host CSS variables) — Word
   *  lists these at the top of the font gallery. */
  #themeFonts(): string[] {
    const style = getComputedStyle(this);
    return ["--docen-theme-font-major", "--docen-theme-font-minor"]
      .map((name) =>
        (style.getPropertyValue(name) || this.style.getPropertyValue(name))
          .trim()
          .replace(/^["']|["']$/g, ""),
      )
      .filter(Boolean);
  }

  #sizeItems(current?: string): string {
    const seen = new Set<string>();
    const sizes: string[] = [];
    for (const size of [...(current ? [current] : []), ...FONT_SIZES_PT.map(String)]) {
      if (!seen.has(size)) {
        seen.add(size);
        sizes.push(size);
      }
    }
    return JSON.stringify(sizes.map((text) => ({ text, value: text })));
  }

  #renderIcons(): void {
    if (this.boldIcon) renderIcon(this.boldIcon, "bold");
    if (this.italicIcon) renderIcon(this.italicIcon, "italic");
    if (this.underlineIcon) renderIcon(this.underlineIcon, "underline");
    if (this.fontColorIcon) renderIcon(this.fontColorIcon, "font-color");
    if (this.highlightIcon) renderIcon(this.highlightIcon, "highlight");
    if (this.bulletIcon) renderIcon(this.bulletIcon, "list");
    if (this.painterIcon) renderIcon(this.painterIcon, "format-painter");
  }

  /** Localized tooltips and accessible names for every control — the mini
   *  toolbar ships its own component-local translations so it works wherever
   *  the ribbon's business tables are not loaded. */
  #applyLabels(): void {
    const name = (el: Element | undefined, key: string): void => {
      if (!el) return;
      const text = t(key, this);
      el.setAttribute("aria-label", text);
      el.setAttribute("title", text);
    };
    name(this.boldBtn, "miniToolbar.bold");
    name(this.italicBtn, "miniToolbar.italic");
    name(this.underlineBtn, "miniToolbar.underline");
    name(this.growBtn, "miniToolbar.growFont");
    name(this.shrinkBtn, "miniToolbar.shrinkFont");
    name(this.textColorBtn, "miniToolbar.fontColor");
    name(this.highlightBtn, "miniToolbar.highlight");
    name(this.bulletBtn, "miniToolbar.bulletList");
    name(this.formatPainterBtn, "miniToolbar.formatPainter");
    this.setAttribute("aria-label", t("miniToolbar.label", this));
    const fontLabel = t("miniToolbar.fontName", this);
    this.fontCombo?.setAttribute("aria-label", fontLabel);
    this.fontCombo?.setAttribute("title", fontLabel);
    this.fontCombo?.setAttribute("label", fontLabel);
    const sizeLabel = t("miniToolbar.fontSize", this);
    this.sizeCombo?.setAttribute("aria-label", sizeLabel);
    this.sizeCombo?.setAttribute("title", sizeLabel);
    this.sizeCombo?.setAttribute("label", sizeLabel);
    name(this.textColorInput, "miniToolbar.fontColorPicker");
    name(this.highlightInput, "miniToolbar.highlightPicker");
  }

  #bindEvents(): void {
    this.addEventListener("mousedown", this.#onMouseDown);
    this.addEventListener("command", this.#onComboCommand);
    document.addEventListener("pointermove", this.#onDocPointerMove);

    const fontCombo = this.fontCombo ?? this.shadowRoot?.querySelector<HTMLElement>(".font-combo");
    if (fontCombo) {
      this.#lastFont = "";
      fontCombo.setAttribute("items", this.#fontItems());
    }
    const sizeCombo = this.sizeCombo ?? this.shadowRoot?.querySelector<HTMLElement>(".size-combo");
    if (sizeCombo) {
      this.#lastSize = "";
      sizeCombo.setAttribute("items", this.#sizeItems());
    }

    const textColorInput =
      this.textColorInput ?? this.shadowRoot?.querySelector<HTMLInputElement>(".color-input");
    const textColorBar =
      this.textColorBar ??
      this.shadowRoot?.querySelector<HTMLElement>(".color-indicator:not(.highlight-indicator)");
    textColorInput?.addEventListener("input", () => {
      const val = textColorInput.value ?? "#000000";
      if (textColorBar) textColorBar.style.backgroundColor = val;
      this.#emit("font-color", val.replace("#", ""));
    });

    const highlightInput =
      this.highlightInput ??
      this.shadowRoot?.querySelector<HTMLInputElement>(".color-wrap:nth-of-type(2) .color-input");
    const highlightColorBar =
      this.highlightColorBar ?? this.shadowRoot?.querySelector<HTMLElement>(".highlight-indicator");
    highlightInput?.addEventListener("input", () => {
      const val = highlightInput.value ?? "#ffff00";
      if (highlightColorBar) highlightColorBar.style.backgroundColor = val;
      this.#emit("highlight", val.replace("#", ""));
    });

    this.shadowRoot?.querySelectorAll<HTMLButtonElement>("button[data-cmd]").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        if (btn.getAttribute("aria-disabled") === "true") return;
        const cmd = btn.dataset.cmd;
        if (cmd) {
          if (cmd === "format-painter") {
            this.#emit("copy-format");
          } else {
            this.#emit(cmd);
          }
        }
      });
    });
  }

  #unbindEvents(): void {
    this.removeEventListener("mousedown", this.#onMouseDown);
    this.removeEventListener("command", this.#onComboCommand);
    document.removeEventListener("pointermove", this.#onDocPointerMove);
  }

  /** Prevent toolbar button clicks from blurring the editor and collapsing the
   *  selection. Focusable controls (the typeable comboboxes, the color
   *  pickers) must keep their default mousedown so they can take focus and
   *  open. */
  #onMouseDown = (e: MouseEvent): void => {
    const target = e.composedPath()[0] as HTMLElement | undefined;
    if (target?.tagName === "BUTTON" && target.hasAttribute("data-cmd")) e.preventDefault();
  };

  /** Commands from the two comboboxes: apply Word's point-size clamp, keep the
   *  recently-used font list fresh, and re-emit from the toolbar so the host
   *  sees one event source (the child's own event never escapes the toolbar). */
  #onComboCommand = (e: Event): void => {
    // `e.target` is retargeted to this host for events observed outside the
    // shadow tree, so identify our own re-emission by the original target.
    if (e.composedPath()[0] === this) return;
    const detail = (e as CustomEvent<{ event?: string; value?: string }>).detail;
    if (!detail?.event) return;
    // The child's raw event must not escape the toolbar (immediate: listeners
    // on this host itself are skipped too, so the clamp can't be bypassed).
    e.stopImmediatePropagation();
    if (detail.event === "font-name") {
      const font = (detail.value ?? "").trim();
      if (!font) return;
      rememberFont(font);
      const combo = this.fontCombo ?? this.shadowRoot?.querySelector<HTMLElement>(".font-combo");
      combo?.setAttribute("items", this.#fontItems(font));
      this.#emit("font-name", font);
      return;
    }
    if (detail.event === "font-size") {
      const text = (detail.value ?? "").trim();
      if (!text) return;
      const raw = Number(text);
      if (!Number.isFinite(raw)) return;
      const size = Math.min(MAX_SIZE_PT, Math.max(MIN_SIZE_PT, raw));
      // Reflect a clamped/mis-typed value back into the box immediately.
      this.#lastSize = String(size);
      this.sizeCombo?.setAttribute("value", String(size));
      this.#emit("font-size", String(size));
    }
  };

  /** Smooth fade-out as pointer moves away from the toolbar and selection.
   *  While a control holds focus (typing in a combo, picking a color) the
   *  toolbar stays put — Word keeps the mini toolbar anchored during use. */
  #onDocPointerMove = (e: PointerEvent): void => {
    if (!this.#isOpen) return;
    const active = this.shadowRoot?.activeElement;
    if (active && this.shadowRoot?.contains(active)) return;

    const tbRect = this.getBoundingClientRect();
    if (tbRect.width === 0 || tbRect.height === 0) return;

    // Distance to toolbar rect
    const dx = Math.max(tbRect.left - e.clientX, 0, e.clientX - tbRect.right);
    const dy = Math.max(tbRect.top - e.clientY, 0, e.clientY - tbRect.bottom);
    const distToTb = Math.hypot(dx, dy);

    // Also consider distance to original selection rect so hovering near the text keeps it
    let distToSel = Infinity;
    if (this.#targetRect) {
      const sx = Math.max(this.#targetRect.left - e.clientX, 0, e.clientX - this.#targetRect.right);
      const sy = Math.max(this.#targetRect.top - e.clientY, 0, e.clientY - this.#targetRect.bottom);
      distToSel = Math.hypot(sx, sy);
    }

    const dist = Math.min(distToTb, distToSel);

    if (dist <= 40) {
      this.style.opacity = "1";
    } else if (dist <= 160) {
      const factor = 1 - (dist - 40) / 120;
      this.style.opacity = String(Math.max(0.1, factor));
    } else {
      this.hide();
    }
  };

  showNear(rect: { left: number; top: number; right: number; bottom: number }): void {
    this.#targetRect = rect;
    this.#isOpen = true;
    this.setAttribute("data-open", "");

    // Measure after the display flip so the placement uses the real box.
    const tbWidth = this.offsetWidth || 420;
    const tbHeight = this.offsetHeight || 32;

    // Position above selection if space permits, otherwise below
    let top = rect.top - tbHeight - 8;
    if (top < 10) {
      top = rect.bottom + 8;
    }

    // Horizontally center or align to selection left
    let left = rect.left + (rect.right - rect.left) / 2 - tbWidth / 2;
    // Viewport bounds clamping
    left = Math.max(10, Math.min(window.innerWidth - tbWidth - 10, left));
    top = Math.max(10, Math.min(window.innerHeight - tbHeight - 10, top));

    this.style.left = `${Math.round(left)}px`;
    this.style.top = `${Math.round(top)}px`;

    // Smooth fade-in
    this.style.opacity = "0";
    requestAnimationFrame(() => {
      this.style.opacity = "0.9";
    });
  }

  hide(): void {
    if (!this.#isOpen) return;
    this.#isOpen = false;
    this.style.opacity = "0";
    this.#targetRect = null;
    if (this.#fadeTimeout) clearTimeout(this.#fadeTimeout);
    this.#fadeTimeout = window.setTimeout(() => {
      if (!this.#isOpen) {
        this.removeAttribute("data-open");
      }
    }, 160);
  }

  updateFormatting(state: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontName?: string;
    fontSize?: string;
    fontColor?: string;
    highlightColor?: string;
    /** False in read-only/protected contexts — every control greys out. */
    editable?: boolean;
  }): void {
    const setPressed = (el: HTMLElement | undefined, on?: boolean): void => {
      if (!el) return;
      el.setAttribute("aria-pressed", on ? "true" : "false");
      el.classList.toggle("active", Boolean(on));
    };
    setPressed(this.boldBtn, state.bold);
    setPressed(this.italicBtn, state.italic);
    setPressed(this.underlineBtn, state.underline);

    const editable = state.editable !== false;
    const size = Number(state.fontSize);
    const setDisabled = (el: HTMLButtonElement | undefined, off: boolean): void => {
      if (!el) return;
      el.setAttribute("aria-disabled", off ? "true" : "false");
    };
    setDisabled(this.growBtn, !editable || (Number.isFinite(size) && size >= MAX_SIZE_PT));
    setDisabled(this.shrinkBtn, !editable || (Number.isFinite(size) && size <= MIN_SIZE_PT));

    // Word clears the boxes for a mixed selection — never leave a stale value
    // from the previous caret position.
    const fontCombo = this.fontCombo ?? this.shadowRoot?.querySelector<HTMLElement>(".font-combo");
    const font = state.fontName?.trim() ?? "";
    if (fontCombo && font !== this.#lastFont) {
      this.#lastFont = font;
      fontCombo.setAttribute("items", this.#fontItems(font || undefined));
      if (font) fontCombo.setAttribute("value", font);
      else fontCombo.removeAttribute("value");
    }
    const sizeCombo = this.sizeCombo ?? this.shadowRoot?.querySelector<HTMLElement>(".size-combo");
    const sizeText = state.fontSize?.trim() ?? "";
    if (sizeCombo && sizeText !== this.#lastSize) {
      this.#lastSize = sizeText;
      sizeCombo.setAttribute("items", this.#sizeItems(sizeText || undefined));
      if (sizeText) sizeCombo.setAttribute("value", sizeText);
      else sizeCombo.removeAttribute("value");
    }

    if (this.textColorBar && state.fontColor) {
      this.textColorBar.style.backgroundColor = state.fontColor.startsWith("#")
        ? state.fontColor
        : `#${state.fontColor}`;
    }
    if (this.highlightColorBar && state.highlightColor) {
      this.highlightColorBar.style.backgroundColor = state.highlightColor.startsWith("#")
        ? state.highlightColor
        : `#${state.highlightColor}`;
    }
  }

  #emit(cmd: string, value?: string): void {
    this.dispatchEvent(
      new CustomEvent("command", {
        bubbles: true,
        composed: true,
        detail: {
          event: cmd,
          value,
          source: this,
        },
      }),
    );
  }
}

if (!customElements.get("docen-mini-toolbar")) {
  customElements.define("docen-mini-toolbar", DocenMiniToolbar);
}

export default DocenMiniToolbar;
