import {
  FASTElement,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { renderIcon } from "../ribbon/command-helpers";

const FONTS = [
  "Calibri",
  "Arial",
  "Times New Roman",
  "Segoe UI",
  "Georgia",
  "Tahoma",
  "Verdana",
  "Courier New",
];

const FONT_SIZES = [
  "8",
  "9",
  "10",
  "11",
  "12",
  "14",
  "16",
  "18",
  "20",
  "24",
  "28",
  "36",
  "48",
  "72",
];

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

  .btn svg {
    display: block;
    width: 16px;
    height: 16px;
    fill: currentColor;
  }

  .select {
    height: 24px;
    border: 1px solid var(--docen-color-stroke-1, #d1d1d1);
    border-radius: 3px;
    background: var(--docen-color-bg, #ffffff);
    color: inherit;
    font-size: 11px;
    font-family: inherit;
    padding: 0 2px;
    margin: 0;
    cursor: pointer;
  }

  .select:focus {
    outline: 1px solid var(--docen-color-accent, #0f6cbd);
  }

  .font-select {
    width: 100px;
  }

  .size-select {
    width: 44px;
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
  <select class="select font-select" ${ref("fontSelect")} title="Font family">
    ${repeat(() => FONTS, html`<option :value="${(x) => x}">${(x) => x}</option>`)}
  </select>

  <select class="select size-select" ${ref("sizeSelect")} title="Font size">
    ${repeat(() => FONT_SIZES, html`<option :value="${(x) => x}">${(x) => x}</option>`)}
  </select>

  <button type="button" class="btn" ${ref("growBtn")} data-cmd="grow-font" title="Grow font">
    <svg viewBox="0 0 16 16" width="16" height="16" fill="currentColor">
      <path
        d="M7.5 2.5l-4 9.5h1.7l.9-2.2h3.8l.9 2.2h1.7l-4-9.5h-1zm-.1 2.3l1.4 3.6H6l1.4-3.6zM13 3l-2.5 3h5L13 3z"
      />
    </svg>
  </button>

  <button type="button" class="btn" ${ref("shrinkBtn")} data-cmd="shrink-font" title="Shrink font">
    <svg viewBox="0 0 16 16" width="16" height="16" fill="currentColor">
      <path
        d="M7 4.5l-3.5 8h1.5l.8-1.9h3.4l.8 1.9h1.5l-3.5-8h-1zm-.1 2l1.2 3.1H5.7l1.2-3.1zM12.5 9l-2-2.5h4L12.5 9z"
      />
    </svg>
  </button>

  <span class="sep"></span>

  <button type="button" class="btn" ${ref("boldBtn")} data-cmd="bold" title="Bold">
    <span ${ref("boldIcon")}></span>
  </button>

  <button type="button" class="btn" ${ref("italicBtn")} data-cmd="italic" title="Italic">
    <span ${ref("italicIcon")}></span>
  </button>

  <button type="button" class="btn" ${ref("underlineBtn")} data-cmd="underline" title="Underline">
    <span ${ref("underlineIcon")}></span>
  </button>

  <span class="sep"></span>

  <div class="color-wrap">
    <button
      type="button"
      class="btn"
      ${ref("textColorBtn")}
      data-cmd="font-color"
      title="Font color"
    >
      <span ${ref("fontColorIcon")}></span>
      <span class="color-indicator" ${ref("textColorBar")}></span>
    </button>
    <input
      type="color"
      class="color-input"
      ${ref("textColorInput")}
      value="#000000"
      title="Choose font color"
    />
  </div>

  <div class="color-wrap">
    <button
      type="button"
      class="btn"
      ${ref("highlightBtn")}
      data-cmd="highlight"
      title="Highlight color"
    >
      <span ${ref("highlightIcon")}></span>
      <span class="color-indicator highlight-indicator" ${ref("highlightColorBar")}></span>
    </button>
    <input
      type="color"
      class="color-input"
      ${ref("highlightInput")}
      value="#ffff00"
      title="Choose highlight color"
    />
  </div>

  <span class="sep"></span>

  <button type="button" class="btn" ${ref("bulletBtn")} data-cmd="bullet-list" title="Bullet list">
    <span ${ref("bulletIcon")}></span>
  </button>

  <button
    type="button"
    class="btn"
    ${ref("formatPainterBtn")}
    data-cmd="format-painter"
    title="Format painter"
  >
    <span ${ref("painterIcon")}></span>
  </button>
`;

/**
 * `<docen-mini-toolbar>` — Word-style floating mini toolbar that appears near
 * a text selection, fading in smoothly and fading out as the pointer moves away.
 */
@customElement({ name: "docen-mini-toolbar", template, styles })
export class DocenMiniToolbar extends FASTElement {
  @observable fontSelect?: HTMLSelectElement;
  @observable sizeSelect?: HTMLSelectElement;
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

  get isOpen(): boolean {
    return this.#isOpen;
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#populateOptions();
    this.#renderIcons();
    this.#bindEvents();
  }

  disconnectedCallback(): void {
    this.#unbindEvents();
    if (this.#fadeTimeout) clearTimeout(this.#fadeTimeout);
    super.disconnectedCallback();
  }

  #populateOptions(): void {
    const fontSelect =
      this.fontSelect ?? this.shadowRoot?.querySelector<HTMLSelectElement>(".font-select");
    if (fontSelect && fontSelect.options.length === 0) {
      for (const font of FONTS) {
        const opt = document.createElement("option");
        opt.value = font;
        opt.textContent = font;
        fontSelect.appendChild(opt);
      }
    }
    const sizeSelect =
      this.sizeSelect ?? this.shadowRoot?.querySelector<HTMLSelectElement>(".size-select");
    if (sizeSelect && sizeSelect.options.length === 0) {
      for (const size of FONT_SIZES) {
        const opt = document.createElement("option");
        opt.value = size;
        opt.textContent = size;
        sizeSelect.appendChild(opt);
      }
    }
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

  #bindEvents(): void {
    this.addEventListener("mousedown", this.#onMouseDown);
    document.addEventListener("pointermove", this.#onDocPointerMove);

    const fontSelect =
      this.fontSelect ?? this.shadowRoot?.querySelector<HTMLSelectElement>(".font-select");
    fontSelect?.addEventListener("change", () => {
      this.#emit("font-name", fontSelect.value);
    });

    const sizeSelect =
      this.sizeSelect ?? this.shadowRoot?.querySelector<HTMLSelectElement>(".size-select");
    sizeSelect?.addEventListener("change", () => {
      this.#emit("font-size", sizeSelect.value);
    });

    const textColorInput =
      this.textColorInput ??
      this.shadowRoot?.querySelector<HTMLInputElement>(".color-wrap:nth-of-type(1) .color-input");
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
    document.removeEventListener("pointermove", this.#onDocPointerMove);
  }

  /** Prevent toolbar clicks from blurring the editor and collapsing selection. */
  #onMouseDown = (e: MouseEvent): void => {
    e.preventDefault();
  };

  /** Smooth fade-out as pointer moves away from the toolbar and selection. */
  #onDocPointerMove = (e: PointerEvent): void => {
    if (!this.#isOpen) return;

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

    // Layout dimensions
    const tbWidth = 340;
    const tbHeight = 34;

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
  }): void {
    if (this.boldBtn) {
      this.boldBtn.setAttribute("aria-pressed", state.bold ? "true" : "false");
      this.boldBtn.classList.toggle("active", Boolean(state.bold));
    }
    if (this.italicBtn) {
      this.italicBtn.setAttribute("aria-pressed", state.italic ? "true" : "false");
      this.italicBtn.classList.toggle("active", Boolean(state.italic));
    }
    if (this.underlineBtn) {
      this.underlineBtn.setAttribute("aria-pressed", state.underline ? "true" : "false");
      this.underlineBtn.classList.toggle("active", Boolean(state.underline));
    }
    const fontSelect =
      this.fontSelect ?? this.shadowRoot?.querySelector<HTMLSelectElement>(".font-select");
    if (fontSelect && state.fontName) {
      fontSelect.value = state.fontName;
    }
    const sizeSelect =
      this.sizeSelect ?? this.shadowRoot?.querySelector<HTMLSelectElement>(".size-select");
    if (sizeSelect && state.fontSize) {
      sizeSelect.value = state.fontSize;
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
