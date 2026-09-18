import {
  FASTElement,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface WrapOption {
  key: string;
  labelKey: string;
  iconSvg: string;
}

const WRAP_OPTIONS: readonly WrapOption[] = [
  {
    key: "inline",
    labelKey: "wrap.inline",
    iconSvg: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><rect x="3" y="5" width="18" height="2" fill="#777"/><rect x="3" y="11" width="7" height="2" fill="#777"/><rect x="11" y="8" width="5" height="8" fill="#2b7cd3"/><rect x="17" y="11" width="4" height="2" fill="#777"/><rect x="3" y="17" width="18" height="2" fill="#777"/></svg>`,
  },
  {
    key: "square",
    labelKey: "wrap.square",
    iconSvg: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><rect x="3" y="4" width="18" height="2" fill="#777"/><rect x="3" y="8" width="5" height="2" fill="#777"/><rect x="16" y="8" width="5" height="2" fill="#777"/><rect x="3" y="12" width="5" height="2" fill="#777"/><rect x="16" y="12" width="5" height="2" fill="#777"/><rect x="3" y="16" width="5" height="2" fill="#777"/><rect x="16" y="16" width="5" height="2" fill="#777"/><rect x="9" y="7" width="6" height="10" fill="#2b7cd3"/><rect x="3" y="20" width="18" height="2" fill="#777"/></svg>`,
  },
  {
    key: "tight",
    labelKey: "wrap.tight",
    iconSvg: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><rect x="3" y="4" width="18" height="2" fill="#777"/><rect x="3" y="8" width="6" height="2" fill="#777"/><rect x="15" y="8" width="6" height="2" fill="#777"/><rect x="3" y="12" width="5" height="2" fill="#777"/><rect x="16" y="12" width="5" height="2" fill="#777"/><rect x="3" y="16" width="6" height="2" fill="#777"/><rect x="15" y="16" width="6" height="2" fill="#777"/><circle cx="12" cy="12" r="4.5" fill="#2b7cd3"/><rect x="3" y="20" width="18" height="2" fill="#777"/></svg>`,
  },
  {
    key: "through",
    labelKey: "wrap.through",
    iconSvg: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><rect x="3" y="4" width="18" height="2" fill="#777"/><rect x="3" y="8" width="7" height="2" fill="#777"/><rect x="14" y="8" width="7" height="2" fill="#777"/><polygon points="12,7 8,17 16,17" fill="#2b7cd3"/><rect x="3" y="12" width="4" height="2" fill="#777"/><rect x="17" y="12" width="4" height="2" fill="#777"/><rect x="3" y="20" width="18" height="2" fill="#777"/></svg>`,
  },
  {
    key: "top-bottom",
    labelKey: "wrap.top-bottom",
    iconSvg: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><rect x="3" y="4" width="18" height="2" fill="#777"/><rect x="9" y="8" width="6" height="8" fill="#2b7cd3"/><rect x="3" y="18" width="18" height="2" fill="#777"/></svg>`,
  },
  {
    key: "behind",
    labelKey: "wrap.behind",
    iconSvg: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><rect x="8" y="7" width="8" height="10" fill="#c0d4ec"/><rect x="3" y="5" width="18" height="2" fill="#555"/><rect x="3" y="10" width="18" height="2" fill="#555"/><rect x="3" y="15" width="18" height="2" fill="#555"/></svg>`,
  },
  {
    key: "front",
    labelKey: "wrap.front",
    iconSvg: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><rect x="3" y="5" width="18" height="2" fill="#ccc"/><rect x="3" y="10" width="18" height="2" fill="#ccc"/><rect x="3" y="15" width="18" height="2" fill="#ccc"/><rect x="8" y="7" width="8" height="10" fill="#2b7cd3"/></svg>`,
  },
];

const styles = css`
  :host {
    display: block;
    position: absolute;
    z-index: 10;
    pointer-events: auto;
  }
  .flyout-btn {
    width: 24px;
    height: 24px;
    border-radius: 3px;
    background: #fff;
    border: 1px solid #c8c8c8;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.16);
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
    padding: 0;
    transition:
      background 0.1s ease,
      border-color 0.1s ease;
  }
  .flyout-btn:hover {
    background: #f3f3f3;
    border-color: #2b7cd3;
  }
  .flyout-btn.active {
    background: #e5f0fb;
    border-color: #2b7cd3;
  }
  .popover {
    position: absolute;
    top: 28px;
    right: 0;
    width: 220px;
    background: #fff;
    border: 1px solid #d1d1d1;
    border-radius: 4px;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.18);
    padding: 10px;
    display: none;
    flex-direction: column;
    gap: 8px;
    font-size: 12px;
    color: #333;
  }
  .popover.open {
    display: flex;
  }
  .section-title {
    font-weight: 600;
    font-size: 11px;
    text-transform: uppercase;
    color: #666;
    margin-bottom: 2px;
  }
  .grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 4px;
  }
  .wrap-item {
    width: 44px;
    height: 40px;
    border: 1px solid #e0e0e0;
    border-radius: 3px;
    background: #fafafa;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 0;
  }
  .wrap-item:hover {
    background: #f0f7ff;
    border-color: #2b7cd3;
  }
  .wrap-item.selected {
    background: #cce4f7;
    border-color: #0078d4;
    outline: 1px solid #0078d4;
  }
  .position-group {
    display: flex;
    flex-direction: column;
    gap: 6px;
    padding-top: 4px;
    border-top: 1px solid #ebebeb;
  }
  .radio-row {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
  .radio-row input {
    cursor: pointer;
    margin: 0;
  }
  .see-more {
    margin-top: 4px;
    padding-top: 6px;
    border-top: 1px solid #ebebeb;
    color: #0078d4;
    cursor: pointer;
    text-decoration: none;
    font-size: 11px;
  }
  .see-more:hover {
    text-decoration: underline;
  }
`;

const template = html<DocenLayoutOptionsFlyout>`
  <button
    class="flyout-btn ${(x) => (x.isOpen ? "active" : "")}"
    title="Layout Options"
    @click="${(x, c) => x.toggleOpen(c.event as MouseEvent)}"
    ${ref("btnEl")}
  >
    <svg width="16" height="16" viewBox="0 0 20 20" fill="none">
      <path d="M4 4a6 6 0 0 1 12 0v2H4V4z" fill="#2b7cd3" />
      <rect x="3" y="8" width="14" height="9" rx="1" fill="#fff" stroke="#555" stroke-width="1.2" />
      <line x1="6" y1="11" x2="14" y2="11" stroke="#777" stroke-width="1.2" />
      <line x1="6" y1="14" x2="14" y2="14" stroke="#777" stroke-width="1.2" />
    </svg>
  </button>

  <div class="popover ${(x) => (x.isOpen ? "open" : "")}" ${ref("popoverEl")}>
    <div class="section-title">${(x) => t("wrap.sectionWithTextWrapping", x)}</div>
    <div class="grid">
      ${repeat(
        () => WRAP_OPTIONS,
        html<WrapOption, DocenLayoutOptionsFlyout>`
          <button
            class="wrap-item ${(item, c) => (c.parent.isWrapSelected(item.key) ? "selected" : "")}"
            title="${(item, c) => t(item.labelKey, c.parent)}"
            @click="${(item, c) => c.parent.selectWrap(item.key)}"
            :innerHTML="${(item) => item.iconSvg}"
          ></button>
        `,
      )}
    </div>

    ${(x) =>
      x.isFloating
        ? html<DocenLayoutOptionsFlyout>`
            <div class="position-group">
              <div class="section-title">${(c) => t("wrap.sectionPosition", c)}</div>
              <label class="radio-row">
                <input
                  type="radio"
                  name="posMode"
                  value="moveWithText"
                  ?checked="${(c) => c.positionMode === "moveWithText"}"
                  @change="${(c) => c.selectPosition("moveWithText")}"
                />
                <span>${(c) => t("wrap.moveWithText", c)}</span>
              </label>
              <label class="radio-row">
                <input
                  type="radio"
                  name="posMode"
                  value="fixPosition"
                  ?checked="${(c) => c.positionMode === "fixPosition"}"
                  @change="${(c) => c.selectPosition("fixPosition")}"
                />
                <span>${(c) => t("wrap.fixPosition", c)}</span>
              </label>
            </div>
          `
        : ""}

    <a class="see-more" @click="${(x) => x.openMoreDialog()}"> ${(x) => t("wrap.seeMore", x)} </a>
  </div>
`;

@customElement({ name: "docen-layout-options-flyout", template, styles })
export default class DocenLayoutOptionsFlyout extends FASTElement {
  @observable wrapMode = "inline";
  @observable positionMode: "moveWithText" | "fixPosition" = "moveWithText";
  @observable isFloating = false;
  @observable isOpen = false;

  btnEl!: HTMLButtonElement;
  popoverEl!: HTMLDivElement;

  #unobserveLang?: () => void;
  #onDocClick = (e: MouseEvent): void => {
    if (!this.isOpen) return;
    if (!this.contains(e.target as Node)) {
      this.isOpen = false;
    }
  };

  override connectedCallback(): void {
    super.connectedCallback();
    document.addEventListener("pointerdown", this.#onDocClick);
    this.#unobserveLang = observeLang(() => {});
  }

  override disconnectedCallback(): void {
    document.removeEventListener("pointerdown", this.#onDocClick);
    this.#unobserveLang?.();
    super.disconnectedCallback();
  }

  toggleOpen(e: MouseEvent): void {
    e.stopPropagation();
    e.preventDefault();
    this.isOpen = !this.isOpen;
  }

  isWrapSelected(key: string): boolean {
    if (this.wrapMode === key) return true;
    if (key === "top-bottom" && (this.wrapMode === "topBottom" || this.wrapMode === "top-bottom")) {
      return true;
    }
    return false;
  }

  selectWrap(wrap: string): void {
    this.wrapMode = wrap;
    this.dispatchEvent(
      new CustomEvent("select-wrap", { detail: { wrap }, bubbles: true, composed: true }),
    );
  }

  selectPosition(mode: "moveWithText" | "fixPosition"): void {
    this.positionMode = mode;
    this.dispatchEvent(
      new CustomEvent("select-position", { detail: { mode }, bubbles: true, composed: true }),
    );
  }

  openMoreDialog(): void {
    this.isOpen = false;
    this.dispatchEvent(new CustomEvent("open-dialog", { bubbles: true, composed: true }));
  }
}
