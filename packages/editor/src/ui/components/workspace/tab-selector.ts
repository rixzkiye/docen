import { FASTElement, attr, css, customElement, html, observable } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export type TabSelectorType =
  | "left"
  | "center"
  | "right"
  | "decimal"
  | "bar"
  | "first-line"
  | "hanging";

export const TAB_SELECTOR_TYPES: TabSelectorType[] = [
  "left",
  "center",
  "right",
  "decimal",
  "bar",
  "first-line",
  "hanging",
];

const styles = css`
  :host {
    display: block;
    width: 20px;
    height: 20px;
    box-sizing: border-box;
    user-select: none;
    border-right: 1px solid var(--docen-ruler-border-strong, #c8c8c8);
    border-bottom: 1px solid var(--docen-ruler-border-strong, #c8c8c8);
  }

  .selector-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 100%;
    height: 100%;
    padding: 0;
    margin: 0;
    border: none;
    background: transparent;
    cursor: pointer;
    color: #404040;
    outline: none;
  }

  .selector-btn:hover {
    background: rgba(0, 0, 0, 0.06);
    color: #0f6cbd;
  }

  .selector-btn:active {
    background: rgba(0, 0, 0, 0.12);
  }

  .selector-btn:focus-visible {
    outline: 1px solid #0f6cbd;
    outline-offset: -1px;
  }

  .tab-icon {
    width: 12px;
    height: 12px;
    pointer-events: none;
  }
`;

function tabGlyph(type: TabSelectorType): ReturnType<typeof html> {
  switch (type) {
    case "left":
      return html`<path
        d="M3 2 v8 h6"
        fill="none"
        stroke="currentColor"
        stroke-width="1.5"
        stroke-linecap="square"
      />`;
    case "center":
      return html`<path
        d="M2 10 h8 M6 10 v-8"
        fill="none"
        stroke="currentColor"
        stroke-width="1.5"
        stroke-linecap="square"
      />`;
    case "right":
      return html`<path
        d="M9 2 v8 h-6"
        fill="none"
        stroke="currentColor"
        stroke-width="1.5"
        stroke-linecap="square"
      />`;
    case "decimal":
      return html`
        <path
          d="M2 10 h6 M5 10 v-8"
          fill="none"
          stroke="currentColor"
          stroke-width="1.5"
          stroke-linecap="square"
        />
        <circle cx="10" cy="9.5" r="1" fill="currentColor" />
      `;
    case "bar":
      return html`<path
        d="M6 1 v10"
        fill="none"
        stroke="currentColor"
        stroke-width="1.5"
        stroke-linecap="square"
      />`;
    case "first-line":
      return html`<polygon points="2,3 10,3 6,8" fill="currentColor" />`;
    case "hanging":
      return html`<polygon points="6,4 2,9 10,9" fill="currentColor" />`;
    default:
      return html`<path
        d="M3 2 v8 h6"
        fill="none"
        stroke="currentColor"
        stroke-width="1.5"
        stroke-linecap="square"
      />`;
  }
}

const template = html<DocenTabSelector>`
  <button
    type="button"
    class="selector-btn"
    part="button"
    title="${(x) => x.titleText}"
    aria-label="${(x) => x.titleText}"
    @click="${(x) => x.cycle()}"
  >
    <svg class="tab-icon" viewBox="0 0 12 12" aria-hidden="true">
      ${(x) => tabGlyph(x.activeType)}
    </svg>
  </button>
`;

/**
 * `<docen-tab-selector>` — Word's top-left corner ruler button: a 20px square
 * that sits at the intersection of the horizontal and vertical rulers. Clicking
 * it cycles through tab-stop types (Left, Center, Right, Decimal, Bar) and
 * paragraph indents (First-Line, Hanging).
 */
@customElement({ name: "docen-tab-selector", template, styles })
export class DocenTabSelector extends FASTElement {
  @attr({ attribute: "active-type" })
  @observable
  activeType: TabSelectorType = "left";

  #unobserveLang?: () => void;

  get titleText(): string {
    switch (this.activeType) {
      case "left":
        return t("ruler.tab.left", this) || "Left Tab";
      case "center":
        return t("ruler.tab.center", this) || "Center Tab";
      case "right":
        return t("ruler.tab.right", this) || "Right Tab";
      case "decimal":
        return t("ruler.tab.decimal", this) || "Decimal Tab";
      case "bar":
        return t("ruler.tab.bar", this) || "Bar Tab";
      case "first-line":
        return t("ruler.firstLine", this) || "First Line Indent";
      case "hanging":
        return t("ruler.hanging", this) || "Hanging Indent";
      default:
        return "Left Tab";
    }
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unobserveLang = observeLang(() => {
      this.$emit("tab-selector:lang-change");
    });
  }

  override disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  cycle(): void {
    const idx = TAB_SELECTOR_TYPES.indexOf(this.activeType);
    const nextIdx = (idx + 1) % TAB_SELECTOR_TYPES.length;
    this.activeType = TAB_SELECTOR_TYPES[nextIdx]!;
    this.$emit("tab-selector:change", { type: this.activeType });
  }

  override $emit(
    type: string,
    detail?: unknown,
    options?: Omit<CustomEventInit, "detail">,
  ): boolean {
    return this.dispatchEvent(
      new CustomEvent(type, {
        bubbles: true,
        composed: true,
        detail,
        ...options,
      }),
    );
  }
}

export default DocenTabSelector;
