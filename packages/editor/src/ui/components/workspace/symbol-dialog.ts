import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const NORMAL_SUBSETS: Record<string, string[]> = {
  common: [
    "©",
    "®",
    "™",
    "§",
    "¶",
    "†",
    "‡",
    "•",
    "…",
    "‰",
    "°",
    "′",
    "″",
    "℃",
    "℉",
    "№",
    "±",
    "×",
    "÷",
    "≈",
    "≠",
    "≤",
    "≥",
    "∞",
    "√",
    "∑",
    "∏",
    "∫",
    "π",
    "µ",
    "Ω",
    "∆",
    "→",
    "←",
    "↑",
    "↓",
    "↔",
    "⇐",
    "⇒",
    "⇑",
    "⇓",
    "⇔",
    "∼",
    "∝",
    "∴",
    "∵",
    "⊂",
    "⊃",
    "€",
    "£",
    "¥",
    "¢",
    "₩",
    "₽",
    "①",
    "②",
    "③",
    "④",
    "⑤",
    "⑥",
    "⑦",
    "⑧",
    "⑨",
    "⑩",
    "★",
    "☆",
    "♦",
    "♠",
    "♣",
    "♥",
    "♪",
    "♫",
    "☀",
    "☂",
    "✓",
    "✗",
    "☐",
    "☑",
    "◼",
    "◻",
  ],
  currency: [
    "$",
    "¢",
    "£",
    "¤",
    "¥",
    "֏",
    "؋",
    "৲",
    "৳",
    "৻",
    "૱",
    "௹",
    "฿",
    "៛",
    "₠",
    "₡",
    "₢",
    "₣",
    "₤",
    "₥",
    "₦",
    "₧",
    "₨",
    "₩",
    "₪",
    "₫",
    "€",
    "₭",
    "₮",
    "₯",
    "₰",
    "₱",
    "₲",
    "₳",
    "₴",
    "₵",
    "₸",
    "₹",
    "₺",
    "₼",
    "₽",
    "₾",
    "₿",
  ],
  math: [
    "∀",
    "∁",
    "∂",
    "∃",
    "∄",
    "∅",
    "∆",
    "∇",
    "∈",
    "∉",
    "∊",
    "∋",
    "∌",
    "∍",
    "∎",
    "∏",
    "∐",
    "∑",
    "−",
    "∓",
    "∔",
    "∕",
    "∖",
    "∗",
    "∘",
    "∙",
    "√",
    "∛",
    "∜",
    "∝",
    "∞",
    "∟",
    "∠",
    "∡",
    "∢",
    "∣",
    "∤",
    "∥",
    "∦",
    "∧",
    "∨",
    "∩",
    "∪",
    "∫",
    "∬",
    "∭",
    "∮",
    "∯",
    "∴",
    "∵",
    "∶",
    "∷",
    "∼",
    "∽",
    "∾",
    "∿",
    "≀",
    "≁",
    "≂",
    "≃",
    "≅",
    "≈",
    "≊",
    "≋",
    "≠",
    "≡",
    "≤",
    "≥",
    "≦",
    "≧",
    "≨",
    "≩",
    "≪",
    "≫",
    "≬",
    "⊂",
    "⊃",
    "⊆",
    "⊇",
    "⊕",
  ],
  arrows: [
    "←",
    "↑",
    "→",
    "↓",
    "↔",
    "↕",
    "↖",
    "↗",
    "↘",
    "↙",
    "↚",
    "↛",
    "↞",
    "↟",
    "↠",
    "↡",
    "↢",
    "↣",
    "↤",
    "↥",
    "↦",
    "↧",
    "↨",
    "↩",
    "↪",
    "↫",
    "↬",
    "↭",
    "↮",
    "↯",
    "↰",
    "↱",
    "⇐",
    "⇑",
    "⇒",
    "⇓",
    "⇔",
    "⇕",
    "⇖",
    "⇗",
    "⇘",
    "⇙",
    "⇚",
    "⇛",
    "⇜",
    "⇝",
    "⇞",
    "⇟",
  ],
  latin1: [
    "¡",
    "¢",
    "£",
    "¤",
    "¥",
    "¦",
    "§",
    "¨",
    "©",
    "ª",
    "«",
    "¬",
    "®",
    "¯",
    "°",
    "±",
    "²",
    "³",
    "´",
    "µ",
    "¶",
    "·",
    "¸",
    "¹",
    "º",
    "»",
    "¼",
    "½",
    "¾",
    "¿",
    "À",
    "Á",
    "Â",
    "Ã",
    "Ä",
    "Å",
    "Æ",
    "Ç",
    "È",
    "É",
    "Ê",
    "Ë",
    "Ì",
    "Í",
    "Î",
    "Ï",
    "Ð",
    "Ñ",
    "Ò",
    "Ó",
    "Ô",
    "Õ",
    "Ö",
    "×",
    "Ø",
    "Ù",
    "Ú",
    "Û",
    "Ü",
    "Ý",
    "Þ",
    "ß",
    "à",
    "á",
  ],
};

const SYMBOL_FONT_GLYPHS = [
  "Α",
  "Β",
  "Γ",
  "Δ",
  "Ε",
  "Ζ",
  "Η",
  "Θ",
  "Ι",
  "Κ",
  "Λ",
  "Μ",
  "Ν",
  "Ξ",
  "Ο",
  "Π",
  "Ρ",
  "Σ",
  "Τ",
  "Υ",
  "Φ",
  "Χ",
  "Ψ",
  "Ω",
  "α",
  "β",
  "γ",
  "δ",
  "ε",
  "ζ",
  "η",
  "θ",
  "ι",
  "κ",
  "λ",
  "μ",
  "ν",
  "ξ",
  "ο",
  "π",
  "ρ",
  "σ",
  "τ",
  "υ",
  "φ",
  "χ",
  "ψ",
  "ω",
  "∀",
  "∃",
  "∅",
  "∇",
  "∈",
  "∉",
  "∋",
  "∏",
  "∑",
  "−",
  "∓",
  "∗",
  "∘",
  "√",
  "∝",
  "∞",
  "∠",
  "∧",
  "∨",
  "∩",
  "∪",
  "∫",
  "∴",
  "∼",
  "≅",
  "≈",
  "≠",
  "≡",
  "≤",
  "≥",
  "⊂",
  "⊃",
  "⊆",
  "⊇",
  "⊕",
  "⊗",
  "⊥",
  "⋅",
  "⌈",
  "⌉",
  "⌊",
  "⌋",
  "〈",
  "〉",
  "◊",
  "♠",
  "♣",
  "♥",
];

const WINGDINGS_GLYPHS = [
  "✌",
  "✍",
  "✎",
  "✏",
  "✑",
  "✒",
  "✂",
  "✈",
  "✉",
  "⌛",
  "⏳",
  "☎",
  "☏",
  "✡",
  "☸",
  "☯",
  "✝",
  "✞",
  "✟",
  "✠",
  "✰",
  "✵",
  "✷",
  "✸",
  "✹",
  "✺",
  "✻",
  "✼",
  "✽",
  "✾",
  "✿",
  "❀",
  "❁",
  "❂",
  "❃",
  "❄",
  "❅",
  "❆",
  "❇",
  "❈",
  "❉",
  "❊",
  "❋",
  "✓",
  "✔",
  "✕",
  "✖",
  "✗",
  "✘",
  "✙",
  "✚",
  "✛",
  "✜",
  "✝",
  "✞",
  "✟",
  "✠",
  "✡",
  "✢",
  "✣",
  "✤",
  "✥",
  "✦",
  "✧",
];

const WEBDINGS_GLYPHS = [
  "🕷",
  "🕸",
  "🏔",
  "🏖",
  "🏗",
  "🏛",
  "🏙",
  "🏚",
  "🏠",
  "🏡",
  "🏢",
  "🏣",
  "🏤",
  "🏥",
  "🏦",
  "🏧",
  "🏨",
  "🏩",
  "🏪",
  "🏫",
  "🏬",
  "🏭",
  "🏮",
  "🏯",
  "🏰",
  "🛩",
  "🛰",
  "🚢",
  "🚤",
  "🛥",
  "🛳",
  "⛵",
  "⚓",
  "🚧",
  "⛽",
  "🚏",
  "🚦",
  "🚥",
  "🏁",
  "🚩",
  "🎌",
  "🗾",
  "🗿",
  "🕐",
  "🕑",
  "🕒",
  "🕓",
  "🕔",
];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(460px, 94vw);
  }
  .sym-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .controls-row {
    display: flex;
    gap: 10px;
    align-items: center;
  }
  .control-group {
    display: flex;
    align-items: center;
    gap: 6px;
    flex: 1;
  }
  .control-group label {
    font-size: 12px;
    color: var(--docen-color-foreground, #242424);
  }
  .control-group select {
    flex: 1;
    padding: 4px 6px;
    font-size: 12px;
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 4px;
  }
  .sym-grid {
    display: grid;
    grid-template-columns: repeat(8, 1fr);
    gap: 2px;
    max-height: 220px;
    overflow-y: auto;
    border: 1px solid var(--docen-color-divider, #e0e0e0);
    border-radius: 4px;
    padding: 4px;
  }
  .sym-grid button {
    aspect-ratio: 1;
    border: 1px solid transparent;
    border-radius: 4px;
    background: transparent;
    font-size: 18px;
    line-height: 1;
    cursor: pointer;
    font-family: inherit;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .sym-grid button:hover {
    background: var(--colorNeutralBackground1Hover, #f5f5f5);
  }
  .sym-grid button.selected {
    border-color: var(--colorBrandForeground1, #0f6cbd);
    background: var(--colorBrandBackground2, #ebf3fc);
  }
  .recents-title {
    font-size: 11px;
    font-weight: 600;
    color: var(--docen-color-foreground-secondary, #616161);
    margin-top: 2px;
  }
  .recents-grid {
    display: flex;
    gap: 4px;
    overflow-x: auto;
    padding: 2px 0;
  }
  .recents-grid button {
    width: 28px;
    height: 28px;
    flex-shrink: 0;
    border: 1px solid var(--docen-color-divider, #e0e0e0);
    border-radius: 3px;
    background: transparent;
    font-size: 15px;
    cursor: pointer;
  }
  .recents-grid button:hover {
    background: var(--colorNeutralBackground1Hover, #f5f5f5);
  }
  .sym-preview {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 6px 10px;
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    min-height: 48px;
    box-sizing: border-box;
  }
  .sym-preview .glyph {
    font-size: 30px;
    line-height: 1;
    min-width: 36px;
    text-align: center;
  }
  .sym-preview .codepoint {
    font-size: 12px;
    color: var(--colorNeutralForeground3, #616161);
    font-variant-numeric: tabular-nums;
  }
`;

const template = html<DocenSymbolDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="sym-body">
      <div class="controls-row">
        <div class="control-group">
          <label>${(x) => t("symbol.font", x)}:</label>
          <select ${ref("fontSelect")} @change="${(x) => x.onFontChange()}">
            <option value="normal">(normal text)</option>
            <option value="Symbol">Symbol</option>
            <option value="Wingdings">Wingdings</option>
            <option value="Webdings">Webdings</option>
          </select>
        </div>
        <div class="control-group" ${ref("subsetGroup")}>
          <label>${(x) => t("symbol.subset", x)}:</label>
          <select ${ref("subsetSelect")} @change="${(x) => x.onSubsetChange()}">
            <option value="common">Common Symbols</option>
            <option value="currency">Currency Symbols</option>
            <option value="math">Mathematical Operators</option>
            <option value="arrows">Arrows</option>
            <option value="latin1">Latin-1 Supplement</option>
          </select>
        </div>
      </div>

      <div class="sym-grid" ${ref("gridEl")}></div>

      <div class="recents-title">${(x) => t("symbol.recentlyUsed", x)}:</div>
      <div class="recents-grid" ${ref("recentsEl")}></div>

      <div class="sym-preview">
        <span class="glyph" ${ref("glyphEl")}></span>
        <span class="codepoint" ${ref("codepointEl")}></span>
      </div>
    </div>
    <div slot="action">
      <fluent-button
        appearance="accent"
        ${ref("insertBtn")}
        @click="${(x) => x.insertSymbol()}"
      ></fluent-button>
      <fluent-button ${ref("closeBtn")} @click="${(x) => x.hide()}"></fluent-button>
    </div>
  </docen-dialog>
`;

@customElement({ name: "docen-symbol-dialog", template, styles })
export class DocenSymbolDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable gridEl?: HTMLElement;
  @observable recentsEl?: HTMLElement;
  @observable glyphEl?: HTMLElement;
  @observable codepointEl?: HTMLElement;
  @observable fontSelect?: HTMLSelectElement;
  @observable subsetSelect?: HTMLSelectElement;
  @observable subsetGroup?: HTMLElement;
  @observable insertBtn?: HTMLElement;
  @observable closeBtn?: HTMLElement;

  #selected = "";
  #currentFont = "normal";
  #currentSubset = "common";
  #recents: string[] = ["©", "®", "™", "€", "£", "¥", "✓", "★"];
  #unobserveLang?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#renderGrid();
    this.#renderRecents();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  override disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(): void {
    if (!this.#selected) {
      const glyphs = this.#getGlyphs();
      if (glyphs[0]) this.#select(glyphs[0]);
    }
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  onFontChange(): void {
    this.#currentFont = this.fontSelect?.value || "normal";
    if (this.subsetGroup) {
      this.subsetGroup.style.display = this.#currentFont === "normal" ? "flex" : "none";
    }
    this.#renderGrid();
  }

  onSubsetChange(): void {
    this.#currentSubset = this.subsetSelect?.value || "common";
    this.#renderGrid();
  }

  #getGlyphs(): string[] {
    if (this.#currentFont === "Symbol") return SYMBOL_FONT_GLYPHS;
    if (this.#currentFont === "Wingdings") return WINGDINGS_GLYPHS;
    if (this.#currentFont === "Webdings") return WEBDINGS_GLYPHS;
    return NORMAL_SUBSETS[this.#currentSubset] ?? NORMAL_SUBSETS.common;
  }

  #select(char: string): void {
    this.#selected = char;
    if (this.glyphEl) this.glyphEl.textContent = char;
    if (this.codepointEl) {
      const code = char.codePointAt(0)?.toString(16).toUpperCase().padStart(4, "0") ?? "";
      this.codepointEl.textContent = `Unicode: U+${code} (${char})`;
    }
    this.gridEl?.querySelectorAll("button").forEach((b) => {
      b.classList.toggle("selected", b.textContent === char);
    });
  }

  #insert(): void {
    if (!this.#selected) return;
    // Add to recents
    this.#recents = [this.#selected, ...this.#recents.filter((c) => c !== this.#selected)].slice(
      0,
      16,
    );
    this.#renderRecents();

    this.$emit("symbol:insert", { char: this.#selected });
  }

  insertSymbol(): void {
    this.#insert();
  }

  #renderGrid(): void {
    if (!this.gridEl) return;
    this.gridEl.replaceChildren();
    const glyphs = this.#getGlyphs();
    for (const char of glyphs) {
      const cell = document.createElement("button");
      cell.type = "button";
      cell.textContent = char;
      cell.addEventListener("click", () => this.#select(char));
      cell.addEventListener("dblclick", () => {
        this.#select(char);
        this.#insert();
      });
      this.gridEl.append(cell);
    }
    if (glyphs[0]) this.#select(glyphs[0]);
  }

  #renderRecents(): void {
    if (!this.recentsEl) return;
    this.recentsEl.replaceChildren();
    for (const char of this.#recents) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = char;
      btn.addEventListener("click", () => {
        this.#select(char);
      });
      btn.addEventListener("dblclick", () => {
        this.#select(char);
        this.#insert();
      });
      this.recentsEl.append(btn);
    }
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("symbol.title", this);
    if (this.insertBtn) this.insertBtn.textContent = t("symbol.insert", this);
    if (this.closeBtn) this.closeBtn.textContent = t("symbol.close", this);
  }
}

export default DocenSymbolDialog;
