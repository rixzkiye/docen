import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import {
  TranslationEngine,
  TranslationPackManager,
  translationEngine,
  translationPackManager,
} from "../../../document/translation";
import type { LanguageInfo, TranslationPackInfo } from "../../../document/translation/types";
import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: 0;
    box-sizing: border-box;
    font-size: 12px;
    font-family: inherit;
    color: var(--docen-color-text-1, #242424);
  }

  [hidden] {
    display: none !important;
  }

  /* Word-style top tab strip */
  .tab-bar {
    display: flex;
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    background: var(--docen-color-subtle-background, #f9f9f9);
  }

  .tab-btn {
    flex: 1;
    padding: 8px 12px;
    border: none;
    background: transparent;
    font-size: 12px;
    font-weight: 500;
    color: var(--docen-color-text-2, #616161);
    cursor: pointer;
    border-bottom: 2px solid transparent;
    outline: none;
    transition: all 0.15s ease-in-out;
  }

  .tab-btn:hover {
    color: var(--docen-color-text-1, #242424);
    background: rgba(0, 0, 0, 0.03);
  }

  .tab-btn.active {
    color: var(--docen-color-accent, #0f6cbd);
    font-weight: 600;
    border-bottom-color: var(--docen-color-accent, #0f6cbd);
    background: #ffffff;
  }

  .pane-body {
    flex: 1;
    min-height: 0;
    overflow-y: auto;
    padding: 12px;
    display: flex;
    flex-direction: column;
    gap: 12px;
  }

  .tab-panel {
    display: flex;
    flex-direction: column;
    gap: 12px;
  }

  /* Language Selector Row */
  .lang-controls {
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .lang-select-group {
    flex: 1;
    display: flex;
    flex-direction: column;
    gap: 4px;
    min-width: 0;
  }

  .lang-label {
    font-size: 11px;
    font-weight: 600;
    color: var(--docen-color-text-2, #616161);
  }

  .lang-select {
    height: 28px;
    padding: 2px 6px;
    border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 4px;
    background: #ffffff;
    font-size: 12px;
    font-family: inherit;
    color: var(--docen-color-text-1, #242424);
    outline: none;
    cursor: pointer;
  }

  .lang-select:focus {
    border-color: var(--docen-color-accent, #0f6cbd);
  }

  .swap-btn {
    align-self: flex-end;
    margin-bottom: 2px;
    min-width: 28px;
    height: 28px;
    padding: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
  }

  /* Text Areas */
  .section-label {
    font-size: 11px;
    font-weight: 600;
    color: var(--docen-color-text-2, #616161);
    margin-bottom: 4px;
  }

  .source-textarea {
    width: 100%;
    height: 90px;
    box-sizing: border-box;
    padding: 8px;
    border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 4px;
    resize: vertical;
    font-family: inherit;
    font-size: 12px;
    line-height: 1.4;
    outline: none;
  }

  .source-textarea:focus {
    border-color: var(--docen-color-accent, #0f6cbd);
  }

  .target-card {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .target-box {
    min-height: 90px;
    max-height: 180px;
    overflow-y: auto;
    padding: 8px;
    background: var(--docen-color-subtle-background, #fdfdfd);
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 4px;
    font-size: 12px;
    line-height: 1.4;
    white-space: pre-wrap;
    word-break: break-word;
    box-sizing: border-box;
  }

  .target-empty {
    color: var(--docen-color-text-2, #888888);
    font-style: italic;
  }

  .btn-row {
    display: flex;
    gap: 8px;
    margin-top: 4px;
  }

  .action-btn {
    flex: 1;
  }

  /* Document Tab Styling */
  .doc-info {
    font-size: 12px;
    line-height: 1.5;
    color: var(--docen-color-text-2, #616161);
    background: var(--docen-color-subtle-background, #f5f5f5);
    padding: 10px;
    border-radius: 4px;
    border-left: 3px solid var(--docen-color-accent, #0f6cbd);
  }

  .doc-status {
    padding: 8px 10px;
    border-radius: 4px;
    font-size: 12px;
    background: #e6f4ea;
    color: #137333;
    border: 1px solid #ceead6;
  }

  /* Language pack status bar & download simulator */
  .pack-status-footer {
    border-top: 1px solid var(--docen-color-divider, #e2e2e2);
    padding: 8px 12px;
    background: var(--docen-color-subtle-background, #f9f9f9);
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .pack-status-line {
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 11px;
    color: var(--docen-color-text-2, #616161);
  }

  .pack-status-summary {
    display: flex;
    align-items: center;
    gap: 5px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .pack-drawer {
    border: 1px solid var(--docen-color-divider, #e2e2e2);
    border-radius: 4px;
    background: #ffffff;
    padding: 8px;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .pack-drawer-header {
    font-weight: 600;
    font-size: 11px;
    color: var(--docen-color-text-1, #242424);
    margin-bottom: 2px;
  }

  .pack-list {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .pack-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 4px 6px;
    background: var(--docen-color-subtle-background, #fbfbfb);
    border: 1px solid var(--docen-color-divider, #eeeeee);
    border-radius: 4px;
    font-size: 11px;
  }

  .pack-badge {
    font-size: 10px;
    font-weight: 600;
    padding: 2px 6px;
    border-radius: 3px;
  }

  .pack-badge.installed {
    background: #e6f4ea;
    color: #137333;
  }
`;

const template = html<DocenTranslatePane>`
  <!-- Word Two-Tab Header -->
  <div class="tab-bar" part="tab-bar">
    <button class="tab-btn active" part="tab-selection">
      ${(x) => t("translate.selection", x) || "Selection"}
    </button>
    <button class="tab-btn" part="tab-document">
      ${(x) => t("translate.document", x) || "Document"}
    </button>
  </div>

  <div class="pane-body" part="body">
    <!-- TAB 1: SELECTION -->
    <div class="tab-panel" part="panel-selection">
      <!-- Language selectors -->
      <div class="lang-controls">
        <div class="lang-select-group">
          <span class="lang-label">${(x) => t("translate.from", x) || "From"}</span>
          <select class="lang-select" part="source-lang-select" :value="${(x) => x.sourceLang}">
            <option value="auto">${(x) => t("translate.autoDetect", x) || "Auto-detect"}</option>
            ${repeat(
              (x) => x.supportedLanguages,
              html<LanguageInfo>`<option :value="${(l) => l.code}">${(l) => l.name}</option>`,
            )}
          </select>
        </div>

        <fluent-button appearance="subtle" class="swap-btn" part="swap-btn" title="Swap languages">
          ⇄
        </fluent-button>

        <div class="lang-select-group">
          <span class="lang-label">${(x) => t("translate.to", x) || "To"}</span>
          <select class="lang-select" part="target-lang-select" :value="${(x) => x.targetLang}">
            ${repeat(
              (x) => x.supportedLanguages,
              html<LanguageInfo>`<option :value="${(l) => l.code}">${(l) => l.name}</option>`,
            )}
          </select>
        </div>
      </div>

      <!-- Source Text Box -->
      <div>
        <div class="section-label">${(x) => t("translate.sourceText", x) || "Source text"}</div>
        <textarea
          class="source-textarea"
          part="source-text"
          ${ref("sourceInputEl")}
          placeholder="${(x) =>
            t("translate.placeholder", x) || "Enter text or select in document..."}"
          :value="${(x) => x.sourceText}"
        ></textarea>
      </div>

      <!-- Target Translated Text Box -->
      <div class="target-card">
        <div class="section-label">${(x) => t("translate.translatedText", x) || "Translation"}</div>
        <div class="target-box" part="target-text">
          ${(x) =>
            x.translatedText
              ? x.translatedText
              : html`<span class="target-empty"
                  >${
                    t("translate.noSelection", x) || "Select text in the document or type above."
                  }</span
                >`}
        </div>
      </div>

      <!-- Selection Action Buttons -->
      <div class="btn-row">
        <fluent-button appearance="accent" class="action-btn" part="insert-btn">
          ${(x) => t("translate.insert", x) || "Insert"}
        </fluent-button>
        <fluent-button appearance="neutral" class="action-btn" part="copy-btn">
          ${(x) => (x.isCopied ? t("translate.copied", x) || "Copied!" : t("translate.copy", x) || "Copy")}
        </fluent-button>
      </div>
    </div>

    <!-- TAB 2: DOCUMENT -->
    <div class="tab-panel" part="panel-document" hidden>
      <div class="doc-info">
        ${(x) =>
          t("translate.docDesc", x) ||
          "Translate paragraph blocks in your document using offline translation packs."}
      </div>

      <div class="lang-controls">
        <div class="lang-select-group">
          <span class="lang-label">${(x) => t("translate.from", x) || "From"}</span>
          <select
            class="lang-select"
            part="doc-source-lang-select"
            :value="${(x) => x.docSourceLang}"
          >
            <option value="auto">${(x) => t("translate.autoDetect", x) || "Auto-detect"}</option>
            ${repeat(
              (x) => x.supportedLanguages,
              html<LanguageInfo>`<option :value="${(l) => l.code}">${(l) => l.name}</option>`,
            )}
          </select>
        </div>

        <div class="lang-select-group">
          <span class="lang-label">${(x) => t("translate.to", x) || "To"}</span>
          <select
            class="lang-select"
            part="doc-target-lang-select"
            :value="${(x) => x.docTargetLang}"
          >
            ${repeat(
              (x) => x.supportedLanguages,
              html<LanguageInfo>`<option :value="${(l) => l.code}">${(l) => l.name}</option>`,
            )}
          </select>
        </div>
      </div>

      <fluent-button appearance="accent" class="action-btn" part="translate-doc-btn">
        ${(x) =>
          x.isTranslatingDoc
            ? t("translate.translatingDoc", x) || "Translating Document..."
            : t("translate.translateDoc", x) || "Translate Document"}
      </fluent-button>

      <div class="doc-status" part="doc-status" hidden>${(x) => x.docStatusMessage}</div>
    </div>
  </div>

  <!-- Language Pack Status Bar & On-Demand Download Simulator -->
  <div class="pack-status-footer" part="pack-footer">
    <div class="pack-status-line">
      <div class="pack-status-summary">
        <span>🌐</span>
        <span part="pack-status-text">${(x) => x.packStatusText}</span>
      </div>
      <fluent-button appearance="subtle" class="download-toggle-btn" part="download-toggle-btn">
        ${(x) =>
          x.showPacksModal
            ? t("translate.closePacks", x) || "Hide"
            : t("translate.downloadPacks", x) || "Download packs"}
      </fluent-button>
    </div>

    <div class="pack-drawer" part="pack-drawer" hidden>
      <div class="pack-drawer-header">
        ${(x) => t("translate.availablePacks", x) || "Available Offline Packs"}
      </div>
      <div class="pack-list">
        ${repeat(
          (x) => x.availablePacks,
          html<TranslationPackInfo>`
            <div class="pack-row" part="pack-row">
              <div>
                <strong>${(p) => p.name}</strong>
                <span style="opacity: 0.7; margin-left: 4px;">(${(p) => p.size})</span>
              </div>
              ${(p, c) =>
                p.installed
                  ? html`<span class="pack-badge installed"
                      >${t("translate.installed", c.parent) || "Installed"}</span
                    >`
                  : html`
                      <fluent-button
                        appearance="neutral"
                        size="small"
                        class="download-action-btn"
                        @click="${() => c.parent.downloadPack(p.id)}"
                      >
                        ${t("translate.download", c.parent) || "Download"}
                      </fluent-button>
                    `}
            </div>
          `,
        )}
      </div>
    </div>
  </div>
`;

@customElement({ name: "docen-translate-pane", template, styles })
export class DocenTranslatePane extends FASTElement {
  @attr tab: "selection" | "document" = "selection";
  @attr sourceLang = "auto";
  @attr targetLang = "id";
  @attr docSourceLang = "auto";
  @attr docTargetLang = "id";

  @observable sourceText = "";
  @observable translatedText = "";
  @observable isTranslatingDoc = false;
  @observable docStatusMessage = "";
  @observable showPacksModal = false;
  @observable packStatusText = "Packs: English ↔ Indonesian installed";
  @observable availablePacks: TranslationPackInfo[] = [];
  @observable supportedLanguages: LanguageInfo[] = [];
  @observable isCopied = false;

  @observable sourceInputEl?: HTMLTextAreaElement;

  readonly engine: TranslationEngine = translationEngine;
  readonly packManager: TranslationPackManager = translationPackManager;

  #unobserveLang?: () => void;

  tabChanged(): void {
    const root = this.shadowRoot;
    if (!root) return;
    const selTab = root.querySelector("[part='tab-selection']");
    const docTab = root.querySelector("[part='tab-document']");
    const selPanel = root.querySelector("[part='panel-selection']");
    const docPanel = root.querySelector("[part='panel-document']");

    if (this.tab === "selection") {
      selTab?.classList.add("active");
      docTab?.classList.remove("active");
      selPanel?.removeAttribute("hidden");
      docPanel?.setAttribute("hidden", "");
    } else {
      selTab?.classList.remove("active");
      docTab?.classList.add("active");
      selPanel?.setAttribute("hidden", "");
      docPanel?.removeAttribute("hidden");
    }
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.refreshPacks();
    this.performTranslation();

    const root = this.shadowRoot;
    if (root) {
      root
        .querySelector("[part='tab-selection']")
        ?.addEventListener("click", () => this.selectTab("selection"));
      root
        .querySelector("[part='tab-document']")
        ?.addEventListener("click", () => this.selectTab("document"));
      root
        .querySelector("[part='insert-btn']")
        ?.addEventListener("click", () => this.handleInsert());
      root
        .querySelector("[part='copy-btn']")
        ?.addEventListener("click", () => void this.handleCopy());
      root
        .querySelector("[part='swap-btn']")
        ?.addEventListener("click", () => this.swapLanguages());
      root
        .querySelector("[part='translate-doc-btn']")
        ?.addEventListener("click", () => this.handleTranslateDocument());
      root
        .querySelector("[part='download-toggle-btn']")
        ?.addEventListener("click", () => this.togglePacksModal());

      const srcSelect = root.querySelector<HTMLSelectElement>("[part='source-lang-select']");
      srcSelect?.addEventListener("change", (e) => this.onSourceLangChange(e));

      const tgtSelect = root.querySelector<HTMLSelectElement>("[part='target-lang-select']");
      tgtSelect?.addEventListener("change", (e) => this.onTargetLangChange(e));

      const docSrcSelect = root.querySelector<HTMLSelectElement>("[part='doc-source-lang-select']");
      docSrcSelect?.addEventListener("change", (e) => this.onDocSourceChange(e));

      const docTgtSelect = root.querySelector<HTMLSelectElement>("[part='doc-target-lang-select']");
      docTgtSelect?.addEventListener("change", (e) => this.onDocTargetChange(e));

      const srcTextarea = root.querySelector<HTMLTextAreaElement>("[part='source-text']");
      srcTextarea?.addEventListener("input", (e) => this.onSourceInput(e));
    }

    this.#unobserveLang = observeLang(() => {
      this.refreshPacks();
      this.performTranslation();
    });
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this.#unobserveLang?.();
  }

  selectTab(tab: "selection" | "document"): void {
    this.tab = tab;
    this.tabChanged();
  }

  setSelectionText(text: string): void {
    this.sourceText = text;
    if (this.sourceInputEl) {
      this.sourceInputEl.value = text;
    }
    const root = this.shadowRoot;
    if (root) {
      const srcTextarea = root.querySelector<HTMLTextAreaElement>("[part='source-text']");
      if (srcTextarea) srcTextarea.value = text;
    }
    this.performTranslation();
  }

  onSourceInput(event: Event): void {
    this.sourceText = (event.target as HTMLTextAreaElement).value;
    this.performTranslation();
  }

  onSourceLangChange(event: Event): void {
    this.sourceLang = (event.target as HTMLSelectElement).value;
    this.performTranslation();
  }

  onTargetLangChange(event: Event): void {
    this.targetLang = (event.target as HTMLSelectElement).value;
    this.performTranslation();
  }

  onDocSourceChange(event: Event): void {
    this.docSourceLang = (event.target as HTMLSelectElement).value;
  }

  onDocTargetChange(event: Event): void {
    this.docTargetLang = (event.target as HTMLSelectElement).value;
  }

  swapLanguages(): void {
    const newSource = this.targetLang;
    const newTarget = this.sourceLang === "auto" ? "en" : this.sourceLang;

    this.sourceLang = newSource;
    this.targetLang = newTarget;

    if (this.translatedText) {
      this.sourceText = this.translatedText;
      if (this.sourceInputEl) {
        this.sourceInputEl.value = this.sourceText;
      }
      const root = this.shadowRoot;
      if (root) {
        const srcTextarea = root.querySelector<HTMLTextAreaElement>("[part='source-text']");
        if (srcTextarea) srcTextarea.value = this.sourceText;
      }
    }
    this.performTranslation();
  }

  performTranslation(): void {
    if (!this.sourceText.trim()) {
      this.translatedText = "";
      return;
    }

    const res = this.engine.translate(this.sourceText, {
      from: this.sourceLang,
      to: this.targetLang,
    });
    this.translatedText = res.translatedText;
  }

  handleInsert(): void {
    if (!this.translatedText) return;
    this.dispatchEvent(
      new CustomEvent("translate:insert", {
        bubbles: true,
        composed: true,
        detail: this.translatedText,
      }),
    );
    this.$emit("translate:insert", this.translatedText);
  }

  async handleCopy(): Promise<void> {
    if (!this.translatedText) return;
    try {
      await navigator.clipboard?.writeText(this.translatedText);
      this.isCopied = true;
      setTimeout(() => {
        this.isCopied = false;
      }, 2000);
    } catch {
      // Ignore clipboard error in test env
    }
  }

  handleTranslateDocument(): void {
    this.isTranslatingDoc = true;
    this.docStatusMessage = "";
    this.dispatchEvent(
      new CustomEvent("translate:document", {
        bubbles: true,
        composed: true,
        detail: { from: this.docSourceLang, to: this.docTargetLang },
      }),
    );
    this.$emit("translate:document", { from: this.docSourceLang, to: this.docTargetLang });
  }

  onDocumentTranslated(count: number): void {
    this.isTranslatingDoc = false;
    this.docStatusMessage = `Document translated (${count} paragraph${count === 1 ? "" : "s"}).`;
    const statusEl = this.shadowRoot?.querySelector("[part='doc-status']");
    if (statusEl) {
      statusEl.removeAttribute("hidden");
      statusEl.textContent = this.docStatusMessage;
    }
  }

  togglePacksModal(): void {
    this.showPacksModal = !this.showPacksModal;
    const drawer = this.shadowRoot?.querySelector("[part='pack-drawer']");
    if (this.showPacksModal) {
      drawer?.removeAttribute("hidden");
    } else {
      drawer?.setAttribute("hidden", "");
    }
    this.refreshPacks();
  }

  async downloadPack(packId: string): Promise<void> {
    await this.packManager.downloadPack(packId);
    this.refreshPacks();
    this.performTranslation();
  }

  refreshPacks(): void {
    this.availablePacks = this.packManager.listAvailablePacks();
    this.supportedLanguages = this.packManager.getSupportedLanguages();

    const installedCount = this.availablePacks.filter((p) => p.installed).length;
    this.packStatusText = `Packs: ${installedCount} installed (en, id${
      installedCount > 1 ? ", +" + (installedCount - 1) : ""
    })`;
  }
}

export default DocenTranslatePane;
