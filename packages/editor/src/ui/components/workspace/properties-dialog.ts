import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";
import type { WordCountStats } from "./word-count-dialog";

export interface DocumentPropertiesCore {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
}

export interface DocumentPropertiesStats extends Partial<WordCountStats> {
  revision?: number;
}

type FluentTextInput = HTMLElement & { value: string };

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(440px, 92vw);
  }
  .prop-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .tabs-bar {
    display: flex;
    border-bottom: 1px solid var(--colorNeutralStroke2, #d1d1d1);
    gap: 12px;
  }
  .tab-btn {
    padding: 6px 12px;
    cursor: pointer;
    border: none;
    background: none;
    font-size: 13px;
    font-weight: 500;
    color: var(--colorNeutralForeground2, #555);
    border-bottom: 2px solid transparent;
  }
  .tab-btn:hover {
    color: var(--colorNeutralForeground1, #111);
  }
  .tab-btn.active {
    color: var(--colorBrandForeground1, #0078d4);
    border-bottom-color: var(--colorBrandStroke1, #0078d4);
    font-weight: 600;
  }
  .panel {
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .field > label {
    width: 90px;
    flex-shrink: 0;
    font-size: 13px;
    font-weight: 500;
  }
  fluent-text-input {
    flex: 1 1 auto;
    min-width: 0;
  }
  .stats-grid {
    display: flex;
    flex-direction: column;
  }
  .stat-row {
    display: flex;
    justify-content: space-between;
    padding: 6px 2px;
    border-bottom: 1px solid var(--colorNeutralStroke2, #e8e8e8);
  }
  .stat-row:last-child {
    border-bottom: none;
  }
  .stat-label {
    font-size: 13px;
    color: var(--colorNeutralForeground2, #444);
  }
  .stat-value {
    font-size: 13px;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
  }
`;

const template = html<DocenPropertiesDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="prop-body">
      <div class="tabs-bar">
        <button
          class="tab-btn"
          ${ref("tabSummaryBtn")}
          @click="${(x) => x.setTab("summary")}"
        ></button>
        <button class="tab-btn" ${ref("tabStatsBtn")} @click="${(x) => x.setTab("stats")}"></button>
      </div>

      <div class="panel" ?hidden="${(x) => x.activeTab !== "summary"}">
        <div class="field">
          <label ${ref("titleLabel")}></label>
          <fluent-text-input ${ref("titleInput")} type="text"></fluent-text-input>
        </div>
        <div class="field">
          <label ${ref("subjectLabel")}></label>
          <fluent-text-input ${ref("subjectInput")} type="text"></fluent-text-input>
        </div>
        <div class="field">
          <label ${ref("authorLabel")}></label>
          <fluent-text-input ${ref("authorInput")} type="text"></fluent-text-input>
        </div>
        <div class="field">
          <label ${ref("keywordsLabel")}></label>
          <fluent-text-input ${ref("keywordsInput")} type="text"></fluent-text-input>
        </div>
        <div class="field">
          <label ${ref("commentsLabel")}></label>
          <fluent-text-input ${ref("commentsInput")} type="text"></fluent-text-input>
        </div>
      </div>

      <div class="panel" ?hidden="${(x) => x.activeTab !== "stats"}">
        <div class="stats-grid">
          <div class="stat-row">
            <span class="stat-label" ${ref("statPagesLabel")}></span>
            <span class="stat-value" ${ref("statPagesVal")}>0</span>
          </div>
          <div class="stat-row">
            <span class="stat-label" ${ref("statWordsLabel")}></span>
            <span class="stat-value" ${ref("statWordsVal")}>0</span>
          </div>
          <div class="stat-row">
            <span class="stat-label" ${ref("statCharsLabel")}></span>
            <span class="stat-value" ${ref("statCharsVal")}>0</span>
          </div>
          <div class="stat-row">
            <span class="stat-label" ${ref("statParasLabel")}></span>
            <span class="stat-value" ${ref("statParasVal")}>0</span>
          </div>
          <div class="stat-row">
            <span class="stat-label" ${ref("statLinesLabel")}></span>
            <span class="stat-value" ${ref("statLinesVal")}>0</span>
          </div>
          <div class="stat-row">
            <span class="stat-label" ${ref("statRevLabel")}></span>
            <span class="stat-value" ${ref("statRevVal")}>1</span>
          </div>
        </div>
      </div>
    </div>

    <div slot="action">
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.onOk()}"
      ></fluent-button>
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-properties-dialog>` — Word's Document Properties dialog.
 * View/edit core properties (Title, Subject, Author, Keywords, Comments)
 * and view statistics (Pages, Words, Characters, Paragraphs, Lines, Revision).
 */
@customElement({ name: "docen-properties-dialog", template, styles })
export class DocenPropertiesDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable tabSummaryBtn?: HTMLElement;
  @observable tabStatsBtn?: HTMLElement;

  @observable titleLabel?: HTMLElement;
  @observable titleInput?: FluentTextInput;
  @observable subjectLabel?: HTMLElement;
  @observable subjectInput?: FluentTextInput;
  @observable authorLabel?: HTMLElement;
  @observable authorInput?: FluentTextInput;
  @observable keywordsLabel?: HTMLElement;
  @observable keywordsInput?: FluentTextInput;
  @observable commentsLabel?: HTMLElement;
  @observable commentsInput?: FluentTextInput;

  @observable statPagesLabel?: HTMLElement;
  @observable statPagesVal?: HTMLElement;
  @observable statWordsLabel?: HTMLElement;
  @observable statWordsVal?: HTMLElement;
  @observable statCharsLabel?: HTMLElement;
  @observable statCharsVal?: HTMLElement;
  @observable statParasLabel?: HTMLElement;
  @observable statParasVal?: HTMLElement;
  @observable statLinesLabel?: HTMLElement;
  @observable statLinesVal?: HTMLElement;
  @observable statRevLabel?: HTMLElement;
  @observable statRevVal?: HTMLElement;

  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  @observable activeTab: "summary" | "stats" = "summary";

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#updateTabButtons();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(core: DocumentPropertiesCore = {}, stats: DocumentPropertiesStats = {}): void {
    this.activeTab = "summary";
    this.#updateTabButtons();

    if (this.titleInput) this.titleInput.value = core.title ?? "";
    if (this.subjectInput) this.subjectInput.value = core.subject ?? "";
    if (this.authorInput) this.authorInput.value = core.creator ?? "";
    if (this.keywordsInput) this.keywordsInput.value = core.keywords ?? "";
    if (this.commentsInput) this.commentsInput.value = core.description ?? "";

    if (this.statPagesVal) this.statPagesVal.textContent = String(stats.pages ?? 0);
    if (this.statWordsVal) this.statWordsVal.textContent = String(stats.words ?? 0);
    if (this.statCharsVal)
      this.statCharsVal.textContent = String(stats.charsNoSpaces ?? stats.charsWithSpaces ?? 0);
    if (this.statParasVal) this.statParasVal.textContent = String(stats.paragraphs ?? 0);
    if (this.statLinesVal) this.statLinesVal.textContent = String(stats.lines ?? 0);
    if (this.statRevVal) this.statRevVal.textContent = String(stats.revision ?? 1);

    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  setTab(tab: "summary" | "stats"): void {
    this.activeTab = tab;
    this.#updateTabButtons();
  }

  onOk(): void {
    const patch: DocumentPropertiesCore = {
      title: this.titleInput?.value ?? "",
      subject: this.subjectInput?.value ?? "",
      creator: this.authorInput?.value ?? "",
      keywords: this.keywordsInput?.value ?? "",
      description: this.commentsInput?.value ?? "",
    };
    this.$emit("properties:ok", { core: patch });
    this.hide();
  }

  #updateTabButtons(): void {
    this.tabSummaryBtn?.classList.toggle("active", this.activeTab === "summary");
    this.tabStatsBtn?.classList.toggle("active", this.activeTab === "stats");
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("properties.title", this);
    if (this.tabSummaryBtn) this.tabSummaryBtn.textContent = t("properties.tab-summary", this);
    if (this.tabStatsBtn) this.tabStatsBtn.textContent = t("properties.tab-stats", this);
    if (this.titleLabel) this.titleLabel.textContent = t("properties.prop-title", this);
    if (this.subjectLabel) this.subjectLabel.textContent = t("properties.prop-subject", this);
    if (this.authorLabel) this.authorLabel.textContent = t("properties.prop-author", this);
    if (this.keywordsLabel) this.keywordsLabel.textContent = t("properties.prop-keywords", this);
    if (this.commentsLabel) this.commentsLabel.textContent = t("properties.prop-comments", this);

    if (this.statPagesLabel) this.statPagesLabel.textContent = t("properties.stat-pages", this);
    if (this.statWordsLabel) this.statWordsLabel.textContent = t("properties.stat-words", this);
    if (this.statCharsLabel) this.statCharsLabel.textContent = t("properties.stat-chars", this);
    if (this.statParasLabel)
      this.statParasLabel.textContent = t("properties.stat-paragraphs", this);
    if (this.statLinesLabel) this.statLinesLabel.textContent = t("properties.stat-lines", this);
    if (this.statRevLabel) this.statRevLabel.textContent = t("properties.stat-revision", this);

    if (this.okBtn) this.okBtn.textContent = t("properties.btn-ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("properties.btn-cancel", this);
  }
}

export default DocenPropertiesDialog;
