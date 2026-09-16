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

import { lookupThesaurus } from "../../../document/proofing/thesaurus";
import type { ThesaurusEntry, ThesaurusMeaning } from "../../../document/proofing/types";
import { observeLang, t } from "../../i18n/localize";

const synonymTemplate = html<string>`
  <fluent-button
    appearance="neutral"
    class="synonym-btn"
    @click="${(s: string, c) => c.parent.emitInsert(s, c.event as Event)}"
  >
    <span>${(s: string) => s}</span>
    <span
      class="action-hint"
      title="Look up"
      @click="${(s: string, c) => c.parent.emitLookup(s, c.event as Event)}"
      >🔍</span
    >
  </fluent-button>
`;

const meaningTemplate = html<ThesaurusMeaning>`
  <div class="meaning-block">
    <div class="pos-line">
      <span class="pos-badge">${(m: ThesaurusMeaning) => m.partOfSpeech}</span>
      ${(m: ThesaurusMeaning) =>
        m.definition ? html`<span class="definition">${m.definition}</span>` : ""}
    </div>
    <div class="synonym-list">${repeat((m: ThesaurusMeaning) => m.synonyms, synonymTemplate)}</div>
    ${(m: ThesaurusMeaning) =>
      m.antonyms && m.antonyms.length
        ? html`
            <div class="antonyms-section">
              <span class="antonyms-label">Antonyms:</span>
              <span class="antonyms-text">${m.antonyms.join(", ")}</span>
            </div>
          `
        : ""}
  </div>
`;

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: 0;
    box-sizing: border-box;
    font-size: 12px;
  }
  .search-bar {
    display: flex;
    gap: 6px;
    padding: 10px;
    border-block-end: 1px solid var(--docen-color-divider, #e2e2e2);
  }
  .search-input {
    flex: 1;
    min-width: 0;
    height: 28px;
    padding: 0 8px;
    box-sizing: border-box;
    border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 4px;
    font-size: 12px;
    font-family: inherit;
    outline: none;
  }
  .search-input:focus {
    border-color: var(--docen-color-accent, #0f6cbd);
  }
  .body {
    flex: 1;
    min-height: 0;
    overflow: auto;
    padding: 10px;
    display: flex;
    flex-direction: column;
    gap: 12px;
  }
  .empty,
  .not-found {
    color: var(--docen-color-text-2, #616161);
    padding: 18px 8px;
    text-align: center;
    line-height: 1.5;
  }
  .word-header {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 8px;
  }
  .word-title {
    font-size: 18px;
    font-weight: 600;
    color: var(--docen-color-text-1, #242424);
  }
  .meaning-block {
    display: flex;
    flex-direction: column;
    gap: 6px;
    padding: 8px;
    background: var(--docen-color-subtle-background, #fbfbfb);
    border: 1px solid var(--docen-color-divider, #e2e2e2);
    border-radius: 6px;
  }
  .pos-line {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .pos-badge {
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
    background: #e1dfdd;
    color: #323130;
    padding: 2px 6px;
    border-radius: 3px;
  }
  .definition {
    color: var(--docen-color-text-2, #616161);
    font-style: italic;
  }
  .synonym-list {
    display: flex;
    flex-direction: column;
    gap: 3px;
    margin-top: 4px;
  }
  .synonym-btn {
    display: flex;
    justify-content: space-between;
    align-items: center;
    width: 100%;
    text-align: start;
    padding: 4px 8px;
  }
  .action-hint {
    font-size: 10px;
    opacity: 0.6;
    padding: 2px 4px;
    border-radius: 3px;
    cursor: pointer;
  }
  .action-hint:hover {
    opacity: 1;
    background: rgba(0, 0, 0, 0.08);
  }
  .antonyms-section {
    font-size: 11px;
    color: var(--docen-color-text-2, #616161);
    margin-top: 4px;
  }
  .antonyms-label {
    font-weight: 600;
    margin-right: 4px;
  }
`;

const template = html<DocenThesaurusPane>`
  <div class="search-bar">
    <input
      type="text"
      class="search-input"
      ${ref("inputEl")}
      placeholder="${(x) => t("thesaurus.search", x)}"
      :value="${(x) => x.query}"
      @keydown="${(x, c) => x.onInputKeydown(c.event as KeyboardEvent)}"
    />
    <fluent-button appearance="accent" @click="${(x) => x.triggerLookup()}"
      >${(x) => t("thesaurus.lookup", x)}</fluent-button
    >
  </div>
  <div class="body">
    ${(x) =>
      x.entry
        ? html`
            <div class="word-header">
              <span class="word-title">${(x) => x.entry?.word}</span>
            </div>
            ${repeat((x) => x.entry?.meanings ?? [], meaningTemplate)}
          `
        : x.notFound
          ? html`<div class="not-found">${(x) => t("thesaurus.no-results", x)} "${x.query}"</div>`
          : html`<div class="empty">${(x) => t("thesaurus.empty", x)}</div>`}
  </div>
`;

@customElement({ name: "docen-thesaurus-pane", template, styles })
class DocenThesaurusPane extends FASTElement {
  @attr query = "";
  @attr lang = "en";
  @observable entry: ThesaurusEntry | null = null;
  @observable notFound = false;
  @observable inputEl?: HTMLInputElement;

  onInputKeydown(event: KeyboardEvent): void {
    if (event.key === "Enter") {
      event.preventDefault();
      this.triggerLookup();
    }
  }

  triggerLookup(): void {
    const val = this.inputEl?.value.trim() ?? "";
    if (!val) return;
    this.lookup(val);
  }

  lookup(word: string): void {
    this.query = word;
    if (this.inputEl) this.inputEl.value = word;
    const res = lookupThesaurus(word, this.lang);
    if (res) {
      this.entry = res;
      this.notFound = false;
    } else {
      this.entry = null;
      this.notFound = true;
    }
  }

  emitInsert(word: string, event: Event): void {
    event.stopPropagation();
    this.$emit("thesaurus:insert", word);
  }

  emitLookup(word: string, event: Event): void {
    event.stopPropagation();
    this.lookup(word);
  }

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#unobserveLang = observeLang(() => {
      if (this.entry) this.entry = { ...this.entry };
    });
  }

  disconnectedCallback(): void {
    super.disconnectedCallback();
    this.#unobserveLang?.();
  }
}

export default DocenThesaurusPane;
