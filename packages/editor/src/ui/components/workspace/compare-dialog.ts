import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface CompareDialogSubmitPayload {
  originalFile?: File;
  revisedFile?: File;
  originalAuthor?: string;
  revisedAuthor?: string;
  granularity?: "word" | "character";
  compareFormatting?: boolean;
  compareComments?: boolean;
  compareCase?: boolean;
  compareWhitespace?: boolean;
  compareTables?: boolean;
  compareHeadersFooters?: boolean;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(500px, 94vw);
  }
  .compare-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 14px;
    font-size: 13px;
  }
  .section-title {
    font-weight: 600;
    color: var(--docen-color-foreground, #242424);
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    padding-bottom: 4px;
    margin-bottom: 4px;
  }
  .file-row {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .file-inputs {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .file-inputs input[type="file"] {
    flex: 1;
    font-size: 12px;
  }
  .author-row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .author-label {
    width: 80px;
    font-size: 12px;
    color: var(--docen-color-foreground-secondary, #616161);
  }
  .author-row fluent-text-input {
    flex: 1;
  }
  .checkbox-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
  }
  .granularity-row {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-top: 4px;
  }
  .footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 12px;
  }
`;

const template = html<DocenCompareDialog>`
  <docen-dialog ${ref("dialog")} modal heading="${(x) => t("compare.dialogTitle", x)}">
    <div class="compare-body" slot="body">
      <div class="section-title">${(x) => t("compare.originalDocument", x)}</div>
      <div class="file-row">
        <div class="file-inputs">
          <input type="file" accept=".docx" ${ref("originalFileInput")} />
        </div>
        <div class="author-row">
          <span class="author-label">${(x) => t("compare.labelChangesWith", x)}</span>
          <fluent-text-input ${ref("originalAuthorInput")} value="Original"></fluent-text-input>
        </div>
      </div>

      <div class="section-title">${(x) => t("compare.revisedDocument", x)}</div>
      <div class="file-row">
        <div class="file-inputs">
          <input type="file" accept=".docx" ${ref("revisedFileInput")} />
        </div>
        <div class="author-row">
          <span class="author-label">${(x) => t("compare.labelChangesWith", x)}</span>
          <fluent-text-input ${ref("revisedAuthorInput")} value="Comparison"></fluent-text-input>
        </div>
      </div>

      <div class="section-title">${(x) => t("compare.comparisonSettings", x)}</div>
      <div class="checkbox-grid">
        <fluent-checkbox ${ref("cbFormatting")} checked>
          ${(x) => t("compare.formatting", x)}
        </fluent-checkbox>
        <fluent-checkbox ${ref("cbCase")} checked> ${(x) => t("compare.case", x)} </fluent-checkbox>
        <fluent-checkbox ${ref("cbWhitespace")} checked>
          ${(x) => t("compare.whitespace", x)}
        </fluent-checkbox>
        <fluent-checkbox ${ref("cbTables")} checked>
          ${(x) => t("compare.tables", x)}
        </fluent-checkbox>
        <fluent-checkbox ${ref("cbHeaders")} checked>
          ${(x) => t("compare.headersFooters", x)}
        </fluent-checkbox>
        <fluent-checkbox ${ref("cbComments")} checked>
          ${(x) => t("compare.comments", x)}
        </fluent-checkbox>
      </div>

      <div class="granularity-row">
        <span>${(x) => t("compare.showChangesAt", x)}:</span>
        <label
          ><input type="radio" name="granularity" value="word" checked />
          ${(x) => t("compare.wordLevel", x)}</label
        >
        <label
          ><input type="radio" name="granularity" value="character" />
          ${(x) => t("compare.characterLevel", x)}</label
        >
      </div>
    </div>
    <div class="footer" slot="footer">
      <fluent-button appearance="primary" @click="${(x) => x.onCompare()}"
        >${(x) => t("compare.compareButton", x)}</fluent-button
      >
      <fluent-button @click="${(x) => x.close()}">${(x) => t("dialog.cancel", x)}</fluent-button>
    </div>
  </docen-dialog>
`;

@customElement({ name: "docen-compare-dialog", template, styles })
export class DocenCompareDialog extends FASTElement {
  @observable dialog?: HTMLElement & { open: boolean };
  @observable originalFileInput?: HTMLInputElement;
  @observable revisedFileInput?: HTMLInputElement;
  @observable originalAuthorInput?: HTMLElement & { value: string };
  @observable revisedAuthorInput?: HTMLElement & { value: string };
  @observable cbFormatting?: HTMLElement & { checked: boolean };
  @observable cbCase?: HTMLElement & { checked: boolean };
  @observable cbWhitespace?: HTMLElement & { checked: boolean };
  @observable cbTables?: HTMLElement & { checked: boolean };
  @observable cbHeaders?: HTMLElement & { checked: boolean };
  @observable cbComments?: HTMLElement & { checked: boolean };

  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => {});
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  show(): void {
    if (this.dialog) this.dialog.open = true;
  }

  close(): void {
    if (this.dialog) this.dialog.open = false;
  }

  onCompare(): void {
    const origFile = this.originalFileInput?.files?.[0];
    const revFile = this.revisedFileInput?.files?.[0];

    const payload: CompareDialogSubmitPayload = {
      originalFile: origFile,
      revisedFile: revFile,
      originalAuthor: this.originalAuthorInput?.value || "Original",
      revisedAuthor: this.revisedAuthorInput?.value || "Comparison",
      granularity: "word",
      compareFormatting: Boolean(this.cbFormatting?.checked),
      compareCase: Boolean(this.cbCase?.checked),
      compareWhitespace: Boolean(this.cbWhitespace?.checked),
      compareTables: Boolean(this.cbTables?.checked),
      compareHeadersFooters: Boolean(this.cbHeaders?.checked),
      compareComments: Boolean(this.cbComments?.checked),
    };

    this.$emit("compare:execute", payload);
    this.close();
  }
}
