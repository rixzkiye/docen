import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface SignatureLineData {
  name: string;
  title: string;
  email: string;
  instructions: string;
  showDate: boolean;
  allowComments: boolean;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(440px, 94vw);
  }
  .sig-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .field-row {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .field-row label {
    font-weight: 500;
    color: var(--docen-color-foreground, #242424);
  }
  .field-row fluent-text-input {
    width: 100%;
  }
  .field-row textarea {
    width: 100%;
    box-sizing: border-box;
    font-family: inherit;
    font-size: 12px;
    padding: 6px;
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 4px;
    resize: vertical;
  }
  .checkboxes {
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-top: 4px;
  }
  .footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 12px;
  }
`;

const template = html<DocenSignatureLineDialog>`
  <docen-dialog ${ref("dialog")} modal heading="${(x) => t("sig.dialogTitle", x)}">
    <div class="sig-body">
      <div class="field-row">
        <label>${(x) => t("sig.suggestedSigner", x)}:</label>
        <fluent-text-input
          ${ref("nameInput")}
          placeholder="${(x) => t("sig.placeholderName", x)}"
        ></fluent-text-input>
      </div>

      <div class="field-row">
        <label>${(x) => t("sig.suggestedTitle", x)}:</label>
        <fluent-text-input
          ${ref("titleInput")}
          placeholder="${(x) => t("sig.placeholderTitle", x)}"
        ></fluent-text-input>
      </div>

      <div class="field-row">
        <label>${(x) => t("sig.suggestedEmail", x)}:</label>
        <fluent-text-input
          ${ref("emailInput")}
          placeholder="${(x) => t("sig.placeholderEmail", x)}"
        ></fluent-text-input>
      </div>

      <div class="field-row">
        <label>${(x) => t("sig.instructions", x)}:</label>
        <textarea ${ref("instructionsInput")} rows="2">
${(x) => t("sig.defaultInstructions", x)}</textarea>
      </div>

      <div class="checkboxes">
        <fluent-checkbox ${ref("allowCommentsCheckbox")}>
          ${(x) => t("sig.allowComments", x)}
        </fluent-checkbox>
        <fluent-checkbox ${ref("showDateCheckbox")} checked>
          ${(x) => t("sig.showDate", x)}
        </fluent-checkbox>
      </div>
    </div>
    <div class="footer" slot="action">
      <fluent-button appearance="primary" @click="${(x) => x.onOk()}"
        >${(x) => t("dialog.ok", x)}</fluent-button
      >
      <fluent-button @click="${(x) => x.close()}">${(x) => t("dialog.cancel", x)}</fluent-button>
    </div>
  </docen-dialog>
`;

@customElement({ name: "docen-signature-line-dialog", template, styles })
export class DocenSignatureLineDialog extends FASTElement {
  @observable dialog?: HTMLElement & { open: boolean };
  @observable nameInput?: HTMLElement & { value: string };
  @observable titleInput?: HTMLElement & { value: string };
  @observable emailInput?: HTMLElement & { value: string };
  @observable instructionsInput?: HTMLTextAreaElement;
  @observable allowCommentsCheckbox?: HTMLElement & { checked: boolean };
  @observable showDateCheckbox?: HTMLElement & { checked: boolean };

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

  onOk(): void {
    const data: SignatureLineData = {
      name: this.nameInput?.value ?? "",
      title: this.titleInput?.value ?? "",
      email: this.emailInput?.value ?? "",
      instructions: this.instructionsInput?.value ?? "",
      allowComments: Boolean(this.allowCommentsCheckbox?.checked),
      showDate: Boolean(this.showDateCheckbox?.checked),
    };

    this.$emit("signature-line:insert", data);
    this.close();
  }
}
