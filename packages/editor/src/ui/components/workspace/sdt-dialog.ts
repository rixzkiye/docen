import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface SdtPropertiesValues {
  title?: string;
  tag?: string;
  cannotDelete?: boolean;
  cannotEdit?: boolean;
  color?: string;
  dateFormat?: string;
  checkedSymbol?: string;
  uncheckedSymbol?: string;
  listItems?: Array<{ displayText: string; value: string }>;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(440px, 92vw);
  }
  .sdt-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .section-title {
    font-weight: 600;
    color: var(--docen-color-foreground, #242424);
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    padding-bottom: 4px;
    margin-bottom: 4px;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .label {
    width: 90px;
    flex: 0 0 auto;
  }
  fluent-text-input {
    flex: 1;
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
    margin-top: 8px;
  }
`;

const template = html<DocenSdtDialog>`
  <docen-dialog ${ref("dialog")} modal heading="${(x) => t("sdt.dialogTitle", x)}">
    <div class="sdt-body" slot="body">
      <div class="section-title">${(x) => t("sdt.general", x)}</div>
      <div class="row">
        <label class="label">${(x) => t("sdt.title", x)}:</label>
        <fluent-text-input
          ${ref("titleInput")}
          :value="${(x) => x.values?.title ?? ""}"
        ></fluent-text-input>
      </div>
      <div class="row">
        <label class="label">${(x) => t("sdt.tag", x)}:</label>
        <fluent-text-input
          ${ref("tagInput")}
          :value="${(x) => x.values?.tag ?? ""}"
        ></fluent-text-input>
      </div>

      <div class="section-title">${(x) => t("sdt.locking", x)}</div>
      <div class="checkboxes">
        <fluent-checkbox
          ${ref("lockDeleteCheckbox")}
          ?checked="${(x) => Boolean(x.values?.cannotDelete)}"
        >
          ${(x) => t("sdt.cannotDelete", x)}
        </fluent-checkbox>
        <fluent-checkbox
          ${ref("lockEditCheckbox")}
          ?checked="${(x) => Boolean(x.values?.cannotEdit)}"
        >
          ${(x) => t("sdt.cannotEdit", x)}
        </fluent-checkbox>
      </div>
    </div>
    <div class="footer" slot="footer">
      <fluent-button appearance="primary" @click="${(x) => x.onOk()}"
        >${(x) => t("dialog.ok", x)}</fluent-button
      >
      <fluent-button @click="${(x) => x.close()}">${(x) => t("dialog.cancel", x)}</fluent-button>
    </div>
  </docen-dialog>
`;

@customElement({ name: "docen-sdt-dialog", template, styles })
export class DocenSdtDialog extends FASTElement {
  @observable dialog?: HTMLElement & { open: boolean };
  @observable titleInput?: HTMLElement & { value: string };
  @observable tagInput?: HTMLElement & { value: string };
  @observable lockDeleteCheckbox?: HTMLElement & { checked: boolean };
  @observable lockEditCheckbox?: HTMLElement & { checked: boolean };

  @observable values?: SdtPropertiesValues;

  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => {});
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  show(initial?: SdtPropertiesValues): void {
    this.values = initial ? { ...initial } : {};
    if (this.dialog) this.dialog.open = true;
  }

  close(): void {
    if (this.dialog) this.dialog.open = false;
  }

  onOk(): void {
    const updated: SdtPropertiesValues = {
      ...this.values,
      title: this.titleInput?.value ?? "",
      tag: this.tagInput?.value ?? "",
      cannotDelete: Boolean(this.lockDeleteCheckbox?.checked),
      cannotEdit: Boolean(this.lockEditCheckbox?.checked),
    };
    this.$emit("sdt-dialog:ok", { properties: updated });
    this.close();
  }
}
