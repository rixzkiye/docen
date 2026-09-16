import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface HyphenationDialogOptions {
  auto: boolean;
  doNotHyphenateCaps: boolean;
  zoneTw?: number;
  limit?: number;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(380px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 8px;
    cursor: pointer;
  }
  .row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
  }
  .row > label {
    white-space: nowrap;
    flex: 1 1 auto;
  }
  .row > input {
    width: 90px;
    padding: 4px 8px;
    border: 1px solid #d1d5db;
    border-radius: 4px;
    box-sizing: border-box;
    font-size: 13px;
  }
`;

const template = html<DocenHyphenationDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <label class="check-field">
        <fluent-checkbox part="auto" ${ref("autoCheck")}></fluent-checkbox>
        <span ${ref("autoLabel")}></span>
      </label>
      <label class="check-field">
        <fluent-checkbox part="caps" ${ref("capsCheck")}></fluent-checkbox>
        <span ${ref("capsLabel")}></span>
      </label>
      <div class="row">
        <label ${ref("zoneLabel")}></label>
        <input type="text" ${ref("zoneInput")} />
      </div>
      <div class="row">
        <label ${ref("limitLabel")}></label>
        <input type="number" min="0" max="10" ${ref("limitInput")} />
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.apply()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

type FluentCheckbox = HTMLElement & { checked?: boolean };

@customElement({ name: "docen-hyphenation-dialog", template, styles })
export class DocenHyphenationDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable autoCheck?: FluentCheckbox;
  @observable capsCheck?: FluentCheckbox;
  @observable zoneInput?: HTMLInputElement;
  @observable limitInput?: HTMLInputElement;

  @observable autoLabel?: HTMLElement;
  @observable capsLabel?: HTMLElement;
  @observable zoneLabel?: HTMLElement;
  @observable limitLabel?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(opts?: Partial<HyphenationDialogOptions>): void {
    if (this.autoCheck) this.autoCheck.checked = opts?.auto ?? false;
    // Word checkbox is "Hyphenate words in CAPS" -> checked means DO hyphenate caps
    if (this.capsCheck) this.capsCheck.checked = !(opts?.doNotHyphenateCaps ?? true);
    if (this.zoneInput) this.zoneInput.value = "0.25 in";
    if (this.limitInput) this.limitInput.value = String(opts?.limit ?? 0);
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  apply(): void {
    const auto = this.autoCheck?.checked ?? false;
    const doNotHyphenateCaps = !(this.capsCheck?.checked ?? false);
    const limit = Number(this.limitInput?.value ?? 0);
    const zoneStr = this.zoneInput?.value ?? "";
    const zoneTw = zoneStr.includes("cm") ? 567 : 360;

    const result: HyphenationDialogOptions = {
      auto,
      doNotHyphenateCaps,
      zoneTw,
      limit,
    };
    this.$emit("hyphenation:ok", result);
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("hyphenationDialog.title", this);
    if (this.autoLabel) this.autoLabel.textContent = t("hyphenationDialog.auto", this);
    if (this.capsLabel) this.capsLabel.textContent = t("hyphenationDialog.caps", this);
    if (this.zoneLabel) this.zoneLabel.textContent = t("hyphenationDialog.zone", this);
    if (this.limitLabel) this.limitLabel.textContent = t("hyphenationDialog.limit", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenHyphenationDialog;
