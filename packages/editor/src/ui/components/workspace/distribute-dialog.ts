import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(360px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .field {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .radios {
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
`;

const template = html<DocenDistributeDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="field">
        <label ${ref("directionLabel")}></label>
        <div class="radios">
          <fluent-field label-position="after">
            <fluent-radio
              name="direction"
              value="h"
              ?checked="${(x) => x.direction === "h"}"
              @change="${(x) => (x.direction = "h")}"
            ></fluent-radio>
            <label slot="label" ${ref("horizontalLabel")}></label>
          </fluent-field>
          <fluent-field label-position="after">
            <fluent-radio
              name="direction"
              value="v"
              ?checked="${(x) => x.direction === "v"}"
              @change="${(x) => (x.direction = "v")}"
            ></fluent-radio>
            <label slot="label" ${ref("verticalLabel")}></label>
          </fluent-field>
        </div>
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

@customElement({ name: "docen-distribute-dialog", template, styles })
export default class DocenDistributeDialog extends FASTElement {
  @observable direction: "h" | "v" = "h";

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable directionLabel?: HTMLElement;
  @observable horizontalLabel?: HTMLElement;
  @observable verticalLabel?: HTMLElement;
  @observable cancelBtn?: HTMLElement;
  @observable okBtn?: HTMLElement;

  #unobserveLang?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  override disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(direction: "h" | "v" = "h"): void {
    this.direction = direction;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  apply(): void {
    this.$emit("distribute:apply", { direction: this.direction });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("distributeDialog.title", this);
    if (this.directionLabel)
      this.directionLabel.textContent = t("distributeDialog.direction", this);
    if (this.horizontalLabel)
      this.horizontalLabel.textContent = t("distributeDialog.horizontal", this);
    if (this.verticalLabel) this.verticalLabel.textContent = t("distributeDialog.vertical", this);
    if (this.okBtn) this.okBtn.textContent = t("common.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("common.cancel", this);
  }
}
