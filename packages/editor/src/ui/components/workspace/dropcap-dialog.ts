import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export type DropCapPosition = "none" | "dropped" | "margin";

export interface DropCapOptions {
  position: DropCapPosition;
  fontFamily?: string;
  lines: number;
  distancePt: number;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(440px, 94vw);
  }
  .dropcap-body {
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
    margin-bottom: 6px;
  }
  .positions {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 10px;
  }
  .pos-card {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 6px;
    padding: 10px 4px;
    border: 2px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 6px;
    cursor: pointer;
    background: var(--docen-color-canvas, #ffffff);
    text-align: center;
    user-select: none;
  }
  .pos-card.selected {
    border-color: var(--docen-color-accent, #0f6cbd);
    background: var(--docen-color-accent-subtle, #ebf3fc);
  }
  .pos-icon {
    width: 44px;
    height: 44px;
    border: 1px dashed var(--docen-color-divider, #888);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 20px;
    font-weight: bold;
  }
  .pos-label {
    font-size: 12px;
  }
  .row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
  }
  .row label {
    font-size: 12px;
  }
  .row input[type="number"],
  .row select {
    padding: 4px 8px;
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 4px;
    width: 140px;
  }
  .footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 10px;
  }
`;

const template = html<DocenDropCapDialog>`
  <docen-dialog ${ref("dialog")} modal heading="${(x) => t("dropcap.dialogTitle", x)}">
    <div class="dropcap-body">
      <div class="section-title">${(x) => t("dropcap.position", x)}</div>
      <div class="positions">
        <div
          class="pos-card ${(x) => (x.selectedPos === "none" ? "selected" : "")}"
          @click="${(x) => x.selectPos("none")}"
        >
          <div class="pos-icon">A</div>
          <span class="pos-label">${(x) => t("dropcap.none", x)}</span>
        </div>
        <div
          class="pos-card ${(x) => (x.selectedPos === "dropped" ? "selected" : "")}"
          @click="${(x) => x.selectPos("dropped")}"
        >
          <div class="pos-icon" style="color: var(--docen-color-accent, #0f6cbd);">A</div>
          <span class="pos-label">${(x) => t("dropcap.dropped", x)}</span>
        </div>
        <div
          class="pos-card ${(x) => (x.selectedPos === "margin" ? "selected" : "")}"
          @click="${(x) => x.selectPos("margin")}"
        >
          <div class="pos-icon" style="border-left: 2px solid var(--docen-color-accent, #0f6cbd);">
            A
          </div>
          <span class="pos-label">${(x) => t("dropcap.inMargin", x)}</span>
        </div>
      </div>

      <div class="section-title">${(x) => t("dropcap.options", x)}</div>
      <div class="row">
        <label>${(x) => t("dropcap.font", x)}:</label>
        <select ${ref("fontSelect")}>
          <option value="">(Default Font)</option>
          <option value="Calibri">Calibri</option>
          <option value="Aptos">Aptos</option>
          <option value="Times New Roman">Times New Roman</option>
          <option value="Georgia">Georgia</option>
          <option value="Garamond">Garamond</option>
          <option value="Arial">Arial</option>
        </select>
      </div>

      <div class="row">
        <label>${(x) => t("dropcap.linesToDrop", x)}:</label>
        <input type="number" min="1" max="10" value="3" ${ref("linesInput")} />
      </div>

      <div class="row">
        <label>${(x) => t("dropcap.distanceFromText", x)} (pt):</label>
        <input type="number" min="0" max="72" step="1" value="0" ${ref("distanceInput")} />
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

@customElement({ name: "docen-dropcap-dialog", template, styles })
export class DocenDropCapDialog extends FASTElement {
  @observable dialog?: HTMLElement & { open: boolean };
  @observable selectedPos: DropCapPosition = "dropped";
  @observable fontSelect?: HTMLSelectElement;
  @observable linesInput?: HTMLInputElement;
  @observable distanceInput?: HTMLInputElement;

  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => {});
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  show(current?: Partial<DropCapOptions>): void {
    if (current?.position) this.selectedPos = current.position;
    if (this.linesInput && typeof current?.lines === "number") {
      this.linesInput.value = String(current.lines);
    }
    if (this.dialog) this.dialog.open = true;
  }

  close(): void {
    if (this.dialog) this.dialog.open = false;
  }

  selectPos(pos: DropCapPosition): void {
    this.selectedPos = pos;
  }

  onOk(): void {
    const lines = parseInt(this.linesInput?.value ?? "3", 10) || 3;
    const distancePt = parseFloat(this.distanceInput?.value ?? "0") || 0;
    const font = this.fontSelect?.value || undefined;

    const options: DropCapOptions = {
      position: this.selectedPos,
      fontFamily: font,
      lines,
      distancePt,
    };

    this.$emit("dropcap:apply", options);
    this.close();
  }
}
