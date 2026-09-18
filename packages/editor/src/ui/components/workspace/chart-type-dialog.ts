import {
  FASTElement,
  css,
  customElement,
  html,
  observable,
  ref,
  repeat,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";
import "./dialog";

export interface ChartTypeOption {
  type: string;
  key: string;
  defaultLabel: string;
}

export const CHART_TYPES: readonly ChartTypeOption[] = [
  { type: "column", key: "chart-type.column", defaultLabel: "Clustered Column" },
  { type: "bar", key: "chart-type.bar", defaultLabel: "Clustered Bar" },
  { type: "line", key: "chart-type.line", defaultLabel: "Line" },
  { type: "area", key: "chart-type.area", defaultLabel: "Area" },
  { type: "pie", key: "chart-type.pie", defaultLabel: "Pie" },
  { type: "ofPie", key: "chart-type.of-pie", defaultLabel: "Of Pie" },
  { type: "doughnut", key: "chart-type.doughnut", defaultLabel: "Doughnut" },
  { type: "scatter", key: "chart-type.scatter", defaultLabel: "Scatter" },
  { type: "bubble", key: "chart-type.bubble", defaultLabel: "Bubble" },
  { type: "radar", key: "chart-type.radar", defaultLabel: "Radar" },
  { type: "stock", key: "chart-type.stock", defaultLabel: "Stock" },
  { type: "surface", key: "chart-type.surface", defaultLabel: "Surface" },
  { type: "combo", key: "chart-type.combo", defaultLabel: "Combo" },
] as const;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(520px, 94vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .types-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 8px;
    max-height: 360px;
    overflow-y: auto;
    padding: 2px;
  }
  .type-card {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 52px;
    padding: 8px 6px;
    border-radius: 4px;
    border: 1px solid var(--neutral-stroke-rest, #d1d1d1);
    background: var(--neutral-fill-card-rest, #fafafa);
    cursor: pointer;
    font-size: 12px;
    text-align: center;
    transition: all 0.15s ease;
    user-select: none;
    box-sizing: border-box;
  }
  .type-card:hover {
    background: var(--neutral-fill-hover, #f0f0f0);
    border-color: var(--neutral-stroke-hover, #b0b0b0);
  }
  .type-card.selected {
    background: var(--accent-fill-subtle-selected, #e8f3ff);
    border-color: var(--accent-fill-rest, #2b7cd3);
    color: var(--accent-foreground-rest, #005a9e);
    font-weight: 600;
  }
`;

const template = html<DocenChartTypeDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="types-grid" ${ref("typesGrid")}>
        ${repeat(
          () => CHART_TYPES,
          html<ChartTypeOption, DocenChartTypeDialog>`
            <div
              class="type-card ${(item, c) => (c.parent.selectedType === item.type ? "selected" : "")}"
              data-type="${(item) => item.type}"
              @click="${(item, c) => c.parent.selectType(item.type)}"
            >
              ${(item, c) => c.parent.getLabel(item)}
            </div>
          `,
        )}
      </div>
    </div>
    <div slot="action">
      <fluent-button
        data-action="cancel"
        ${ref("cancelBtn")}
        @click="${(x) => x.hide()}"
      ></fluent-button>
      <fluent-button
        data-action="ok"
        appearance="accent"
        ${ref("okBtn")}
        @click="${(x) => x.apply()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-chart-type-dialog>` — modal dialog allowing the user to select
 * from all native chart types and change the selected chart.
 */
@customElement({ name: "docen-chart-type-dialog", template, styles })
export default class DocenChartTypeDialog extends FASTElement {
  @observable selectedType = "column";

  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable typesGrid?: HTMLElement;
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

  getLabel(item: ChartTypeOption): string {
    const loc = t(`ribbon.opt.${item.key}`, this);
    return loc && loc !== `ribbon.opt.${item.key}` ? loc : item.defaultLabel;
  }

  selectType(type: string): void {
    this.selectedType = type;
  }

  show(currentType = "column"): void {
    this.selectedType = currentType;
    if (this.dialogEl && typeof this.dialogEl.show === "function") {
      this.dialogEl.show();
    }
    this.setAttribute("open", "");
  }

  hide(): void {
    if (this.dialogEl && typeof this.dialogEl.hide === "function") {
      this.dialogEl.hide();
    }
    this.removeAttribute("open");
  }

  apply(): void {
    this.dispatchEvent(
      new CustomEvent("chart-type:ok", {
        detail: { type: this.selectedType },
        bubbles: true,
        composed: true,
      }),
    );
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) {
      this.dialogEl.heading = t("chart-type.dialog.title", this) || "Change Chart Type";
    }
    if (this.cancelBtn) this.cancelBtn.textContent = t("action.cancel", this) || "Cancel";
    if (this.okBtn) this.okBtn.textContent = t("action.ok", this) || "OK";
  }
}
