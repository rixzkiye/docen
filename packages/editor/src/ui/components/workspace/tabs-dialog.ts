import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface TabStopEntry {
  position: number; // in twips
  type: "left" | "center" | "right" | "decimal" | "bar";
  leader?: "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
}

export interface TabsDialogOptions {
  tabStops?: TabStopEntry[];
  defaultTabStop?: number; // twips (default 720 = 0.5")
}

function parsePositionTwip(str: string): number | null {
  const trimmed = str.trim();
  if (!trimmed) return null;
  const match = /^([\d.]+)\s*(in|inch|"|cm|mm|pt)?$/i.exec(trimmed);
  if (!match) return null;
  const val = parseFloat(match[1]);
  if (isNaN(val) || val < 0) return null;
  const unit = (match[2] || "").toLowerCase();
  if (unit === "cm") return Math.round(val * 567);
  if (unit === "mm") return Math.round(val * 56.7);
  if (unit === "pt") return Math.round(val * 20);
  // Default to inches
  return Math.round(val * 1440);
}

function formatPositionTwip(twip: number, metric = false): string {
  if (metric) {
    const cm = twip / 567;
    return `${cm.toFixed(2).replace(/\.?0+$/, "")} cm`;
  }
  const inches = twip / 1440;
  return `${inches.toFixed(2).replace(/\.?0+$/, "")}"`;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(420px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
  }
  .top-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    align-items: start;
  }
  .input-col {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .input-col label {
    font-size: 12px;
    font-weight: 500;
  }
  .text-input {
    padding: 4px 8px;
    border: 1px solid #d1d5db;
    border-radius: 4px;
    font-size: 13px;
    box-sizing: border-box;
    width: 100%;
  }
  .listbox {
    border: 1px solid #d1d5db;
    border-radius: 4px;
    height: 90px;
    overflow-y: auto;
    background: #fff;
    padding: 2px 0;
  }
  .list-item {
    padding: 3px 8px;
    cursor: pointer;
    user-select: none;
  }
  .list-item:hover {
    background: #f3f4f6;
  }
  .list-item.selected {
    background: #e0e7ff;
    color: #1e40af;
    font-weight: 500;
  }
  .section-box {
    border: 1px solid #e5e7eb;
    border-radius: 6px;
    padding: 8px 12px;
  }
  .section-legend {
    font-size: 12px;
    font-weight: 600;
    color: #374151;
    margin-bottom: 6px;
  }
  .radio-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 6px;
  }
  .radio-grid-4 {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 6px;
  }
  .radio-label {
    display: flex;
    align-items: center;
    gap: 4px;
    font-size: 12px;
    cursor: pointer;
  }
  .btn-row {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 4px;
  }
`;

const template = html<DocenTabsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="top-grid">
        <div class="input-col">
          <label ${ref("posLabel")}></label>
          <input
            type="text"
            class="text-input"
            ${ref("posInput")}
            @input="${(x) => x.onPosInput()}"
          />
          <div class="listbox" ${ref("listboxEl")}>
            ${(x) =>
              x.stops.map(
                (s, i) => html`
                  <div
                    class="${i === x.selectedIndex ? "selected" : ""} list-item"
                    @click="${() => x.selectIndex(i)}"
                  >
                    ${formatPositionTwip(s.position, x.isMetric)}
                  </div>
                `,
              )}
          </div>
        </div>
        <div class="input-col">
          <label ${ref("defaultLabel")}></label>
          <input type="text" class="text-input" ${ref("defaultInput")} />
        </div>
      </div>

      <div class="section-box">
        <div class="section-legend" ${ref("alignLegend")}></div>
        <div class="radio-grid">
          <label class="radio-label">
            <input
              type="radio"
              name="tab-align"
              value="left"
              checked="${(x) => x.selectedAlign === "left"}"
              @change="${(x) => (x.selectedAlign = "left")}"
            />
            <span ${ref("alignLeftLabel")}></span>
          </label>
          <label class="radio-label">
            <input
              type="radio"
              name="tab-align"
              value="center"
              checked="${(x) => x.selectedAlign === "center"}"
              @change="${(x) => (x.selectedAlign = "center")}"
            />
            <span ${ref("alignCenterLabel")}></span>
          </label>
          <label class="radio-label">
            <input
              type="radio"
              name="tab-align"
              value="right"
              checked="${(x) => x.selectedAlign === "right"}"
              @change="${(x) => (x.selectedAlign = "right")}"
            />
            <span ${ref("alignRightLabel")}></span>
          </label>
          <label class="radio-label">
            <input
              type="radio"
              name="tab-align"
              value="decimal"
              checked="${(x) => x.selectedAlign === "decimal"}"
              @change="${(x) => (x.selectedAlign = "decimal")}"
            />
            <span ${ref("alignDecimalLabel")}></span>
          </label>
          <label class="radio-label">
            <input
              type="radio"
              name="tab-align"
              value="bar"
              checked="${(x) => x.selectedAlign === "bar"}"
              @change="${(x) => (x.selectedAlign = "bar")}"
            />
            <span ${ref("alignBarLabel")}></span>
          </label>
        </div>
      </div>

      <div class="section-box">
        <div class="section-legend" ${ref("leaderLegend")}></div>
        <div class="radio-grid-4">
          <label class="radio-label">
            <input
              type="radio"
              name="tab-leader"
              value="none"
              checked="${(x) => x.selectedLeader === "none"}"
              @change="${(x) => (x.selectedLeader = "none")}"
            />
            <span ${ref("leaderNoneLabel")}></span>
          </label>
          <label class="radio-label">
            <input
              type="radio"
              name="tab-leader"
              value="dot"
              checked="${(x) => x.selectedLeader === "dot"}"
              @change="${(x) => (x.selectedLeader = "dot")}"
            />
            <span>2 ......</span>
          </label>
          <label class="radio-label">
            <input
              type="radio"
              name="tab-leader"
              value="hyphen"
              checked="${(x) => x.selectedLeader === "hyphen"}"
              @change="${(x) => (x.selectedLeader = "hyphen")}"
            />
            <span>3 ------</span>
          </label>
          <label class="radio-label">
            <input
              type="radio"
              name="tab-leader"
              value="underscore"
              checked="${(x) => x.selectedLeader === "underscore"}"
              @change="${(x) => (x.selectedLeader = "underscore")}"
            />
            <span>4 ______</span>
          </label>
        </div>
      </div>

      <div class="btn-row">
        <fluent-button ${ref("setBtn")} @click="${(x) => x.setStop()}"></fluent-button>
        <fluent-button ${ref("clearBtn")} @click="${(x) => x.clearStop()}"></fluent-button>
        <fluent-button ${ref("clearAllBtn")} @click="${(x) => x.clearAllStops()}"></fluent-button>
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

@customElement({ name: "docen-tabs-dialog", template, styles })
export class DocenTabsDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable posInput?: HTMLInputElement;
  @observable defaultInput?: HTMLInputElement;
  @observable listboxEl?: HTMLElement;

  @observable posLabel?: HTMLElement;
  @observable defaultLabel?: HTMLElement;
  @observable alignLegend?: HTMLElement;
  @observable alignLeftLabel?: HTMLElement;
  @observable alignCenterLabel?: HTMLElement;
  @observable alignRightLabel?: HTMLElement;
  @observable alignDecimalLabel?: HTMLElement;
  @observable alignBarLabel?: HTMLElement;

  @observable leaderLegend?: HTMLElement;
  @observable leaderNoneLabel?: HTMLElement;

  @observable setBtn?: HTMLElement;
  @observable clearBtn?: HTMLElement;
  @observable clearAllBtn?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  @observable stops: TabStopEntry[] = [];
  @observable selectedIndex = -1;
  @observable selectedAlign: TabStopEntry["type"] = "left";
  @observable selectedLeader: "none" | "dot" | "hyphen" | "underscore" = "none";
  @observable isMetric = false;

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.isMetric = !/^en/i.test(navigator.language || "");
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => {
      this.isMetric = !/^en/i.test(navigator.language || "");
      this.#applyLabels();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(opts?: TabsDialogOptions): void {
    this.stops = (opts?.tabStops ?? []).map((s) => ({ ...s }));
    this.selectedIndex = this.stops.length > 0 ? 0 : -1;
    this.selectedAlign = this.stops[0]?.type ?? "left";
    const lead = this.stops[0]?.leader;
    this.selectedLeader =
      lead === "dot" || lead === "hyphen" || lead === "underscore" ? lead : "none";

    if (this.posInput) {
      this.posInput.value =
        this.stops.length > 0 ? formatPositionTwip(this.stops[0].position, this.isMetric) : "";
    }
    if (this.defaultInput) {
      this.defaultInput.value = formatPositionTwip(opts?.defaultTabStop ?? 720, this.isMetric);
    }
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  selectIndex(index: number): void {
    this.selectedIndex = index;
    const stop = this.stops[index];
    if (stop) {
      if (this.posInput) this.posInput.value = formatPositionTwip(stop.position, this.isMetric);
      this.selectedAlign = stop.type;
      const lead = stop.leader;
      this.selectedLeader =
        lead === "dot" || lead === "hyphen" || lead === "underscore" ? lead : "none";
    }
  }

  onPosInput(): void {
    const val = this.posInput?.value ?? "";
    const twip = parsePositionTwip(val);
    if (twip != null) {
      const idx = this.stops.findIndex((s) => Math.abs(s.position - twip) < 15);
      if (idx !== -1) {
        this.selectedIndex = idx;
        this.selectedAlign = this.stops[idx].type;
        const lead = this.stops[idx].leader;
        this.selectedLeader =
          lead === "dot" || lead === "hyphen" || lead === "underscore" ? lead : "none";
      }
    }
  }

  setStop(): void {
    const val = this.posInput?.value ?? "";
    const twip = parsePositionTwip(val);
    if (twip == null) return;
    const leader = this.selectedLeader === "none" ? undefined : this.selectedLeader;
    const entry: TabStopEntry = {
      position: twip,
      type: this.selectedAlign,
      leader,
    };
    const next = this.stops.filter((s) => Math.abs(s.position - twip) >= 15);
    next.push(entry);
    next.sort((a, b) => a.position - b.position);
    this.stops = next;
    this.selectedIndex = this.stops.findIndex((s) => Math.abs(s.position - twip) < 15);
  }

  clearStop(): void {
    if (this.selectedIndex >= 0 && this.selectedIndex < this.stops.length) {
      this.stops = this.stops.filter((_, i) => i !== this.selectedIndex);
      this.selectedIndex = -1;
      if (this.posInput) this.posInput.value = "";
    }
  }

  clearAllStops(): void {
    this.stops = [];
    this.selectedIndex = -1;
    if (this.posInput) this.posInput.value = "";
  }

  apply(): void {
    // If user typed a position and didn't press Set, commit it now
    if (this.posInput?.value) {
      this.setStop();
    }
    const defTwip = this.defaultInput ? parsePositionTwip(this.defaultInput.value) : undefined;
    this.$emit("tabs:ok", {
      tabStops: this.stops,
      defaultTabStop: defTwip ?? undefined,
    });
    this.hide();
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("tabsDialog.title", this);
    if (this.posLabel) this.posLabel.textContent = t("tabsDialog.position", this);
    if (this.defaultLabel) this.defaultLabel.textContent = t("tabsDialog.default", this);
    if (this.alignLegend) this.alignLegend.textContent = t("tabsDialog.alignment", this);
    if (this.alignLeftLabel) this.alignLeftLabel.textContent = t("tabsDialog.left", this);
    if (this.alignCenterLabel) this.alignCenterLabel.textContent = t("tabsDialog.center", this);
    if (this.alignRightLabel) this.alignRightLabel.textContent = t("tabsDialog.right", this);
    if (this.alignDecimalLabel) this.alignDecimalLabel.textContent = t("tabsDialog.decimal", this);
    if (this.alignBarLabel) this.alignBarLabel.textContent = t("tabsDialog.bar", this);
    if (this.leaderLegend) this.leaderLegend.textContent = t("tabsDialog.leader", this);
    if (this.leaderNoneLabel) this.leaderNoneLabel.textContent = t("tabsDialog.none", this);
    if (this.setBtn) this.setBtn.textContent = t("tabsDialog.set", this);
    if (this.clearBtn) this.clearBtn.textContent = t("tabsDialog.clear", this);
    if (this.clearAllBtn) this.clearAllBtn.textContent = t("tabsDialog.clearAll", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenTabsDialog;
