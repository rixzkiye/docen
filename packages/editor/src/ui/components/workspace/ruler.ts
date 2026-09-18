import type { Editor } from "@docen/docx/core";
import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, resolveDir } from "../../i18n/localize";

export interface RulerTabStop {
  position: number; // in twips
  type: "left" | "center" | "right" | "decimal" | "bar";
}

export interface RulerIndent {
  left?: number; // twips
  right?: number; // twips
  firstLine?: number; // twips
  hanging?: number; // twips
}

const styles = css`
  :host {
    display: block;
    height: 24px;
    background: #f0f0f0;
    border-bottom: 1px solid var(--docen-color-divider, #c8c8c8);
    box-sizing: border-box;
    position: relative;
    user-select: none;
    font-family: var(--docen-font-family, "Segoe UI", -apple-system, sans-serif);
    font-size: 10px;
    color: #404040;
    overflow: visible;
  }

  :host([dir="rtl"]) .unit-toggle-btn,
  :host([data-dir="rtl"]) .unit-toggle-btn,
  :host-context([dir="rtl"]) .unit-toggle-btn {
    left: auto;
    right: 0;
    border-right: none;
    border-left: 1px solid #c8c8c8;
  }

  .ruler-container {
    position: relative;
    height: 100%;
    width: 100%;
    display: flex;
    align-items: stretch;
  }

  .unit-toggle-btn {
    position: absolute;
    left: 0;
    top: 0;
    bottom: 0;
    width: 24px;
    background: #e0e0e0;
    border: none;
    border-right: 1px solid #c8c8c8;
    font-size: 9px;
    font-weight: 700;
    color: #505050;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 5;
    padding: 0;
  }

  .unit-toggle-btn:hover {
    background: #d0d0d0;
  }

  .ruler-track {
    position: relative;
    height: 100%;
    flex: 1;
    overflow: hidden;
  }

  .margin-left-bg,
  .margin-right-bg {
    position: absolute;
    top: 0;
    bottom: 0;
    background: #e4e4e4;
  }

  .content-bg {
    position: absolute;
    top: 0;
    bottom: 0;
    background: #ffffff;
    cursor: crosshair;
  }

  .ticks-svg {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    pointer-events: none;
  }

  /* Indent Markers */
  .marker {
    position: absolute;
    z-index: 10;
    cursor: ew-resize;
    transform: translateX(-50%);
  }

  .first-line-marker {
    top: 0;
    width: 10px;
    height: 8px;
    clip-path: polygon(0 0, 100% 0, 50% 100%);
    background: #4a4a4a;
  }

  .first-line-marker:hover,
  .first-line-marker.active {
    background: #0f6cbd;
  }

  .hanging-marker {
    top: 8px;
    width: 10px;
    height: 8px;
    clip-path: polygon(50% 0, 0 100%, 100% 100%);
    background: #4a4a4a;
  }

  .hanging-marker:hover,
  .hanging-marker.active {
    background: #0f6cbd;
  }

  .left-marker {
    top: 16px;
    width: 10px;
    height: 7px;
    background: #4a4a4a;
    border-radius: 1px;
  }

  .left-marker:hover,
  .left-marker.active {
    background: #0f6cbd;
  }

  .right-marker {
    top: 10px;
    width: 10px;
    height: 12px;
    clip-path: polygon(50% 0, 0 100%, 100% 100%);
    background: #4a4a4a;
  }

  .right-marker:hover,
  .right-marker.active {
    background: #0f6cbd;
  }

  /* Tab Stops */
  .tab-stop-item {
    position: absolute;
    top: 10px;
    width: 8px;
    height: 12px;
    cursor: ew-resize;
    z-index: 9;
    transform: translateX(-50%);
  }

  .tab-stop-item.drag-off {
    opacity: 0.3;
  }

  .tab-stop-icon {
    width: 100%;
    height: 100%;
    display: flex;
  }

  .tab-stop-icon.type-left {
    border-left: 2px solid #0f6cbd;
    border-bottom: 2px solid #0f6cbd;
    height: 7px;
    width: 5px;
  }

  .tab-stop-icon.type-right {
    border-right: 2px solid #0f6cbd;
    border-bottom: 2px solid #0f6cbd;
    height: 7px;
    width: 5px;
    margin-left: 3px;
  }

  .tab-stop-icon.type-center {
    border-left: 2px solid #0f6cbd;
    border-bottom: 2px solid #0f6cbd;
    border-right: 2px solid #0f6cbd;
    height: 7px;
    width: 6px;
  }

  .tab-stop-icon.type-decimal {
    border-left: 2px solid #0f6cbd;
    border-bottom: 2px solid #0f6cbd;
    border-right: 2px solid #0f6cbd;
    height: 7px;
    width: 6px;
    position: relative;
  }

  .tab-stop-icon.type-bar {
    border-left: 2px solid #0f6cbd;
    height: 11px;
    width: 2px;
  }

  /* Tooltip & Guideline */
  .alt-tooltip {
    position: absolute;
    top: -24px;
    transform: translateX(-50%);
    background: #242424;
    color: #ffffff;
    padding: 2px 6px;
    font-size: 11px;
    font-weight: 600;
    border-radius: 3px;
    white-space: nowrap;
    z-index: 100;
    pointer-events: none;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
  }

  .guide-line {
    position: absolute;
    top: 24px;
    bottom: -2000px;
    width: 1px;
    border-left: 1px dashed #0f6cbd;
    z-index: 99;
    pointer-events: none;
  }
`;

const template = html<DocenRuler>`
  <div class="ruler-container" part="container">
    <button
      class="unit-toggle-btn"
      part="unit-toggle"
      title="Toggle Unit (in / cm)"
      @click="${(x) => x.toggleUnit()}"
    >
      ${(x) => x.unit.toUpperCase()}
    </button>
    <div
      class="ruler-track"
      part="track"
      @click="${(x, c) => x.onTrackClick(c.event as MouseEvent)}"
    >
      <div
        class="margin-left-bg"
        style="left: 0; width: ${(x) => x.marginLeftPx * x.scale}px;"
      ></div>
      <div
        class="content-bg"
        style="left: ${(x) => x.marginLeftPx * x.scale}px; width: ${(x) => (x.pageWidthPx - x.marginLeftPx - x.marginRightPx) * x.scale}px;"
      ></div>
      <div
        class="margin-right-bg"
        style="left: ${(x) => (x.pageWidthPx - x.marginRightPx) * x.scale}px; width: ${(x) => x.marginRightPx * x.scale}px;"
      ></div>

      <svg class="ticks-svg" part="ticks" ${ref("ticksSvg")}></svg>

      <!-- First-Line Indent Marker -->
      <div
        class="marker first-line-marker ${(x) => (x.dragActiveMarker === "firstLine" ? "active" : "")}"
        part="marker-first-line"
        title="First Line Indent"
        style="left: ${(x) => x.firstLineMarkerX}px;"
        @pointerdown="${(x, c) => x.onMarkerPointerDown("firstLine", c.event as PointerEvent)}"
      ></div>

      <!-- Hanging Indent Marker -->
      <div
        class="marker hanging-marker ${(x) => (x.dragActiveMarker === "hanging" ? "active" : "")}"
        part="marker-hanging"
        title="Hanging Indent"
        style="left: ${(x) => x.hangingMarkerX}px;"
        @pointerdown="${(x, c) => x.onMarkerPointerDown("hanging", c.event as PointerEvent)}"
      ></div>

      <!-- Left Indent Base Marker -->
      <div
        class="marker left-marker ${(x) => (x.dragActiveMarker === "left" ? "active" : "")}"
        part="marker-left"
        title="Left Indent"
        style="left: ${(x) => x.hangingMarkerX}px;"
        @pointerdown="${(x, c) => x.onMarkerPointerDown("left", c.event as PointerEvent)}"
      ></div>

      <!-- Right Indent Marker -->
      <div
        class="marker right-marker ${(x) => (x.dragActiveMarker === "right" ? "active" : "")}"
        part="marker-right"
        title="Right Indent"
        style="left: ${(x) => x.rightMarkerX}px;"
        @pointerdown="${(x, c) => x.onMarkerPointerDown("right", c.event as PointerEvent)}"
      ></div>

      <!-- Tab Stops -->
      ${(x) => x.renderTabStops()}

      <!-- Alt Precision Tooltip & Guideline -->
      ${(x) =>
        x.showTooltip && x.tooltipText
          ? html`
              <div class="alt-tooltip" style="left: ${x.tooltipX}px;">${x.tooltipText}</div>
              <div class="guide-line" style="left: ${x.tooltipX}px;"></div>
            `
          : ""}
    </div>
  </div>
`;

/**
 * `<docen-ruler>` — Interactive horizontal ruler with draggable indent markers,
 * tab stops, Alt-key precision measurement, and 2-way ProseMirror sync.
 */
@customElement({ name: "docen-ruler", template, styles })
export class DocenRuler extends FASTElement {
  @attr unit: "in" | "cm" = "in";
  @attr dir: "ltr" | "rtl" = "ltr";
  @observable pageWidthPx = 816; // 8.5" * 96
  @observable marginLeftPx = 96; // 1" * 96
  @observable marginRightPx = 96; // 1" * 96
  @observable scale = 1;

  get isRtl(): boolean {
    return (this.dir ?? resolveDir(this)) === "rtl";
  }

  dirChanged(): void {
    this.#syncDir();
  }

  #syncDir(): void {
    const isRtl = this.isRtl;
    this.toggleAttribute("data-dir", isRtl);
    if (isRtl) {
      this.setAttribute("dir", "rtl");
    } else if (this.getAttribute("dir") === "rtl" && !this.getAttribute("data-dir")) {
      // keep explicit attr
    } else {
      this.removeAttribute("dir");
    }
    this.renderTicks();
  }

  #unobserveLang?: () => void;

  // Indent values in twips
  @observable leftIndentTwips = 0;
  @observable firstLineTwips = 0;
  @observable rightIndentTwips = 0;

  // Tab stops
  @observable tabStops: RulerTabStop[] = [];

  // Drag interaction state
  @observable dragActiveMarker: "firstLine" | "hanging" | "left" | "right" | "tabStop" | null =
    null;
  @observable dragTabIdx = -1;
  @observable dragOffRuler = false;
  @observable showTooltip = false;
  @observable tooltipText = "";
  @observable tooltipX = 0;

  @observable ticksSvg?: SVGSVGElement;

  ticksSvgChanged(): void {
    this.renderTicks();
  }

  unitChanged(): void {
    this.renderTicks();
  }

  pageWidthPxChanged(): void {
    this.renderTicks();
  }

  marginLeftPxChanged(): void {
    this.renderTicks();
  }

  marginRightPxChanged(): void {
    this.renderTicks();
  }

  scaleChanged(): void {
    this.renderTicks();
  }

  #editor: Editor | null = null;
  #onTransaction = (): void => {
    this.syncFromEditor();
  };

  #dragStartX = 0;
  #initialLeftTwips = 0;
  #initialFirstLineTwips = 0;
  #initialRightTwips = 0;
  #initialTabTwips = 0;

  override connectedCallback(): void {
    super.connectedCallback();
    if (!this.getAttribute("unit")) {
      const isMetric = !/^en/i.test(navigator.language || "");
      this.unit = isMetric ? "cm" : "in";
    }
    this.#syncDir();
    this.#unobserveLang = observeLang(() => this.#syncDir());
    this.renderTicks();
  }

  override disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    if (this.#editor) {
      this.#editor.off("transaction", this.#onTransaction);
    }
    super.disconnectedCallback();
  }

  override $emit(
    type: string,
    detail?: unknown,
    options?: Omit<CustomEventInit, "detail">,
  ): boolean {
    return this.dispatchEvent(
      new CustomEvent(type, {
        bubbles: true,
        composed: true,
        detail,
        ...options,
      }),
    );
  }

  toggleUnit(): void {
    this.unit = this.unit === "in" ? "cm" : "in";
    this.setAttribute("unit", this.unit);
    this.renderTicks();
    this.$emit("ruler:unit-change", { unit: this.unit });
  }

  // ── Coordinates & Markers Math ──

  get zeroXPx(): number {
    if (this.isRtl) {
      return (this.pageWidthPx - this.marginRightPx) * this.scale;
    }
    return this.marginLeftPx * this.scale;
  }

  get contentWidthPx(): number {
    return (this.pageWidthPx - this.marginLeftPx - this.marginRightPx) * this.scale;
  }

  get hangingMarkerX(): number {
    const px = (this.leftIndentTwips / 15) * this.scale;
    return this.isRtl ? this.zeroXPx - px : this.zeroXPx + px;
  }

  get firstLineMarkerX(): number {
    const px = ((this.leftIndentTwips + this.firstLineTwips) / 15) * this.scale;
    return this.isRtl ? this.zeroXPx - px : this.zeroXPx + px;
  }

  get rightMarkerX(): number {
    const px = (this.rightIndentTwips / 15) * this.scale;
    if (this.isRtl) {
      return this.marginLeftPx * this.scale + px;
    }
    return this.zeroXPx + this.contentWidthPx - px;
  }

  twipsToContentPx(twips: number): number {
    return (twips / 15) * this.scale;
  }

  contentPxToTwips(px: number): number {
    return Math.round((px / this.scale) * 15);
  }

  formatMeasurement(twips: number): string {
    if (this.unit === "cm") {
      const cm = (twips / 1440) * 2.54;
      return `${cm.toFixed(2)} cm`;
    }
    const inches = twips / 1440;
    return `${inches.toFixed(2)} in`;
  }

  // ── SVG Ticks Rendering ──

  renderTicks(): void {
    const svg = this.ticksSvg;
    if (!svg) return;
    const isCm = this.unit === "cm";
    const unitPx = (isCm ? 96 / 2.54 : 96) * this.scale;
    const minorPx = isCm ? unitPx / 10 : unitPx / 8;
    const zeroX = this.zeroXPx;
    const totalW = this.pageWidthPx * this.scale;

    let lines = "";
    let texts = "";

    if (this.isRtl) {
      // Mirrored to right margin: 0 is at zeroX, numbers increase moving leftwards
      for (let x = zeroX; x >= 0; x -= minorPx) {
        const offset = zeroX - x;
        const unitsVal = offset / unitPx;
        const isMajor = Math.abs(unitsVal - Math.round(unitsVal)) < 1e-4;
        const isHalf = Math.abs(unitsVal * 2 - Math.round(unitsVal * 2)) < 1e-4;

        const tickHeight = isMajor ? 10 : isHalf ? 6 : 3;
        const y1 = 24 - tickHeight;
        const y2 = 24;

        lines += `<line x1="${x.toFixed(1)}" y1="${y1}" x2="${x.toFixed(1)}" y2="${y2}" stroke="#a0a0a0" stroke-width="1"/>`;

        if (isMajor) {
          const num = Math.round(unitsVal);
          const label = String(Math.abs(num));
          texts += `<text x="${x.toFixed(1)}" y="10" font-size="8" text-anchor="middle" fill="#606060">${label}</text>`;
        }
      }
      for (let x = zeroX + minorPx; x <= totalW; x += minorPx) {
        const offset = x - zeroX;
        const unitsVal = offset / unitPx;
        const isMajor = Math.abs(unitsVal - Math.round(unitsVal)) < 1e-4;
        const isHalf = Math.abs(unitsVal * 2 - Math.round(unitsVal * 2)) < 1e-4;

        const tickHeight = isMajor ? 10 : isHalf ? 6 : 3;
        const y1 = 24 - tickHeight;
        const y2 = 24;

        lines += `<line x1="${x.toFixed(1)}" y1="${y1}" x2="${x.toFixed(1)}" y2="${y2}" stroke="#a0a0a0" stroke-width="1"/>`;

        if (isMajor) {
          const num = Math.round(unitsVal);
          const label = String(Math.abs(num));
          texts += `<text x="${x.toFixed(1)}" y="10" font-size="8" text-anchor="middle" fill="#606060">${label}</text>`;
        }
      }
    } else {
      const startOffset = Math.ceil(-zeroX / minorPx) * minorPx;
      const endOffset = totalW - zeroX;

      for (let offset = startOffset; offset <= endOffset; offset += minorPx) {
        const x = zeroX + offset;
        const unitsVal = offset / unitPx;
        const isMajor = Math.abs(unitsVal - Math.round(unitsVal)) < 1e-4;
        const isHalf = Math.abs(unitsVal * 2 - Math.round(unitsVal * 2)) < 1e-4;

        const tickHeight = isMajor ? 10 : isHalf ? 6 : 3;
        const y1 = 24 - tickHeight;
        const y2 = 24;

        lines += `<line x1="${x.toFixed(1)}" y1="${y1}" x2="${x.toFixed(1)}" y2="${y2}" stroke="#a0a0a0" stroke-width="1"/>`;

        if (isMajor) {
          const num = Math.round(unitsVal);
          const label = String(Math.abs(num));
          texts += `<text x="${x.toFixed(1)}" y="10" font-size="8" text-anchor="middle" fill="#606060">${label}</text>`;
        }
      }
    }

    svg.innerHTML = `<g>${lines}</g><g>${texts}</g>`;
  }

  // ── Tab Stops Rendering ──

  renderTabStops(): ReturnType<typeof html> {
    return html`
      ${this.tabStops.map((stop, idx) => {
        const x = this.isRtl
          ? this.zeroXPx - this.twipsToContentPx(stop.position)
          : this.zeroXPx + this.twipsToContentPx(stop.position);
        const isDraggingThis = this.dragActiveMarker === "tabStop" && this.dragTabIdx === idx;
        return html`
          <div
            class="tab-stop-item ${isDraggingThis && this.dragOffRuler ? "drag-off" : ""}"
            style="left: ${x}px;"
            title="Tab Stop: ${this.formatMeasurement(stop.position)} (${stop.type})"
            @pointerdown="${(ruler: DocenRuler, c) => ruler.onTabPointerDown(idx, c.event as PointerEvent)}"
            @dblclick="${(ruler: DocenRuler, c) => ruler.onTabDblClick(idx, c.event as MouseEvent)}"
          >
            <div class="tab-stop-icon type-${stop.type}"></div>
          </div>
        `;
      })}
    `;
  }

  // ── Pointer Drag Handlers ──

  onMarkerPointerDown(
    marker: "firstLine" | "hanging" | "left" | "right",
    event: PointerEvent,
  ): void {
    event.preventDefault();
    event.stopPropagation();
    (event.target as HTMLElement).setPointerCapture(event.pointerId);

    this.dragActiveMarker = marker;
    this.#dragStartX = event.clientX;
    this.#initialLeftTwips = this.leftIndentTwips;
    this.#initialFirstLineTwips = this.firstLineTwips;
    this.#initialRightTwips = this.rightIndentTwips;

    this.showTooltip = event.altKey;
    this.#updateTooltip(event.clientX);

    const onPointerMove = (e: PointerEvent): void => {
      const deltaPx = e.clientX - this.#dragStartX;
      const rawDeltaTwips = this.contentPxToTwips(deltaPx);
      const deltaTwips = this.isRtl ? -rawDeltaTwips : rawDeltaTwips;

      if (this.dragActiveMarker === "firstLine") {
        this.firstLineTwips = this.#initialFirstLineTwips + deltaTwips;
      } else if (this.dragActiveMarker === "hanging") {
        // Dragging hanging moves left indent while keeping first line absolute position stationary
        const newLeft = Math.max(0, this.#initialLeftTwips + deltaTwips);
        const shift = newLeft - this.#initialLeftTwips;
        this.leftIndentTwips = newLeft;
        this.firstLineTwips = this.#initialFirstLineTwips - shift;
      } else if (this.dragActiveMarker === "left") {
        // Dragging left base moves both left and first-line together
        const newLeft = Math.max(0, this.#initialLeftTwips + deltaTwips);
        this.leftIndentTwips = newLeft;
      } else if (this.dragActiveMarker === "right") {
        this.rightIndentTwips = Math.max(0, this.#initialRightTwips - deltaTwips);
      }

      this.showTooltip = e.altKey;
      this.#updateTooltip(e.clientX);
      this.#emitIndentChange();
    };

    const onPointerUp = (_e: PointerEvent): void => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
      this.dragActiveMarker = null;
      this.showTooltip = false;
      this.commitIndentToEditor();
    };

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  onTabPointerDown(idx: number, event: PointerEvent): void {
    event.preventDefault();
    event.stopPropagation();
    (event.target as HTMLElement).setPointerCapture(event.pointerId);

    this.dragActiveMarker = "tabStop";
    this.dragTabIdx = idx;
    this.#dragStartX = event.clientX;
    const stop = this.tabStops[idx];
    if (!stop) return;
    this.#initialTabTwips = stop.position;
    this.dragOffRuler = false;

    this.showTooltip = event.altKey;
    this.#updateTooltip(event.clientX);

    const trackRect = this.shadowRoot?.querySelector(".ruler-track")?.getBoundingClientRect();

    const onPointerMove = (e: PointerEvent): void => {
      if (trackRect) {
        // Check vertical distance off ruler
        const dY = e.clientY - (trackRect.top + trackRect.height / 2);
        this.dragOffRuler = Math.abs(dY) > 28;
      }

      const deltaPx = e.clientX - this.#dragStartX;
      const rawDeltaTwips = this.contentPxToTwips(deltaPx);
      const deltaTwips = this.isRtl ? -rawDeltaTwips : rawDeltaTwips;
      const newTwips = Math.max(0, this.#initialTabTwips + deltaTwips);

      if (this.tabStops[idx]) {
        this.tabStops = this.tabStops.map((s, i) => (i === idx ? { ...s, position: newTwips } : s));
      }

      this.showTooltip = e.altKey;
      this.#updateTooltip(e.clientX);
    };

    const onPointerUp = (_e: PointerEvent): void => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);

      if (this.dragOffRuler) {
        // Remove tab stop!
        const removed = this.tabStops[idx];
        this.tabStops = this.tabStops.filter((_, i) => i !== idx);
        this.$emit("ruler:tabstop-remove", { tabStop: removed });
      } else {
        this.$emit("ruler:tabstop-change", { tabStops: this.tabStops });
      }

      this.dragActiveMarker = null;
      this.dragTabIdx = -1;
      this.dragOffRuler = false;
      this.showTooltip = false;
      this.commitTabStopsToEditor();
    };

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  onTabDblClick(idx: number, event: MouseEvent): void {
    event.stopPropagation();
    const stop = this.tabStops[idx];
    this.$emit("ruler:open-tabs", { tabStop: stop });
    this.#dispatchTabsDialog();
  }

  onTrackClick(event: MouseEvent): void {
    const track = event.currentTarget as HTMLElement | null;
    if (!track) return;
    const rect = track.getBoundingClientRect();
    const clickX = event.clientX - rect.left;

    // Check if clicked inside content zone
    const contentLeft = this.marginLeftPx * this.scale;
    const contentRight = (this.pageWidthPx - this.marginRightPx) * this.scale;

    if (clickX >= contentLeft && clickX <= contentRight) {
      const offsetPx = this.isRtl ? contentRight - clickX : clickX - contentLeft;
      const twips = this.contentPxToTwips(offsetPx);
      const newStop: RulerTabStop = { position: twips, type: "left" };
      this.tabStops = [...this.tabStops, newStop].sort((a, b) => a.position - b.position);
      this.$emit("ruler:tabstop-add", { tabStop: newStop });
      this.commitTabStopsToEditor();
    }
  }

  #updateTooltip(clientX: number): void {
    const track = this.shadowRoot?.querySelector(".ruler-track")?.getBoundingClientRect();
    if (!track) return;
    this.tooltipX = clientX - track.left;

    if (this.dragActiveMarker === "firstLine") {
      this.tooltipText = `First Line: ${this.formatMeasurement(this.firstLineTwips)}`;
    } else if (this.dragActiveMarker === "hanging" || this.dragActiveMarker === "left") {
      this.tooltipText = `Left Indent: ${this.formatMeasurement(this.leftIndentTwips)}`;
    } else if (this.dragActiveMarker === "right") {
      this.tooltipText = `Right Indent: ${this.formatMeasurement(this.rightIndentTwips)}`;
    } else if (this.dragActiveMarker === "tabStop" && this.dragTabIdx >= 0) {
      const stop = this.tabStops[this.dragTabIdx];
      this.tooltipText = stop ? `Tab: ${this.formatMeasurement(stop.position)} (${stop.type})` : "";
    }
  }

  #emitIndentChange(): void {
    this.$emit("ruler:indent-change", {
      indent: {
        left: this.leftIndentTwips,
        right: this.rightIndentTwips,
        firstLine: this.firstLineTwips >= 0 ? this.firstLineTwips : undefined,
        hanging: this.firstLineTwips < 0 ? -this.firstLineTwips : undefined,
      },
    });
  }

  commitIndentToEditor(): void {
    this.#emitIndentChange();
    if (!this.#editor) return;
    const { tr, selection } = this.#editor.state;
    const { $from } = selection;
    for (let d = $from.depth; d > 0; d--) {
      const node = $from.node(d);
      if (node.type.name === "paragraph" || node.type.name === "heading") {
        const pos = $from.before(d);
        const indentObj: Record<string, number | undefined> = {
          left: this.leftIndentTwips || undefined,
          right: this.rightIndentTwips || undefined,
          firstLine: this.firstLineTwips > 0 ? this.firstLineTwips : undefined,
          hanging: this.firstLineTwips < 0 ? -this.firstLineTwips : undefined,
        };
        const cleanIndent = Object.values(indentObj).some((v) => v !== undefined)
          ? indentObj
          : null;
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          indent: cleanIndent,
        });
        this.#editor.view.dispatch(tr);
        return;
      }
    }
  }

  commitTabStopsToEditor(): void {
    this.$emit("ruler:tabstop-change", { tabStops: this.tabStops });
    if (!this.#editor) return;
    const { tr, selection } = this.#editor.state;
    const { $from } = selection;
    for (let d = $from.depth; d > 0; d--) {
      const node = $from.node(d);
      if (node.type.name === "paragraph" || node.type.name === "heading") {
        const pos = $from.before(d);
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          tabStops: this.tabStops.length > 0 ? this.tabStops : null,
        });
        this.#editor.view.dispatch(tr);
        return;
      }
    }
  }

  #dispatchTabsDialog(): void {
    const hostEl =
      (this.closest("docen-document") as HTMLElement | null) ??
      document.querySelector("docen-document");
    if (hostEl) {
      hostEl.dispatchEvent(
        new CustomEvent("command", {
          bubbles: true,
          composed: true,
          detail: { event: "tabs-dialog" },
        }),
      );
    }
  }

  // ── Two-Way Editor Synchronization ──

  bindEditor(editor: Editor): void {
    if (this.#editor) {
      this.#editor.off("transaction", this.#onTransaction);
    }
    this.#editor = editor;
    editor.on("transaction", this.#onTransaction);
    this.syncFromEditor();
  }

  syncFromEditor(): void {
    if (!this.#editor) return;
    const { selection } = this.#editor.state;
    const { $from } = selection;
    for (let d = $from.depth; d > 0; d--) {
      const node = $from.node(d);
      if (node.type.name === "paragraph" || node.type.name === "heading") {
        const indent = node.attrs.indent as RulerIndent | undefined;
        this.leftIndentTwips = indent?.left ?? 0;
        this.rightIndentTwips = indent?.right ?? 0;
        this.firstLineTwips = indent?.firstLine ?? (indent?.hanging ? -indent.hanging : 0);

        const stops = (node.attrs.tabStops ?? node.attrs.tabs ?? []) as RulerTabStop[];
        this.tabStops = Array.isArray(stops) ? [...stops] : [];
        return;
      }
    }
  }

  setParagraphAttrs(
    indent?: RulerIndent,
    tabStops?: RulerTabStop[],
    geometry?: {
      pageWidthPx?: number;
      marginLeftPx?: number;
      marginRightPx?: number;
      scale?: number;
    },
  ): void {
    if (geometry) {
      if (geometry.pageWidthPx != null) this.pageWidthPx = geometry.pageWidthPx;
      if (geometry.marginLeftPx != null) this.marginLeftPx = geometry.marginLeftPx;
      if (geometry.marginRightPx != null) this.marginRightPx = geometry.marginRightPx;
      if (geometry.scale != null) this.scale = geometry.scale;
    }
    if (indent) {
      this.leftIndentTwips = indent.left ?? 0;
      this.rightIndentTwips = indent.right ?? 0;
      this.firstLineTwips = indent.firstLine ?? (indent.hanging ? -indent.hanging : 0);
    }
    if (tabStops) {
      this.tabStops = [...tabStops];
    }
  }
}

export default DocenRuler;
