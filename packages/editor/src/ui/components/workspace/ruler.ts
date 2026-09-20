import type { Editor } from "@docen/docx/core";
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

import { observeLang, resolveDir, t } from "../../i18n/localize";
import { RULER_TICK_LEN, rulerTicks } from "./ruler-ticks";
import type { TabSelectorType } from "./tab-selector";

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

/** The strip thickness in CSS px (Word's horizontal ruler is 20px tall). */
export const HORIZONTAL_RULER_HEIGHT = 20;

/** SVG path (viewBox 0 0 9 10) per Word tab-stop type: left ⌊, center ⊤,
 *  right ⌋, decimal ⌊ with a point, bar |. */
function tabGlyphPath(type: RulerTabStop["type"]): string {
  switch (type) {
    case "center":
      return "M0.5 1.5 H8.5 M4.5 1.5 V9";
    case "right":
      return "M8.5 1 V9 H0.5";
    case "bar":
      return "M4.5 0 V10";
    case "decimal":
      return "M0.5 1 V9 H8.5";
    default:
      return "M0.5 1 V9 H8.5";
  }
}

const styles = css`
  :host {
    display: block;
    /* Keep in sync with HORIZONTAL_RULER_HEIGHT (the css tag rejects a numeric
       interpolation). */
    height: 20px;
    background: var(--docen-ruler-bg, #f3f3f3);
    border-top: 1px solid var(--docen-ruler-border, #e3e3e3);
    border-bottom: 1px solid var(--docen-ruler-border-strong, #c8c8c8);
    box-sizing: border-box;
    position: relative;
    user-select: none;
    font-family: var(--docen-font-family, "Segoe UI", -apple-system, sans-serif);
    font-size: 10px;
    color: #444444;
    /* The Alt-precision tooltip floats above the strip. */
    overflow: visible;
  }

  .ruler-track {
    position: relative;
    height: 100%;
    width: 100%;
    cursor: crosshair;
  }

  /* Word's margin gutter is the neutral strip; the text column is the white
     panel between the two margin hairlines. */
  .content-bg {
    position: absolute;
    top: 0;
    bottom: 0;
    background: var(--docen-ruler-content-bg, #ffffff);
    border-left: 1px solid var(--docen-ruler-margin-line, #d0d0d0);
    border-right: 1px solid var(--docen-ruler-margin-line, #d0d0d0);
    box-sizing: border-box;
  }

  /* Unit badge — Word's small CM/IN label sits at the page's start margin, not
     in a toolbar-like button. Clicking it switches the unit. */
  .unit-label {
    position: absolute;
    top: 1px;
    width: 22px;
    height: 12px;
    padding: 0;
    border: none;
    background: transparent;
    color: #5a5a5a;
    font-family: inherit;
    font-size: 8px;
    font-weight: 700;
    line-height: 12px;
    text-align: center;
    cursor: pointer;
    z-index: 6;
  }

  .unit-label:hover {
    color: #0f6cbd;
    text-decoration: underline;
  }

  .ticks-svg {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    pointer-events: none;
    z-index: 1;
  }

  /* Indent markers — Word's stacked hourglass: first-line triangle at the top,
     hanging triangle below it, left-indent square at the bottom. The right
     indent marker mirrors the hanging triangle at the end margin. */
  .marker {
    position: absolute;
    z-index: 10;
    cursor: ew-resize;
    transform: translateX(-50%);
  }

  .first-line-marker {
    top: 0;
    width: 11px;
    height: 8px;
    clip-path: polygon(0 0, 100% 0, 50% 100%);
    background: var(--docen-ruler-marker, #4a4a4a);
  }

  .hanging-marker {
    top: 7px;
    width: 11px;
    height: 8px;
    clip-path: polygon(50% 0, 0 100%, 100% 100%);
    background: var(--docen-ruler-marker, #4a4a4a);
  }

  .left-marker {
    top: 13px;
    width: 11px;
    height: 5px;
    border-radius: 1px;
    background: var(--docen-ruler-marker, #4a4a4a);
  }

  .right-marker {
    top: 9px;
    width: 11px;
    height: 9px;
    clip-path: polygon(50% 0, 0 100%, 100% 100%);
    background: var(--docen-ruler-marker, #4a4a4a);
  }

  .marker:hover,
  .marker.active {
    background: var(--docen-color-accent, #0f6cbd);
  }

  .marker.active {
    filter: brightness(0.92);
  }

  /* Tab stops */
  .tab-stop-item {
    position: absolute;
    top: 5px;
    width: 11px;
    height: 13px;
    cursor: ew-resize;
    z-index: 9;
    transform: translateX(-50%);
  }

  .tab-stop-item.drag-off {
    opacity: 0.3;
  }

  .tab-stop-item .tab-glyph {
    display: block;
    width: 11px;
    height: 13px;
    stroke: #404040;
    stroke-width: 1.4;
    fill: none;
    stroke-linecap: square;
    shape-rendering: crispEdges;
  }

  .tab-stop-item:hover .tab-glyph,
  .tab-stop-item.active .tab-glyph {
    stroke: var(--docen-color-accent, #0f6cbd);
  }

  /* Tooltip & Guideline. The strip pins to the true pane top, so the tooltip
     floats below it (Word drops it below the ruler at the top of the pane). */
  .alt-tooltip {
    position: absolute;
    top: 22px;
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
    top: 18px;
    bottom: -2000px;
    width: 1px;
    border-left: 1px dashed #0f6cbd;
    z-index: 99;
    pointer-events: none;
  }
`;

const template = html<DocenRuler>`
  <div
    class="ruler-track"
    part="track"
    @click="${(x, c) => x.onTrackClick(c.event as MouseEvent)}"
    @dblclick="${(x, c) => x.onTrackDblClick(c.event as MouseEvent)}"
  >
    <div
      class="content-bg"
      style="left: ${(x) => x.contentStartPx}px; width: ${(x) => x.contentWidthPx}px;"
    ></div>

    <svg class="ticks-svg" part="ticks" ${ref("ticksSvg")}></svg>

    <button
      type="button"
      class="unit-label"
      part="unit-toggle"
      style="left: ${(x) => x.unitLabelX}px;"
      title="${(x) => t("ruler.unitToggle", x)}"
      aria-label="${(x) => t("ruler.unitToggle", x)}"
      @click="${(x) => x.toggleUnit()}"
    >
      ${(x) => x.unit.toUpperCase()}
    </button>

    <!-- First-Line Indent Marker -->
    <div
      class="marker first-line-marker ${(x) => (x.dragActiveMarker === "firstLine" ? "active" : "")}"
      part="marker-first-line"
      title="${(x) => t("ruler.firstLine", x)}"
      style="left: ${(x) => x.firstLineMarkerX}px;"
      @pointerdown="${(x, c) => x.onMarkerPointerDown("firstLine", c.event as PointerEvent)}"
    ></div>

    <!-- Hanging Indent Marker -->
    <div
      class="marker hanging-marker ${(x) => (x.dragActiveMarker === "hanging" ? "active" : "")}"
      part="marker-hanging"
      title="${(x) => t("ruler.hanging", x)}"
      style="left: ${(x) => x.hangingMarkerX}px;"
      @pointerdown="${(x, c) => x.onMarkerPointerDown("hanging", c.event as PointerEvent)}"
    ></div>

    <!-- Left Indent Base Marker -->
    <div
      class="marker left-marker ${(x) => (x.dragActiveMarker === "left" ? "active" : "")}"
      part="marker-left"
      title="${(x) => t("ruler.leftIndent", x)}"
      style="left: ${(x) => x.hangingMarkerX}px;"
      @pointerdown="${(x, c) => x.onMarkerPointerDown("left", c.event as PointerEvent)}"
    ></div>

    <!-- Right Indent Marker -->
    <div
      class="marker right-marker ${(x) => (x.dragActiveMarker === "right" ? "active" : "")}"
      part="marker-right"
      title="${(x) => t("ruler.rightIndent", x)}"
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
`;

/**
 * `<docen-ruler>` — Word's horizontal ruler: a 20px page-aligned strip with
 * the four-level tick hierarchy, margin gutter, draggable indent markers
 * (first-line / hanging / left / right), typed tab-stop glyphs, and 2-way
 * ProseMirror sync. The container spans the page exactly (the host sets its
 * width to the zoomed page width and mounts it above the page column); the
 * unit label overlays the start margin instead of displacing the track.
 */
@customElement({ name: "docen-ruler", template, styles })
export class DocenRuler extends FASTElement {
  @attr unit: "in" | "cm" = "in";
  @attr dir: "ltr" | "rtl" = "ltr";
  @observable pageWidthPx = 816; // 8.5" * 96
  @observable marginLeftPx = 96; // 1" * 96
  @observable marginRightPx = 96; // 1" * 96
  @observable scale = 1;
  /** The page's left edge in unzoomed strip px — the band spans the whole
   *  document pane, so every page-relative coordinate is offset by this. The
   *  host measures it from the centered page column (0 before the first
   *  layout, when the band is exactly the page width). */
  @observable pageLeftPx = 0;

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
      // keep an explicitly-set dir attribute
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
  @attr({ attribute: "active-tab-type" })
  @observable
  activeTabType: TabSelectorType = "left";

  // Drag interaction state
  @observable dragActiveMarker: "firstLine" | "hanging" | "left" | "right" | "tabStop" | null =
    null;
  @observable dragTabIdx = -1;
  /** The stop object under the pointer — repeat reuses DOM nodes, so the
   *  child scope's index can go stale after an array edit; identity survives. */
  @observable dragStop: RulerTabStop | null = null;
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

  pageLeftPxChanged(): void {
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

  /** The auto-added tab stop from the last single click — a double-click on
   *  the ruler removes it before opening the Tabs dialog (Word's double-click
   *  is a dialog gesture, not two tab-stop creations). */
  #lastAutoStop: { position: number; at: number } | null = null;
  #liveFrame = 0;

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
    this.#cancelLiveFrame();
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

  /** The page's left edge in zoomed strip px. */
  get pageLeftPxViewport(): number {
    return this.pageLeftPx * this.scale;
  }

  /** The content-origin (margin line) in strip px — Word's 0. */
  get zeroXPx(): number {
    if (this.isRtl) {
      return this.pageLeftPxViewport + (this.pageWidthPx - this.marginRightPx) * this.scale;
    }
    return this.pageLeftPxViewport + this.marginLeftPx * this.scale;
  }

  get contentWidthPx(): number {
    return (this.pageWidthPx - this.marginLeftPx - this.marginRightPx) * this.scale;
  }

  /** The content panel's physical screen start — margins stay physical in RTL;
   *  only the tick scale and the markers mirror around the right margin line. */
  get contentStartPx(): number {
    return this.pageLeftPxViewport + this.marginLeftPx * this.scale;
  }

  /** The CM/IN badge sits at the page's start margin (RTL: end margin). */
  get unitLabelX(): number {
    if (this.isRtl) {
      return this.pageLeftPxViewport + this.pageWidthPx * this.scale - 24;
    }
    return this.pageLeftPxViewport + 2;
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
      return this.contentStartPx + px;
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
    const pageZeroX = this.isRtl
      ? (this.pageWidthPx - this.marginRightPx) * this.scale
      : this.marginLeftPx * this.scale;
    const ticks = rulerTicks({
      lengthPx: this.pageWidthPx * this.scale,
      zeroPx: pageZeroX,
      unit: this.unit,
      scale: this.scale,
      mirror: this.isRtl,
    });
    // The strip's content box is 18px (20px minus the two hairlines); ticks
    // grow upward from its bottom edge like Word's.
    const height = 18;
    let lines = "";
    let texts = "";
    for (const tick of ticks) {
      const x = Math.round(this.pageLeftPxViewport + tick.pos) + 0.5;
      const len = RULER_TICK_LEN[tick.level];
      lines += `<line x1="${x}" y1="${height}" x2="${x}" y2="${height - len}"/>`;
      if (tick.label !== undefined) {
        texts += `<text x="${x}" y="8" text-anchor="middle">${tick.label}</text>`;
      }
    }
    svg.setAttribute("shape-rendering", "crispEdges");
    svg.innerHTML =
      `<g stroke="#8f8f8f" stroke-width="1" fill="none">${lines}</g>` +
      `<g fill="#555555" font-size="7" font-family="inherit">${texts}</g>`;
  }

  // ── Tab Stops Rendering ──

  renderTabStops(): ReturnType<typeof html> {
    return html`
      ${repeat(
        () => this.tabStops,
        html<RulerTabStop, DocenRuler>`
          <div
            class="tab-stop-item ${(stop, c) =>
              c.parent.dragActiveMarker === "tabStop" &&
              c.parent.dragStop === stop &&
              c.parent.dragOffRuler
                ? "drag-off"
                : ""}"
            style="left: ${(stop, c) =>
              c.parent.isRtl
                ? c.parent.zeroXPx - c.parent.twipsToContentPx(stop.position)
                : c.parent.zeroXPx + c.parent.twipsToContentPx(stop.position)}px;"
            title="${(stop, c) =>
              `${t("ruler.tabStop", c.parent)}: ${c.parent.formatMeasurement(stop.position)} (${stop.type})`}"
            aria-label="${(stop, c) =>
              `${t("ruler.tabStop", c.parent)}: ${c.parent.formatMeasurement(stop.position)}`}"
            @pointerdown="${(stop, c) => c.parent.onTabPointerDown(stop, c.event as PointerEvent)}"
            @click="${(_stop, c) => c.event.stopPropagation()}"
            @dblclick="${(stop, c) => c.parent.onTabDblClick(stop, c.event as MouseEvent)}"
          >
            <svg class="tab-glyph" viewBox="0 0 9 10" aria-hidden="true">
              <path d="${(stop) => tabGlyphPath(stop.type)}"></path>
              ${(stop) =>
                stop.type === "decimal"
                  ? html`<circle cx="6.5" cy="5" r="1.2" fill="#404040" stroke="none"></circle>`
                  : ""}
            </svg>
          </div>
        `,
      )}
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
      // Word reflows the paragraph under the dragged marker live.
      this.#queueLiveCommit();
    };

    const onPointerUp = (_e: PointerEvent): void => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
      this.dragActiveMarker = null;
      this.showTooltip = false;
      this.#cancelLiveFrame();
      this.commitIndentToEditor();
    };

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  onTabPointerDown(ref: number | RulerTabStop, event: PointerEvent): void {
    const idx = typeof ref === "number" ? ref : this.tabStops.indexOf(ref);
    const stop = this.tabStops[idx];
    if (!stop) return;
    event.preventDefault();
    event.stopPropagation();
    (event.target as HTMLElement).setPointerCapture(event.pointerId);

    this.dragActiveMarker = "tabStop";
    this.dragTabIdx = idx;
    this.dragStop = stop;
    this.#dragStartX = event.clientX;
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
        // The dragged object was replaced — keep the identity highlight fresh.
        this.dragStop = this.tabStops[idx] ?? null;
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
      this.dragStop = null;
      this.dragOffRuler = false;
      this.showTooltip = false;
      this.commitTabStopsToEditor();
    };

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  /** `ref` is the stop object from the repeat's child scope (the scoped index
   *  can be stale after an array edit); a raw index stays accepted. */
  onTabDblClick(ref: number | RulerTabStop, event: MouseEvent): void {
    event.stopPropagation();
    const stop = typeof ref === "number" ? this.tabStops[ref] : ref;
    this.$emit("ruler:open-tabs", { tabStop: stop });
    this.#dispatchTabsDialog();
  }

  /** Double-clicking the ruler's open strip is Word's dialog gesture: drop the
   *  tab stop the preceding single click just auto-added, then open Tabs. */
  onTrackDblClick(event: MouseEvent): void {
    const target = event.target as Element | null;
    if (target?.closest?.(".marker, .tab-stop-item, .unit-label")) return;
    const track = event.currentTarget as HTMLElement | null;
    if (track) {
      const rect = track.getBoundingClientRect();
      const clickX = event.clientX - rect.left;
      if (
        clickX < this.pageLeftPxViewport ||
        clickX > this.pageLeftPxViewport + this.pageWidthPx * this.scale
      ) {
        return;
      }
    }
    const auto = this.#lastAutoStop;
    if (auto && Date.now() - auto.at < 700) {
      this.tabStops = this.tabStops.filter((s) => s.position !== auto.position);
      this.#lastAutoStop = null;
      this.commitTabStopsToEditor();
    }
    this.$emit("ruler:open-tabs", {});
    this.#dispatchTabsDialog();
  }

  onTrackClick(event: MouseEvent): void {
    const target = event.target as Element | null;
    if (target?.closest?.(".marker, .tab-stop-item, .unit-label")) return;
    if (event.detail > 1) return; // second click of a double-click: dialog gesture
    const track = event.currentTarget as HTMLElement | null;
    if (!track) return;
    const rect = track.getBoundingClientRect();
    const clickX = event.clientX - rect.left;

    // Check if clicked inside content zone (RTL: zero sits at the right line)
    const contentLeft = this.isRtl ? this.zeroXPx - this.contentWidthPx : this.zeroXPx;
    const contentRight = contentLeft + this.contentWidthPx;

    if (clickX >= contentLeft && clickX <= contentRight) {
      const offsetPx = this.isRtl ? contentRight - clickX : clickX - contentLeft;
      const twips = this.contentPxToTwips(offsetPx);

      if (this.activeTabType === "first-line") {
        this.firstLineTwips = twips - this.leftIndentTwips;
        this.#emitIndentChange();
        this.#queueLiveCommit();
        return;
      }
      if (this.activeTabType === "hanging") {
        this.leftIndentTwips = twips;
        this.#emitIndentChange();
        this.#queueLiveCommit();
        return;
      }

      // Word never stacks two stops on the same spot — an existing stop wins.
      if (this.tabStops.some((s) => Math.abs(s.position - twips) < 30)) return;
      const stopType: RulerTabStop["type"] =
        this.activeTabType === "bar" ||
        this.activeTabType === "center" ||
        this.activeTabType === "right" ||
        this.activeTabType === "decimal"
          ? this.activeTabType
          : "left";
      const newStop: RulerTabStop = { position: twips, type: stopType };
      this.tabStops = [...this.tabStops, newStop].sort((a, b) => a.position - b.position);
      this.#lastAutoStop = { position: twips, at: Date.now() };
      this.$emit("ruler:tabstop-add", { tabStop: newStop });
      this.commitTabStopsToEditor();
    }
  }

  #updateTooltip(clientX: number): void {
    const track = this.shadowRoot?.querySelector(".ruler-track")?.getBoundingClientRect();
    if (!track) return;
    this.tooltipX = clientX - track.left;

    if (this.dragActiveMarker === "firstLine") {
      this.tooltipText = `${t("ruler.firstLine", this)}: ${this.formatMeasurement(this.firstLineTwips)}`;
    } else if (this.dragActiveMarker === "hanging" || this.dragActiveMarker === "left") {
      this.tooltipText = `${t("ruler.leftIndent", this)}: ${this.formatMeasurement(this.leftIndentTwips)}`;
    } else if (this.dragActiveMarker === "right") {
      this.tooltipText = `${t("ruler.rightIndent", this)}: ${this.formatMeasurement(this.rightIndentTwips)}`;
    } else if (this.dragActiveMarker === "tabStop" && this.dragTabIdx >= 0) {
      const stop = this.tabStops[this.dragTabIdx];
      this.tooltipText = stop
        ? `${t("ruler.tabStop", this)}: ${this.formatMeasurement(stop.position)} (${stop.type})`
        : "";
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

  /** The paragraphs a ruler edit lands on — every textblock the selection
   *  touches (Word applies ruler changes to the whole selection), falling
   *  back to the caret's own paragraph for a collapsed selection. */
  #paragraphTargets(): Array<{ pos: number; attrs: Record<string, unknown> }> {
    const editor = this.#editor;
    if (!editor) return [];
    const { state } = editor;
    const { from, to } = state.selection;
    const found = new Map<number, Record<string, unknown>>();
    state.doc.nodesBetween(from, to, (node, pos) => {
      const name = node.type.name;
      if (name === "paragraph" || name === "heading") {
        found.set(pos, node.attrs as Record<string, unknown>);
        return false;
      }
      return true;
    });
    if (found.size === 0) {
      const { $from } = state.selection;
      for (let d = $from.depth; d > 0; d--) {
        const node = $from.node(d);
        if (node.type.name === "paragraph" || node.type.name === "heading") {
          found.set($from.before(d), node.attrs as Record<string, unknown>);
          break;
        }
      }
    }
    return [...found].map(([pos, attrs]) => ({ pos, attrs }));
  }

  #indentAttrs(): Record<string, number> | null {
    const indent: Record<string, number> = {};
    if (this.leftIndentTwips) indent.left = this.leftIndentTwips;
    if (this.rightIndentTwips) indent.right = this.rightIndentTwips;
    if (this.firstLineTwips > 0) indent.firstLine = this.firstLineTwips;
    if (this.firstLineTwips < 0) indent.hanging = -this.firstLineTwips;
    return Object.keys(indent).length > 0 ? indent : null;
  }

  #queueLiveCommit(): void {
    if (!this.#editor || this.#liveFrame) return;
    const run = (): void => {
      this.#liveFrame = 0;
      this.commitIndentToEditor();
    };
    this.#liveFrame =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame(run)
        : (setTimeout(run, 16) as unknown as number);
  }

  #cancelLiveFrame(): void {
    if (!this.#liveFrame) return;
    if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.#liveFrame);
    else clearTimeout(this.#liveFrame);
    this.#liveFrame = 0;
  }

  commitIndentToEditor(): void {
    this.#emitIndentChange();
    const editor = this.#editor;
    if (!editor) return;
    const targets = this.#paragraphTargets();
    if (targets.length === 0) return;
    const indent = this.#indentAttrs();
    const { tr } = editor.state;
    for (const { pos, attrs } of targets) {
      tr.setNodeMarkup(pos, undefined, {
        ...attrs,
        indent: indent ? { ...indent } : null,
      });
    }
    editor.view.dispatch(tr);
  }

  commitTabStopsToEditor(): void {
    this.$emit("ruler:tabstop-change", { tabStops: this.tabStops });
    const editor = this.#editor;
    if (!editor) return;
    const targets = this.#paragraphTargets();
    if (targets.length === 0) return;
    const { tr } = editor.state;
    for (const { pos, attrs } of targets) {
      tr.setNodeMarkup(pos, undefined, {
        ...attrs,
        tabStops: this.tabStops.length > 0 ? [...this.tabStops] : null,
      });
    }
    editor.view.dispatch(tr);
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
      pageLeftPx?: number;
    },
  ): void {
    if (geometry) {
      if (geometry.pageWidthPx != null) this.pageWidthPx = geometry.pageWidthPx;
      if (geometry.marginLeftPx != null) this.marginLeftPx = geometry.marginLeftPx;
      if (geometry.marginRightPx != null) this.marginRightPx = geometry.marginRightPx;
      if (geometry.scale != null) this.scale = geometry.scale;
      if (geometry.pageLeftPx != null) this.pageLeftPx = geometry.pageLeftPx;
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
