import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";
import { RULER_TICK_LEN, rulerTicks, type RulerUnit } from "./ruler-ticks";

/** The strip's thickness in CSS px — Word's vertical ruler is 20px wide. */
export const VERTICAL_RULER_THICKNESS = 20;

const styles = css`
  :host {
    display: block;
    position: fixed;
    /* Keep in sync with VERTICAL_RULER_THICKNESS (the css tag rejects a
       numeric interpolation). */
    width: 20px;
    box-sizing: border-box;
    background: var(--docen-ruler-bg, #f3f3f3);
    /* The horizontal ruler's bottom border is the pane seam — no top border
       (the corner stays flush when both rulers are shown). */
    border: 1px solid var(--docen-ruler-border-strong, #c8c8c8);
    border-top: none;
    /* The page column slides under the fixed strip — nothing may paint past
       its edges. */
    overflow: hidden;
    z-index: 6;
    user-select: none;
    font-family: var(--docen-font-family, "Segoe UI", -apple-system, sans-serif);
    font-size: 10px;
    color: #444444;
  }

  .vr-track {
    position: relative;
    width: 100%;
    height: 100%;
  }

  /* Word's ruler shows the text column as the white band between the two
     margin lines; the margins stay neutral. */
  .content-bg {
    position: absolute;
    left: 0;
    right: 0;
    background: var(--docen-ruler-content-bg, #ffffff);
    border-top: 1px solid var(--docen-ruler-margin-line, #d0d0d0);
    border-bottom: 1px solid var(--docen-ruler-margin-line, #d0d0d0);
    box-sizing: border-box;
    pointer-events: none;
    z-index: 0;
  }

  .ticks-svg {
    position: absolute;
    inset: 0;
    width: 100%;
    height: 100%;
    pointer-events: none;
    z-index: 1;
  }

  /* Word's margin boundaries: a small triangle pointing into the page at the
     content-top and content-bottom lines. Both drag vertically (ns-resize). */
  .margin-handle {
    position: absolute;
    right: 0;
    width: 10px;
    height: 8px;
    transform: translateY(-50%);
    clip-path: polygon(0 0, 100% 50%, 0 100%);
    background: var(--docen-ruler-marker, #4a4a4a);
    cursor: ns-resize;
    z-index: 10;
  }

  .margin-handle:hover,
  .margin-handle.active {
    background: var(--docen-color-accent, #0f6cbd);
  }

  .margin-handle[hidden] {
    display: none;
  }

  .alt-tooltip {
    position: absolute;
    right: 24px;
    transform: translateY(-50%);
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
`;

const template = html<DocenVerticalRuler>`
  <div class="vr-track" part="track">
    <div
      class="content-bg"
      part="content-bg"
      style="top: ${(x) => x.contentTopYPx}px; height: ${(x) => x.contentHeightPx}px;"
    ></div>

    <svg class="ticks-svg" part="ticks" ${ref("ticksSvg")}></svg>

    <div
      class="margin-handle top ${(x) => (x.dragSide === "top" ? "active" : "")}"
      part="margin-top"
      title="${(x) => t("ruler.topMargin", x)}"
      aria-label="${(x) => t("ruler.topMargin", x)}"
      style="top: ${(x) => x.topHandleYPx}px;"
      ?hidden="${(x) => !x.topHandleVisible}"
      @pointerdown="${(x, c) => x.onHandlePointerDown("top", c.event as PointerEvent)}"
    ></div>

    <div
      class="margin-handle bottom ${(x) => (x.dragSide === "bottom" ? "active" : "")}"
      part="margin-bottom"
      title="${(x) => t("ruler.bottomMargin", x)}"
      aria-label="${(x) => t("ruler.bottomMargin", x)}"
      style="top: ${(x) => x.bottomHandleYPx}px;"
      ?hidden="${(x) => !x.bottomHandleVisible}"
      @pointerdown="${(x, c) => x.onHandlePointerDown("bottom", c.event as PointerEvent)}"
    ></div>

    ${(x) =>
      x.showTooltip && x.tooltipText
        ? html`<div class="alt-tooltip" style="top: ${x.tooltipYPx}px;">${x.tooltipText}</div>`
        : ""}
  </div>
`;

/**
 * `<docen-vertical-ruler>` — Word's fixed vertical ruler: a 20px strip pinned
 * to the document pane's left edge (left also for RTL, like Word), showing the
 * current page's vertical scale from the content-top origin. The top/bottom
 * margin boundaries carry draggable handles (the only interactive parts — no
 * indent/tab markers); dragging emits live `v-ruler:margin-input` events and
 * commits on release. The host owns visibility (Print Layout + View → Ruler +
 * the vertical-ruler option) and the section the events land on.
 *
 * Geometry is viewport-relative: `originY` is the current page's top edge
 * relative to the strip's top (negative once the page has scrolled up), so the
 * scale slides under the fixed strip exactly like Word's. All px values are
 * already zoom-scaled; `scale` is only the twips↔px conversion factor.
 */
@customElement({ name: "docen-vertical-ruler", template, styles })
export class DocenVerticalRuler extends FASTElement {
  @attr unit: RulerUnit = "in";
  /** The current page's top edge relative to the strip top (px, may be < 0). */
  @observable originY = 0;
  /** Page height in zoomed px. */
  @observable pageHeightPx = 1056;
  /** Distance from the page's top edge to the content origin (zoomed px). */
  @observable contentTopPx = 96;
  /** The text column's height (zoomed px). */
  @observable contentHeightPx = 864;
  /** Zoom factor (1 = 100%) — used for twips conversion only. */
  @observable scale = 1;

  @observable dragSide: "top" | "bottom" | null = null;
  @observable dragValueTwips: number | null = null;
  @observable showTooltip = false;
  @observable tooltipText = "";
  @observable tooltipYPx = 0;
  /** The strip's own height, tracked so handle visibility and tick coverage
   *  follow pane resizes. */
  @observable hostHeight = 0;
  @observable ticksSvg?: SVGSVGElement;

  #unobserveLang?: () => void;
  #resizeObserver?: ResizeObserver;
  #dragStartY = 0;
  #dragStartTwips = 0;
  #dragStartHandleY = 0;
  #liveFrame = 0;

  ticksSvgChanged(): void {
    this.renderTicks();
  }

  unitChanged(): void {
    this.renderTicks();
  }

  originYChanged(): void {
    this.renderTicks();
  }

  scaleChanged(): void {
    this.renderTicks();
  }

  contentTopPxChanged(): void {
    this.renderTicks();
  }

  contentHeightPxChanged(): void {
    this.renderTicks();
  }

  pageHeightPxChanged(): void {
    this.renderTicks();
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unobserveLang = observeLang(() => {
      this.renderTicks();
    });
    this.#resizeObserver = new ResizeObserver(() => {
      this.hostHeight = this.clientHeight || this.getBoundingClientRect().height || 0;
      this.renderTicks();
    });
    this.#resizeObserver.observe(this);
    this.hostHeight = this.clientHeight || this.getBoundingClientRect().height || 0;
    this.renderTicks();
  }

  override disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    this.#resizeObserver?.disconnect();
    this.#resizeObserver = undefined;
    this.#cancelLiveFrame();
    super.disconnectedCallback();
  }

  override $emit(
    type: string,
    detail?: unknown,
    options?: Omit<CustomEventInit, "detail">,
  ): boolean {
    return this.dispatchEvent(
      new CustomEvent(type, { bubbles: true, composed: true, detail, ...options }),
    );
  }

  // ── Margin geometry (twips) ──

  /** Twips per zoomed px. */
  get #twipsPerPx(): number {
    return this.scale > 0 ? 15 / this.scale : 15;
  }

  get topMarginTwips(): number {
    return Math.round(this.contentTopPx * this.#twipsPerPx);
  }

  get bottomMarginTwips(): number {
    return Math.round(
      (this.pageHeightPx - this.contentTopPx - this.contentHeightPx) * this.#twipsPerPx,
    );
  }

  get pageHeightTwips(): number {
    return Math.round(this.pageHeightPx * this.#twipsPerPx);
  }

  get contentHeightTwips(): number {
    return Math.round(this.contentHeightPx * this.#twipsPerPx);
  }

  /** The largest top margin that leaves the text column on the page. */
  get maxTopMarginTwips(): number {
    return Math.max(0, this.pageHeightTwips - this.contentHeightTwips);
  }

  /** The largest bottom margin that leaves the text column on the page. */
  get maxBottomMarginTwips(): number {
    return Math.max(0, this.pageHeightTwips - this.topMarginTwips);
  }

  get #hostHeight(): number {
    return this.hostHeight || this.clientHeight || this.getBoundingClientRect().height || 0;
  }

  /** The dragged handle tracks the pointer; the resting one tracks layout. */
  get topHandleYPx(): number {
    if (this.dragSide === "top" && this.dragValueTwips != null) {
      return (
        this.#dragStartHandleY + ((this.dragValueTwips - this.#dragStartTwips) * this.scale) / 15
      );
    }
    return this.originY + this.contentTopPx;
  }

  get bottomHandleYPx(): number {
    if (this.dragSide === "bottom" && this.dragValueTwips != null) {
      return (
        this.#dragStartHandleY - ((this.dragValueTwips - this.#dragStartTwips) * this.scale) / 15
      );
    }
    return this.originY + this.contentTopPx + this.contentHeightPx;
  }

  get topHandleVisible(): boolean {
    return this.#handleVisible(this.topHandleYPx);
  }

  get bottomHandleVisible(): boolean {
    return this.#handleVisible(this.bottomHandleYPx);
  }

  #handleVisible(y: number): boolean {
    const h = this.#hostHeight;
    return y >= 0 && y <= h;
  }

  get contentTopYPx(): number {
    return this.originY + this.contentTopPx;
  }

  // ── Pointer drag (Word's draggable margin boundaries) ──

  onHandlePointerDown(side: "top" | "bottom", event: PointerEvent): void {
    event.preventDefault();
    event.stopPropagation();
    (event.target as HTMLElement).setPointerCapture(event.pointerId);

    this.dragSide = side;
    this.#dragStartY = event.clientY;
    this.#dragStartTwips = side === "top" ? this.topMarginTwips : this.bottomMarginTwips;
    this.#dragStartHandleY =
      side === "top"
        ? this.originY + this.contentTopPx
        : this.originY + this.contentTopPx + this.contentHeightPx;
    this.dragValueTwips = this.#dragStartTwips;
    this.showTooltip = true;
    this.#updateTooltip();
    this.$emit("v-ruler:margin-start", { side });

    const onPointerMove = (e: PointerEvent): void => {
      const deltaTwips = Math.round(((e.clientY - this.#dragStartY) / this.scale) * 15);
      // Top margin grows downward; the bottom boundary grows upward (a larger
      // bottom margin moves the content-bottom line up).
      const raw =
        side === "top" ? this.#dragStartTwips + deltaTwips : this.#dragStartTwips - deltaTwips;
      const max = side === "top" ? this.maxTopMarginTwips : this.maxBottomMarginTwips;
      this.dragValueTwips = Math.max(0, Math.min(max, raw));
      this.#updateTooltip();
      this.#queueLiveCommit();
    };

    const onPointerUp = (): void => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
      const twips = this.dragValueTwips ?? this.#dragStartTwips;
      this.#cancelLiveFrame();
      this.$emit("v-ruler:margin-commit", { side, twips });
      this.dragSide = null;
      this.dragValueTwips = null;
      this.showTooltip = false;
    };

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  #updateTooltip(): void {
    const side = this.dragSide;
    if (!side || this.dragValueTwips == null) return;
    const label = t(side === "top" ? "ruler.topMargin" : "ruler.bottomMargin", this);
    this.tooltipText = `${label}: ${this.formatMeasurement(this.dragValueTwips)}`;
    const y = side === "top" ? this.topHandleYPx : this.bottomHandleYPx;
    const h = this.#hostHeight;
    this.tooltipYPx = Math.max(10, Math.min(h - 10, y));
  }

  #queueLiveCommit(): void {
    if (this.#liveFrame) return;
    const run = (): void => {
      this.#liveFrame = 0;
      if (this.dragSide && this.dragValueTwips != null) {
        this.$emit("v-ruler:margin-input", { side: this.dragSide, twips: this.dragValueTwips });
      }
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

  formatMeasurement(twips: number): string {
    if (this.unit === "cm") {
      const cm = (twips / 1440) * 2.54;
      return `${cm.toFixed(2)} cm`;
    }
    const inches = twips / 1440;
    return `${inches.toFixed(2)} in`;
  }

  // ── SVG ticks (the shared four-level hierarchy) ──

  /** Ticks span the full pane height (Word's continuous ruler) with 0 at the
   *  content-top origin; the scale slides under the fixed strip as the page
   *  scrolls. */
  renderTicks(): void {
    const svg = this.ticksSvg;
    if (!svg) return;
    const height = this.#hostHeight;
    if (!(height > 0)) return;
    const ticks = rulerTicks({
      lengthPx: height,
      zeroPx: this.originY,
      unit: this.unit,
      scale: this.scale,
    });
    let lines = "";
    let texts = "";
    for (const tick of ticks) {
      const y = Math.round(tick.pos) + 0.5;
      const len = RULER_TICK_LEN[tick.level];
      lines += `<line x1="${VERTICAL_RULER_THICKNESS}" y1="${y}" x2="${VERTICAL_RULER_THICKNESS - len}" y2="${y}"/>`;
      if (tick.label !== undefined) {
        const tx = VERTICAL_RULER_THICKNESS - RULER_TICK_LEN[0] - 4;
        texts += `<text x="${tx}" y="${y}" text-anchor="middle" transform="rotate(-90 ${tx} ${y})">${tick.label}</text>`;
      }
    }
    svg.setAttribute("shape-rendering", "crispEdges");
    svg.innerHTML =
      `<g stroke="#8f8f8f" stroke-width="1" fill="none">${lines}</g>` +
      `<g fill="#555555" font-size="7" font-family="inherit">${texts}</g>`;
  }
}

export default DocenVerticalRuler;
