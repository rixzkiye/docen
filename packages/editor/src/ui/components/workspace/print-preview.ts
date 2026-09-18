import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { pagesToPdf, type PdfPageShot } from "../../../document/export-pdf";
import { observeLang, t } from "../../i18n/localize";

export type PageRangeType = "all" | "current" | "selection" | "custom";
export type OrientationPreset = "auto" | "portrait" | "landscape";
export type PaperSizePreset = "auto" | "letter" | "a4" | "legal" | "a3" | "a5";
export type LayoutPreset = "1-up" | "booklet";
export type PrintMarkupMode = "document" | "markup";

export interface BookletSheet {
  sheetIndex: number;
  side: "front" | "back";
  leftPage: number | null; // 1-based page number or null
  rightPage: number | null; // 1-based page number or null
}

export interface PaperDimensions {
  width: number;
  height: number;
}

export const PAPER_SIZES: Record<
  Exclude<PaperSizePreset, "auto">,
  { width: number; height: number }
> = {
  letter: { width: 816, height: 1056 },
  a4: { width: 793.7, height: 1122.5 },
  legal: { width: 816, height: 1344 },
  a3: { width: 1122.5, height: 1587.4 },
  a5: { width: 559.4, height: 793.7 },
};

/**
 * Parse a page range specification into a list of 1-based page numbers.
 */
export function parsePageRange(options: {
  type: PageRangeType;
  custom?: string;
  currentPage?: number;
  selectionPages?: number[];
  totalPages: number;
}): number[] {
  const { type, custom, currentPage = 1, selectionPages = [], totalPages } = options;
  if (totalPages <= 0) return [];

  if (type === "all") {
    return Array.from({ length: totalPages }, (_, i) => i + 1);
  }

  if (type === "current") {
    const p = Math.min(Math.max(1, currentPage), totalPages);
    return [p];
  }

  if (type === "selection") {
    const valid = Array.from(new Set(selectionPages.filter((p) => p >= 1 && p <= totalPages))).sort(
      (a, b) => a - b,
    );
    if (valid.length > 0) return valid;
    const p = Math.min(Math.max(1, currentPage), totalPages);
    return [p];
  }

  // Custom range e.g. "1-3, 5, 7-10"
  if (!custom || !custom.trim()) {
    return Array.from({ length: totalPages }, (_, i) => i + 1);
  }

  const result: number[] = [];
  const segments = custom.split(",");
  for (const seg of segments) {
    const s = seg.trim();
    if (!s) continue;
    const rangeMatch = s.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const start = Number.parseInt(rangeMatch[1]!, 10);
      const end = Number.parseInt(rangeMatch[2]!, 10);
      const min = Math.max(1, Math.min(start, end));
      const max = Math.min(totalPages, Math.max(start, end));
      for (let i = min; i <= max; i++) {
        if (!result.includes(i)) result.push(i);
      }
      continue;
    }

    const singleMatch = s.match(/^(\d+)$/);
    if (singleMatch) {
      const p = Number.parseInt(singleMatch[1]!, 10);
      if (p >= 1 && p <= totalPages && !result.includes(p)) {
        result.push(p);
      }
    }
  }

  return result.length > 0 ? result : Array.from({ length: totalPages }, (_, i) => i + 1);
}

/**
 * Expand a list of pages with copies count and collation.
 */
export function expandCopies(pages: number[], copies: number, collate: boolean): number[] {
  const count = Math.max(1, Math.floor(copies));
  if (pages.length === 0) return [];
  if (count === 1) return [...pages];

  const result: number[] = [];
  if (collate) {
    // 1, 2, 3, 1, 2, 3...
    for (let c = 0; c < count; c++) {
      result.push(...pages);
    }
  } else {
    // 1, 1, 2, 2, 3, 3...
    for (const p of pages) {
      for (let c = 0; c < count; c++) {
        result.push(p);
      }
    }
  }
  return result;
}

/**
 * 2-up booklet imposition algorithm.
 * Groups pages into 4-page folded sheets with proper front and back ordering.
 */
export function calculateBookletImposition(pageCount: number): BookletSheet[] {
  if (pageCount <= 0) return [];

  const totalSheets = Math.ceil(pageCount / 4);
  const totalPages = totalSheets * 4;
  const sheets: BookletSheet[] = [];

  for (let s = 0; s < totalSheets; s++) {
    // Front side of sheet s
    const frontLeft = totalPages - 2 * s;
    const frontRight = 2 * s + 1;
    sheets.push({
      sheetIndex: s,
      side: "front",
      leftPage: frontLeft <= pageCount ? frontLeft : null,
      rightPage: frontRight <= pageCount ? frontRight : null,
    });

    // Back side of sheet s
    const backLeft = 2 * s + 2;
    const backRight = totalPages - 2 * s - 1;
    sheets.push({
      sheetIndex: s,
      side: "back",
      leftPage: backLeft <= pageCount ? backLeft : null,
      rightPage: backRight <= pageCount ? backRight : null,
    });
  }

  return sheets;
}

/**
 * Calculate dimensions taking into account paper size preset and orientation override.
 */
export function resolvePaperDimensions(
  preset: PaperSizePreset,
  orientation: OrientationPreset,
  baseWidth: number,
  baseHeight: number,
): PaperDimensions {
  let w = baseWidth;
  let h = baseHeight;

  if (preset !== "auto" && PAPER_SIZES[preset]) {
    w = PAPER_SIZES[preset].width;
    h = PAPER_SIZES[preset].height;
  }

  if (orientation === "portrait" && w > h) {
    const tmp = w;
    w = h;
    h = tmp;
  } else if (orientation === "landscape" && h > w) {
    const tmp = w;
    w = h;
    h = tmp;
  }

  return { width: w, height: h };
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(1080px, 96vw);
    height: min(780px, 94vh);
    max-height: 94vh;
    padding: 0;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }
  .pp-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 16px;
    border-bottom: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    background: var(--colorNeutralBackground1, #ffffff);
  }
  .pp-title {
    font-size: 16px;
    font-weight: 600;
  }
  .pp-actions {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .pp-main {
    display: flex;
    flex: 1;
    min-height: 0;
    overflow: hidden;
  }
  .pp-sidebar {
    width: 310px;
    flex: none;
    border-right: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    background: var(--colorNeutralBackground2, #fafafa);
    overflow-y: auto;
    padding: 16px;
    display: flex;
    flex-direction: column;
    gap: 14px;
    box-sizing: border-box;
  }
  .pp-field {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .pp-label {
    font-size: 12px;
    font-weight: 600;
    color: var(--colorNeutralForeground2, #616161);
  }
  .pp-row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .pp-viewport {
    flex: 1;
    background: #d8d8d8;
    overflow: auto;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 24px;
    box-sizing: border-box;
    position: relative;
  }
  .pp-stage-wrap {
    display: flex;
    align-items: center;
    justify-content: center;
    transition: transform 0.15s ease-out;
    transform-origin: center center;
  }
  .pp-page-shadow {
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
    background: #ffffff;
    display: block;
  }
  .pp-sheet-wrap {
    display: flex;
    gap: 2px;
    background: #e0e0e0;
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
  }
  .pp-sheet-half {
    background: #ffffff;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
  }
  .pp-sheet-blank {
    color: #999999;
    font-size: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .pp-controls {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 8px 16px;
    border-top: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    background: var(--colorNeutralBackground1, #ffffff);
    justify-content: space-between;
  }
  .pp-nav,
  .pp-zoom {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .pp-indicator {
    font-size: 13px;
    min-width: 100px;
    text-align: center;
  }
  fluent-dropdown,
  fluent-text-input {
    width: 100%;
  }
`;

const template = html<DocenPrintPreview>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="pp-header">
      <div class="pp-title" ${ref("titleEl")}></div>
      <div class="pp-actions">
        <fluent-button
          appearance="accent"
          ${ref("printBtn")}
          @click="${(x) => x.onPrint()}"
        ></fluent-button>
        <fluent-button
          appearance="neutral"
          ${ref("exportPdfBtn")}
          @click="${(x) => x.onExportPdf()}"
        ></fluent-button>
        <fluent-button
          appearance="stealth"
          ${ref("closeBtn")}
          @click="${(x) => x.hide()}"
        ></fluent-button>
      </div>
    </div>

    <div class="pp-main">
      <div class="pp-sidebar">
        <!-- Copies and Collation -->
        <div class="pp-field">
          <label class="pp-label" ${ref("copiesLabelEl")}></label>
          <fluent-text-input
            ${ref("copiesInput")}
            type="number"
            min="1"
            value="1"
            @change="${(x) => x.onCopiesChange()}"
          ></fluent-text-input>
        </div>

        <div class="pp-field">
          <label class="pp-label" ${ref("collateLabelEl")}></label>
          <fluent-dropdown ${ref("collateDropdown")} @change="${(x) => x.onCollateChange()}">
            <fluent-option value="collated" selected ${ref("collatedOptEl")}></fluent-option>
            <fluent-option value="uncollated" ${ref("uncollatedOptEl")}></fluent-option>
          </fluent-dropdown>
        </div>

        <!-- Page Range -->
        <div class="pp-field">
          <label class="pp-label" ${ref("pageRangeLabelEl")}></label>
          <fluent-dropdown ${ref("rangeDropdown")} @change="${(x) => x.onRangeTypeChange()}">
            <fluent-option value="all" selected ${ref("rangeAllOptEl")}></fluent-option>
            <fluent-option value="current" ${ref("rangeCurrentOptEl")}></fluent-option>
            <fluent-option value="selection" ${ref("rangeSelectionOptEl")}></fluent-option>
            <fluent-option value="custom" ${ref("rangeCustomOptEl")}></fluent-option>
          </fluent-dropdown>
          <fluent-text-input
            ${ref("customRangeInput")}
            style="margin-top: 4px;"
            placeholder="e.g. 1-3, 5, 7-10"
            @input="${(x) => x.onCustomRangeInput()}"
          ></fluent-text-input>
        </div>

        <!-- Layout / Booklet Imposition -->
        <div class="pp-field">
          <label class="pp-label" ${ref("layoutLabelEl")}></label>
          <fluent-dropdown ${ref("layoutDropdown")} @change="${(x) => x.onLayoutChange()}">
            <fluent-option value="1-up" selected>1 Page per Sheet</fluent-option>
            <fluent-option value="booklet" ${ref("bookletOptEl")}></fluent-option>
          </fluent-dropdown>
        </div>

        <!-- Orientation -->
        <div class="pp-field">
          <label class="pp-label" ${ref("orientationLabelEl")}></label>
          <fluent-dropdown
            ${ref("orientationDropdown")}
            @change="${(x) => x.onOrientationChange()}"
          >
            <fluent-option value="auto" selected ${ref("orientAutoOptEl")}></fluent-option>
            <fluent-option value="portrait" ${ref("orientPortraitOptEl")}></fluent-option>
            <fluent-option value="landscape" ${ref("orientLandscapeOptEl")}></fluent-option>
          </fluent-dropdown>
        </div>

        <!-- Paper Size -->
        <div class="pp-field">
          <label class="pp-label" ${ref("paperSizeLabelEl")}></label>
          <fluent-dropdown ${ref("paperSizeDropdown")} @change="${(x) => x.onPaperSizeChange()}">
            <fluent-option value="auto" selected>Auto (Document)</fluent-option>
            <fluent-option value="letter">Letter</fluent-option>
            <fluent-option value="a4">A4</fluent-option>
            <fluent-option value="legal">Legal</fluent-option>
            <fluent-option value="a3">A3</fluent-option>
            <fluent-option value="a5">A5</fluent-option>
          </fluent-dropdown>
        </div>

        <!-- Print Markup Toggle -->
        <div class="pp-field">
          <label class="pp-label" ${ref("markupLabelEl")}></label>
          <fluent-dropdown ${ref("markupDropdown")} @change="${(x) => x.onMarkupChange()}">
            <fluent-option value="document" selected>Print Document</fluent-option>
            <fluent-option value="markup">Print Document with Markup</fluent-option>
          </fluent-dropdown>
        </div>
      </div>

      <!-- Preview Viewport -->
      <div class="pp-viewport" ${ref("viewportEl")}>
        <div class="pp-stage-wrap" ${ref("stageWrapEl")}>
          <div ${ref("previewContainerEl")}></div>
        </div>
      </div>
    </div>

    <!-- Footer Controls -->
    <div class="pp-controls">
      <div class="pp-nav">
        <fluent-button appearance="subtle" ${ref("prevBtn")} @click="${(x) => x.onPrev()}"
          >&lt;</fluent-button
        >
        <span class="pp-indicator" ${ref("pageIndicatorEl")}></span>
        <fluent-button appearance="subtle" ${ref("nextBtn")} @click="${(x) => x.onNext()}"
          >&gt;</fluent-button
        >
      </div>

      <div class="pp-zoom">
        <fluent-button appearance="subtle" ${ref("zoomOutBtn")} @click="${(x) => x.onZoomOut()}"
          >-</fluent-button
        >
        <span class="pp-indicator" ${ref("zoomIndicatorEl")}>100%</span>
        <fluent-button appearance="subtle" ${ref("zoomInBtn")} @click="${(x) => x.onZoomIn()}"
          >+</fluent-button
        >
        <fluent-button
          appearance="neutral"
          ${ref("fitBtn")}
          @click="${(x) => x.onZoomFit()}"
        ></fluent-button>
      </div>
    </div>
  </docen-dialog>
`;

@customElement({ name: "docen-print-preview", template, styles })
export class DocenPrintPreview extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable titleEl?: HTMLElement;
  @observable printBtn?: HTMLElement;
  @observable exportPdfBtn?: HTMLElement;
  @observable closeBtn?: HTMLElement;
  @observable copiesLabelEl?: HTMLElement;
  @observable copiesInput?: HTMLInputElement & { value: string };
  @observable collateLabelEl?: HTMLElement;
  @observable collateDropdown?: HTMLElement & { value: string };
  @observable collatedOptEl?: HTMLElement;
  @observable uncollatedOptEl?: HTMLElement;
  @observable pageRangeLabelEl?: HTMLElement;
  @observable rangeDropdown?: HTMLElement & { value: string };
  @observable rangeAllOptEl?: HTMLElement;
  @observable rangeCurrentOptEl?: HTMLElement;
  @observable rangeSelectionOptEl?: HTMLElement;
  @observable rangeCustomOptEl?: HTMLElement;
  @observable customRangeInput?: HTMLInputElement & { value: string };
  @observable layoutLabelEl?: HTMLElement;
  @observable layoutDropdown?: HTMLElement & { value: string };
  @observable bookletOptEl?: HTMLElement;
  @observable orientationLabelEl?: HTMLElement;
  @observable orientationDropdown?: HTMLElement & { value: string };
  @observable orientAutoOptEl?: HTMLElement;
  @observable orientPortraitOptEl?: HTMLElement;
  @observable orientLandscapeOptEl?: HTMLElement;
  @observable paperSizeLabelEl?: HTMLElement;
  @observable paperSizeDropdown?: HTMLElement & { value: string };
  @observable markupLabelEl?: HTMLElement;
  @observable markupDropdown?: HTMLElement & { value: string };
  @observable viewportEl?: HTMLElement;
  @observable stageWrapEl?: HTMLElement;
  @observable previewContainerEl?: HTMLElement;
  @observable prevBtn?: HTMLElement;
  @observable nextBtn?: HTMLElement;
  @observable pageIndicatorEl?: HTMLElement;
  @observable zoomOutBtn?: HTMLElement;
  @observable zoomInBtn?: HTMLElement;
  @observable fitBtn?: HTMLElement;
  @observable zoomIndicatorEl?: HTMLElement;

  // Internal state
  #snapshots: PdfPageShot[] = [];
  #documentFilename = "Document";
  #currentPage = 1;
  #selectionPages: number[] = [];
  #rangeType: PageRangeType = "all";
  #customRange = "";
  #copies = 1;
  #collate = true;
  #layout: LayoutPreset = "1-up";
  #orientation: OrientationPreset = "auto";
  #paperSize: PaperSizePreset = "auto";
  #printMarkup: PrintMarkupMode = "document";
  #activeViewIndex = 0; // index into effective views (pages or booklet sheets)
  #zoom = 1.0;
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

  /**
   * Seed the print preview dialog with page snapshots and document context.
   */
  seed(options: {
    snapshots: PdfPageShot[];
    filename?: string;
    currentPage?: number;
    selectionPages?: number[];
    printMarkup?: boolean;
  }): void {
    this.#snapshots = options.snapshots ?? [];
    this.#documentFilename = options.filename ?? "Document";
    this.#currentPage = options.currentPage ?? 1;
    this.#selectionPages = options.selectionPages ?? [];
    this.#printMarkup = options.printMarkup ? "markup" : "document";
    this.#activeViewIndex = 0;
    this.#zoom = 1.0;
  }

  show(): void {
    this.#syncControlsToState();
    this.dialogEl?.show();
    this.updatePreview();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  getSnapshots(): readonly PdfPageShot[] {
    return this.#snapshots;
  }

  getEffectivePageNumbers(): number[] {
    const rawPages = parsePageRange({
      type: this.#rangeType,
      custom: this.#customRange,
      currentPage: this.#currentPage,
      selectionPages: this.#selectionPages,
      totalPages: this.#snapshots.length,
    });
    return expandCopies(rawPages, this.#copies, this.#collate);
  }

  getBookletSheets(): BookletSheet[] {
    const rawPages = parsePageRange({
      type: this.#rangeType,
      custom: this.#customRange,
      currentPage: this.#currentPage,
      selectionPages: this.#selectionPages,
      totalPages: this.#snapshots.length,
    });
    return calculateBookletImposition(rawPages.length);
  }

  // --- UI Event Handlers ---

  onCopiesChange(): void {
    const v = Number.parseInt(this.copiesInput?.value ?? "1", 10);
    this.#copies = Math.max(1, isNaN(v) ? 1 : v);
    this.updatePreview();
  }

  onCollateChange(): void {
    this.#collate = this.collateDropdown?.value !== "uncollated";
    this.updatePreview();
  }

  onRangeTypeChange(): void {
    this.#rangeType = (this.rangeDropdown?.value as PageRangeType) ?? "all";
    if (this.customRangeInput) {
      this.customRangeInput.style.display = this.#rangeType === "custom" ? "block" : "none";
    }
    this.#activeViewIndex = 0;
    this.updatePreview();
  }

  onCustomRangeInput(): void {
    this.#customRange = this.customRangeInput?.value ?? "";
    this.#activeViewIndex = 0;
    this.updatePreview();
  }

  onLayoutChange(): void {
    this.#layout = (this.layoutDropdown?.value as LayoutPreset) ?? "1-up";
    this.#activeViewIndex = 0;
    this.updatePreview();
  }

  onOrientationChange(): void {
    this.#orientation = (this.orientationDropdown?.value as OrientationPreset) ?? "auto";
    this.updatePreview();
  }

  onPaperSizeChange(): void {
    this.#paperSize = (this.paperSizeDropdown?.value as PaperSizePreset) ?? "auto";
    this.updatePreview();
  }

  onMarkupChange(): void {
    this.#printMarkup = (this.markupDropdown?.value as PrintMarkupMode) ?? "document";
    this.dispatchEvent(
      new CustomEvent("print-preview:markup-change", {
        bubbles: true,
        composed: true,
        detail: { markup: this.#printMarkup === "markup" },
      }),
    );
    this.updatePreview();
  }

  onPrev(): void {
    if (this.#activeViewIndex > 0) {
      this.#activeViewIndex--;
      this.updatePreview();
    }
  }

  onNext(): void {
    const max = this.#getMaxViewCount();
    if (this.#activeViewIndex < max - 1) {
      this.#activeViewIndex++;
      this.updatePreview();
    }
  }

  onZoomIn(): void {
    this.#zoom = Math.min(3.0, Math.round((this.#zoom + 0.1) * 10) / 10);
    this.#applyZoom();
  }

  onZoomOut(): void {
    this.#zoom = Math.max(0.2, Math.round((this.#zoom - 0.1) * 10) / 10);
    this.#applyZoom();
  }

  onZoomFit(): void {
    this.#zoom = 1.0;
    this.#applyZoom();
  }

  #applyZoom(): void {
    if (this.stageWrapEl) {
      this.stageWrapEl.style.transform = `scale(${this.#zoom})`;
    }
    if (this.zoomIndicatorEl) {
      this.zoomIndicatorEl.textContent = `${Math.round(this.#zoom * 100)}%`;
    }
  }

  #getMaxViewCount(): number {
    if (this.#layout === "booklet") {
      const sheets = this.getBookletSheets();
      return Math.max(1, sheets.length);
    }
    const pages = this.getEffectivePageNumbers();
    return Math.max(1, pages.length);
  }

  // --- Rendering the Canvas Preview ---

  updatePreview(): void {
    if (!this.previewContainerEl) return;
    this.previewContainerEl.replaceChildren();

    const maxViews = this.#getMaxViewCount();
    if (this.#activeViewIndex >= maxViews) {
      this.#activeViewIndex = Math.max(0, maxViews - 1);
    }

    if (this.#layout === "booklet") {
      this.#renderBookletView();
    } else {
      this.#renderSinglePageView();
    }

    this.#applyZoom();
  }

  #renderSinglePageView(): void {
    const pages = this.getEffectivePageNumbers();
    if (pages.length === 0 || this.#snapshots.length === 0) {
      if (this.pageIndicatorEl) this.pageIndicatorEl.textContent = "0 of 0";
      return;
    }

    const pageNum = pages[this.#activeViewIndex] ?? 1;
    const shot = this.#snapshots[pageNum - 1];

    if (this.pageIndicatorEl) {
      this.pageIndicatorEl.textContent = t("printPreview.pageOf", this)
        .replace("{current}", String(this.#activeViewIndex + 1))
        .replace("{total}", String(pages.length));
    }

    if (!shot) return;

    const dims = resolvePaperDimensions(
      this.#paperSize,
      this.#orientation,
      shot.width,
      shot.height,
    );

    const img = document.createElement("img");
    img.className = "pp-page-shadow";
    img.src = shot.url ?? "";
    img.style.width = `${dims.width * 0.65}px`;
    img.style.height = `${dims.height * 0.65}px`;

    this.previewContainerEl?.append(img);
  }

  #renderBookletView(): void {
    const sheets = this.getBookletSheets();
    if (sheets.length === 0) {
      if (this.pageIndicatorEl) this.pageIndicatorEl.textContent = "0 of 0";
      return;
    }

    const currentSheet = sheets[this.#activeViewIndex];
    if (this.pageIndicatorEl) {
      this.pageIndicatorEl.textContent = t("printPreview.sheetOf", this)
        .replace("{current}", String(this.#activeViewIndex + 1))
        .replace("{total}", String(sheets.length));
    }

    if (!currentSheet) return;

    const wrap = document.createElement("div");
    wrap.className = "pp-sheet-wrap";

    const baseWidth = this.#snapshots[0]?.width ?? 793.7;
    const baseHeight = this.#snapshots[0]?.height ?? 1122.5;
    const dims = resolvePaperDimensions(this.#paperSize, this.#orientation, baseWidth, baseHeight);

    const renderHalf = (pageNum: number | null): HTMLElement => {
      const half = document.createElement("div");
      half.className = "pp-sheet-half";
      half.style.width = `${dims.width * 0.45}px`;
      half.style.height = `${dims.height * 0.45}px`;

      if (pageNum == null) {
        const blank = document.createElement("div");
        blank.className = "pp-sheet-blank";
        blank.textContent = "[ Blank ]";
        half.append(blank);
      } else {
        const shot = this.#snapshots[pageNum - 1];
        if (shot?.url) {
          const img = document.createElement("img");
          img.src = shot.url;
          img.style.width = "100%";
          img.style.height = "100%";
          img.style.display = "block";
          half.append(img);
        }
      }
      return half;
    };

    wrap.append(renderHalf(currentSheet.leftPage));
    wrap.append(renderHalf(currentSheet.rightPage));
    this.previewContainerEl?.append(wrap);
  }

  // --- Output Actions (Print & Export PDF) ---

  async onPrint(): Promise<void> {
    const pages = this.getEffectivePageNumbers();
    if (pages.length === 0 || this.#snapshots.length === 0) return;

    const orderedShots = pages.map((p) => this.#snapshots[p - 1]).filter(Boolean) as PdfPageShot[];

    if (orderedShots.length === 0) return;

    const first = orderedShots[0]!;
    const dims = resolvePaperDimensions(
      this.#paperSize,
      this.#orientation,
      first.width,
      first.height,
    );

    const frame = document.createElement("iframe");
    Object.assign(frame.style, {
      position: "fixed",
      right: "0",
      bottom: "0",
      width: "0",
      height: "0",
      border: "0",
    });
    document.body.append(frame);

    const doc = frame.contentDocument!;
    doc.open();
    doc.write(`<!doctype html><html><head><title>${this.#documentFilename}</title><style>
      @page { size: ${dims.width / 96}in ${dims.height / 96}in; margin: 0; }
      html, body { margin: 0; padding: 0; }
      img { display: block; width: 100%; height: 100%; }
      .pg { page-break-after: always; break-after: page; width: 100%; height: 100%; }
      .pg:last-child { page-break-after: auto; break-after: auto; }
    </style></head><body>`);

    for (const s of orderedShots) {
      if (s.url) {
        doc.write(`<div class="pg"><img src="${s.url}"></div>`);
      }
    }
    doc.write("</body></html>");
    doc.close();

    frame.onload = () => {
      const win = frame.contentWindow;
      if (!win) return;
      const cleanup = (): void => frame.remove();
      win.addEventListener("afterprint", cleanup, { once: true });
      win.focus();
      win.print();
      setTimeout(cleanup, 30_000);
    };

    this.dispatchEvent(
      new CustomEvent("print-preview:print", {
        bubbles: true,
        composed: true,
        detail: { pages: orderedShots.length },
      }),
    );
    this.hide();
  }

  async onExportPdf(): Promise<void> {
    const pages = this.getEffectivePageNumbers();
    if (pages.length === 0 || this.#snapshots.length === 0) return;

    const orderedShots = pages.map((p) => this.#snapshots[p - 1]).filter(Boolean) as PdfPageShot[];

    if (orderedShots.length === 0) return;

    const pdfBytes = await pagesToPdf(orderedShots, {
      metadata: { title: this.#documentFilename },
    });

    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${this.#documentFilename.replace(/\.[^/.]+$/, "")}.pdf`;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);

    this.dispatchEvent(
      new CustomEvent("print-preview:export-pdf", {
        bubbles: true,
        composed: true,
        detail: { blob },
      }),
    );
    this.hide();
  }

  #syncControlsToState(): void {
    if (this.copiesInput) this.copiesInput.value = String(this.#copies);
    if (this.customRangeInput) {
      this.customRangeInput.style.display = this.#rangeType === "custom" ? "block" : "none";
    }
  }

  #applyLabels(): void {
    if (this.titleEl) this.titleEl.textContent = t("printPreview.title", this);
    if (this.printBtn) this.printBtn.textContent = t("printPreview.print", this);
    if (this.exportPdfBtn) this.exportPdfBtn.textContent = t("printPreview.exportPdf", this);
    if (this.closeBtn) this.closeBtn.textContent = t("options.cancel", this);
    if (this.copiesLabelEl) this.copiesLabelEl.textContent = t("printPreview.copies", this);
    if (this.collateLabelEl) this.collateLabelEl.textContent = t("printPreview.collate", this);
    if (this.collatedOptEl) this.collatedOptEl.textContent = t("printPreview.collated", this);
    if (this.uncollatedOptEl) this.uncollatedOptEl.textContent = t("printPreview.uncollated", this);
    if (this.pageRangeLabelEl)
      this.pageRangeLabelEl.textContent = t("printPreview.pageRange", this);
    if (this.rangeAllOptEl) this.rangeAllOptEl.textContent = t("printPreview.rangeAll", this);
    if (this.rangeCurrentOptEl)
      this.rangeCurrentOptEl.textContent = t("printPreview.rangeCurrent", this);
    if (this.rangeSelectionOptEl)
      this.rangeSelectionOptEl.textContent = t("printPreview.rangeSelection", this);
    if (this.rangeCustomOptEl)
      this.rangeCustomOptEl.textContent = t("printPreview.rangeCustom", this);
    if (this.customRangeInput) {
      this.customRangeInput.placeholder = t("printPreview.customRangePlaceholder", this);
    }
    if (this.layoutLabelEl) this.layoutLabelEl.textContent = t("printPreview.pages", this);
    if (this.bookletOptEl) this.bookletOptEl.textContent = t("printPreview.booklet", this);
    if (this.orientationLabelEl)
      this.orientationLabelEl.textContent = t("printPreview.orientation", this);
    if (this.orientAutoOptEl)
      this.orientAutoOptEl.textContent = t("printPreview.orientationAuto", this);
    if (this.orientPortraitOptEl)
      this.orientPortraitOptEl.textContent = t("printPreview.orientationPortrait", this);
    if (this.orientLandscapeOptEl)
      this.orientLandscapeOptEl.textContent = t("printPreview.orientationLandscape", this);
    if (this.paperSizeLabelEl)
      this.paperSizeLabelEl.textContent = t("printPreview.paperSize", this);
    if (this.markupLabelEl) this.markupLabelEl.textContent = t("printPreview.printMarkup", this);
    if (this.fitBtn) this.fitBtn.textContent = t("printPreview.zoomFit", this);
  }
}

export default DocenPrintPreview;
