import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, resolveDir, t } from "../../i18n/localize";

export interface StatusBarWidgetsConfig {
  pageNumber: boolean;
  wordCount: boolean;
  language: boolean;
  extendSelection: boolean;
  capsLock: boolean;
  zoom: boolean;
}

export const STATUS_BAR_STORAGE_KEY = "docen:status-bar-widgets";

export const DEFAULT_STATUS_BAR_CONFIG: StatusBarWidgetsConfig = {
  pageNumber: true,
  wordCount: true,
  language: true,
  extendSelection: true,
  capsLock: true,
  zoom: true,
};

const styles = css`
  :host {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 8px;
    width: 100%;
  }
  :host([dir="rtl"]),
  :host([data-dir="rtl"]),
  :host-context([dir="rtl"]) {
    flex-direction: row-reverse;
    direction: rtl;
  }
  :host([dir="rtl"]) .pct,
  :host([data-dir="rtl"]) .pct,
  :host-context([dir="rtl"]) .pct {
    text-align: left;
  }
  .left {
    display: flex;
    align-items: center;
    gap: 14px;
    min-width: 0;
    overflow: hidden;
  }
  /* Right cluster — Word's zoom control: a minus / plus button flanking a
     draggable slider, then the percent. The slider is a native range input
     styled to a Fluent track + accent thumb. */
  .zoom {
    display: flex;
    align-items: center;
    gap: 4px;
  }
  .step {
    width: 18px;
    height: 18px;
    padding: 0;
    border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 3px;
    background: transparent;
    color: var(--docen-color-text-1, #242424);
    font-size: 13px;
    line-height: 1;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    justify-content: center;
  }
  .step:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  .slider {
    -webkit-appearance: none;
    appearance: none;
    width: 90px;
    height: 3px;
    margin: 0;
    background: var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 2px;
    cursor: pointer;
  }
  .slider::-webkit-slider-thumb {
    -webkit-appearance: none;
    appearance: none;
    width: 11px;
    height: 11px;
    border: none;
    border-radius: 50%;
    background: var(--docen-color-accent, #0f6cbd);
    cursor: pointer;
  }
  .slider::-moz-range-thumb {
    width: 11px;
    height: 11px;
    border: none;
    border-radius: 50%;
    background: var(--docen-color-accent, #0f6cbd);
    cursor: pointer;
  }
  .pct {
    min-width: 38px;
    text-align: right;
    cursor: pointer;
    border-radius: 3px;
    padding: 1px 3px;
  }
  .pct:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  /* View shortcuts left of the zoom slider — Word's status bar carries
     Reading / Print Layout / Web Layout buttons; the host reports the active
     view through the view attribute ("read" | "print" | "web" | "draft" —
     Draft has no status-bar button, Word's doesn't either). */
  .views {
    display: flex;
    align-items: center;
    gap: 2px;
    margin-inline-end: 6px;
  }
  .view-btn {
    width: 22px;
    height: 22px;
    padding: 0;
    border: none;
    border-radius: 3px;
    background: transparent;
    color: var(--docen-color-text-2, #424242);
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    justify-content: center;
  }
  .view-btn:hover:not(:disabled) {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  .view-btn:disabled {
    color: var(--docen-color-text-3, #8a8a8a);
    opacity: 0.55;
    cursor: default;
  }
  .view-btn[aria-pressed="true"] {
    background: var(--docen-color-subtle-background-selected, #e8e8e8);
    color: var(--docen-color-accent, #0f6cbd);
  }
  /* Proofing state (Word's status-bar book): a check when the document is
     clean, a red cross when misspellings are pending. Click opens the
     Spelling pane. */
  .spell-btn {
    width: 22px;
    height: 22px;
    padding: 0;
    border: none;
    border-radius: 3px;
    background: transparent;
    color: var(--docen-color-text-2, #424242);
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    justify-content: center;
  }
  .spell-btn:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  .spell-btn[data-state="ok"] .spell-issues,
  .spell-btn[data-state="issues"] .spell-ok {
    display: none;
  }
  .spell-btn[data-state="issues"] {
    color: #e81123;
  }
  /* Language indicator — sat after the word count. Plain text matching the
     surrounding status copy; a click cycles through every registered locale. */
  .lang-text {
    cursor: pointer;
    padding-inline: 2px;
  }
  /* Word count opens the statistics dialog, same plain-text affordance. */
  .words {
    cursor: pointer;
  }
  .extend-selection,
  .caps-lock {
    padding-inline: 4px;
    font-size: 11px;
    cursor: pointer;
    border-radius: 2px;
  }
  .extend-selection:hover,
  .caps-lock:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  .status-context-menu {
    position: fixed;
    margin: 0;
    padding: 4px 0;
    border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 4px;
    background: var(--docen-color-bg, #ffffff);
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.18);
    font-family: inherit;
    font-size: 12px;
    color: var(--docen-color-text-1, #242424);
    min-width: 180px;
    z-index: 10005;
  }
  .status-menu-header {
    padding: 6px 12px 4px;
    font-size: 11px;
    font-weight: 600;
    color: var(--docen-color-secondary, #595959);
    border-bottom: 1px solid var(--docen-color-divider, #e1e1e1);
    margin-bottom: 2px;
  }
  .status-menu-item {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 5px 12px;
    cursor: pointer;
    user-select: none;
    -webkit-user-select: none;
  }
  .status-menu-item:hover {
    background: var(--docen-color-subtle-background-hover, #f0f0f0);
  }
  .status-menu-check {
    width: 14px;
    font-size: 12px;
    font-weight: bold;
    color: var(--docen-color-accent, #0f6cbd);
  }
  /* Narrow viewports: drop the less essential items progressively and shrink
     the zoom slider, so the bar fits a phone width without overflowing. */
  @media (max-width: 720px) {
    .section {
      display: none;
    }
    .slider {
      width: 48px;
    }
  }
  .extend-mode {
    font-weight: 600;
    padding: 0 4px;
    color: var(--docen-color-accent, #0f6cbd);
  }
  @media (max-width: 560px) {
    .views,
    .lang-text,
    .extend-mode {
      display: none;
    }
  }
`;

const template = html<DocenStatusBar>`
  <span class="left">
    <span class="section" ${ref("sectionEl")}></span>
    <span class="pages" ${ref("pagesEl")}></span>
    <span class="words" ${ref("wordsEl")}></span>
    <button
      type="button"
      class="spell-btn"
      data-state="ok"
      ${ref("spellBtn")}
      aria-label="Proofing"
    >
      <svg
        viewBox="0 0 16 16"
        width="15"
        height="15"
        fill="none"
        stroke="currentColor"
        stroke-width="1.2"
        stroke-linejoin="round"
        stroke-linecap="round"
        aria-hidden="true"
      >
        <path
          d="M8 3.5C6.9 2.6 5.2 2 3 2v11c2.2 0 3.9.6 5 1.5 1.1-.9 2.8-1.5 5-1.5V2c-2.2 0-3.9.6-5 1.5z"
        />
        <path class="spell-ok" d="M5.6 7.6l1.7 1.7 3.2-3.4" />
        <path class="spell-issues" d="M6 6.2l4 4M10 6.2l-4 4" />
      </svg>
    </button>
    <span class="lang-text" ${ref("langBtn")}></span>
    <span class="extend-mode" ${ref("extendModeEl")}></span>
    <span class="extend-selection" ${ref("extendSelEl")}></span>
    <span class="caps-lock" ${ref("capsLockEl")}></span>
  </span>
  <span class="zoom" ${ref("zoomEl")}>
    <span class="views" ${ref("viewsEl")}>
      <button type="button" class="view-btn" data-view="reading">
        <svg
          viewBox="0 0 16 16"
          width="15"
          height="15"
          fill="none"
          stroke="currentColor"
          stroke-width="1.2"
          stroke-linejoin="round"
          aria-hidden="true"
        >
          <path
            d="M8 3.5C6.9 2.6 5.2 2 3 2v11c2.2 0 3.9.6 5 1.5 1.1-.9 2.8-1.5 5-1.5V2c-2.2 0-3.9.6-5 1.5z"
          />
          <path d="M8 3.5v11" />
        </svg>
      </button>
      <button type="button" class="view-btn" data-view="print">
        <svg
          viewBox="0 0 16 16"
          width="15"
          height="15"
          fill="none"
          stroke="currentColor"
          stroke-width="1.2"
          stroke-linejoin="round"
          aria-hidden="true"
        >
          <path d="M3.5 1.5h6l3 3v10h-9z" />
          <path d="M9.5 1.5v3h3" />
        </svg>
      </button>
      <button type="button" class="view-btn" data-view="web">
        <svg
          viewBox="0 0 16 16"
          width="15"
          height="15"
          fill="none"
          stroke="currentColor"
          stroke-width="1.2"
          stroke-linejoin="round"
          aria-hidden="true"
        >
          <circle cx="8" cy="8" r="6.5" />
          <path d="M1.5 8h13M8 1.5c-4.7 4.2-4.7 8.8 0 13 4.7-4.2 4.7-8.8 0-13z" />
        </svg>
      </button>
    </span>
    <button type="button" class="step" ${ref("outBtn")} aria-label="Zoom out">−</button>
    <input
      type="range"
      class="slider"
      min="10"
      max="500"
      step="1"
      value="100"
      ${ref("slider")}
      aria-label="Zoom level"
    />
    <button type="button" class="step" ${ref("inBtn")} aria-label="Zoom in">+</button>
    <span class="pct" ${ref("pctEl")}></span>
  </span>
  <div popover="auto" class="status-context-menu" ${ref("contextMenuEl")}></div>
`;

/**
 * `<docen-status-bar>` — Word's bottom status bar: a left cluster (caret
 * section, "Page X of Y", word count) and a right zoom control (− / slider / +
 * / percent). Numeric state arrives as attributes (`section` / `page` / `total`
 * / `words` / `zoom`); the labels are localized here. Every item renders only
 * when its attribute is stamped — a host without a surface (the presentation
 * host has no section/spell nouns) leaves it out. Zoom interaction emits
 * `zoom:change { zoom }` (percent, 10–500) for the host to apply.
 */
@customElement({ name: "docen-status-bar", template, styles })
class DocenStatusBar extends FASTElement {
  @attr section?: string;
  @attr page?: string;
  @attr total?: string;
  /** Full localized template for the page indicator, overriding the built-in
   *  "Page {page} of {total}" — a host with different page nouns stamps its
   *  own (the presentation element passes "Slide {page} of {total}"). */
  @attr pageLabel?: string;
  @attr words?: string;
  @attr zoom?: string;
  /** The active document view — "read" | "print" | "web" | "draft" — drives
   *  the view buttons' pressed state. */
  @attr view?: string;
  /** The proofing state — "ok" | "issues" — drives the status-bar book icon
   *  (a green check vs a red cross, Word's spell indicator). Named `proofing`
   *  because `spellcheck` is a native HTMLElement property (boolean). */
  @attr proofing?: string;
  /** The caret's proofing-language display name (Word shows the selection's
   *  w:lang in the status bar); a click opens the language dialog. */
  @attr language?: string;
  /** Extend selection mode ("EXT" Word status indicator). */
  @attr extend?: string;
  @attr dir: "ltr" | "rtl" = "ltr";

  get isRtl(): boolean {
    return this.dir === "rtl" || this.getAttribute("dir") === "rtl" || resolveDir(this) === "rtl";
  }

  dirChanged(): void {
    this.#syncDir();
  }

  #syncDir(): void {
    const isRtl = (this.dir ?? resolveDir(this)) === "rtl";
    this.toggleAttribute("data-dir", isRtl);
    if (isRtl) {
      this.setAttribute("dir", "rtl");
    } else if (this.getAttribute("dir") === "rtl" && !this.getAttribute("data-dir")) {
      // keep explicit attr
    } else {
      this.removeAttribute("dir");
    }
  }

  @observable sectionEl?: HTMLElement;
  @observable pagesEl?: HTMLElement;
  @observable wordsEl?: HTMLElement;
  @observable viewsEl?: HTMLElement;
  @observable slider?: HTMLInputElement;
  @observable pctEl?: HTMLElement;
  @observable outBtn?: HTMLButtonElement;
  @observable inBtn?: HTMLButtonElement;
  @observable langBtn?: HTMLElement;
  @observable extendModeEl?: HTMLElement;
  @observable spellBtn?: HTMLButtonElement;
  @observable extendSelEl?: HTMLElement;
  @observable capsLockEl?: HTMLElement;
  @observable zoomEl?: HTMLElement;
  @observable contextMenuEl?: HTMLElement;
  #unsubscribe?: () => void;
  #widgetConfig: StatusBarWidgetsConfig = DEFAULT_STATUS_BAR_CONFIG;
  #onWindowKey?: (e: KeyboardEvent) => void;

  get widgetConfig(): StatusBarWidgetsConfig {
    return { ...this.#widgetConfig };
  }

  get widgetsConfig(): StatusBarWidgetsConfig {
    return { ...this.#widgetConfig };
  }

  setWidgetVisible(widget: keyof StatusBarWidgetsConfig, visible: boolean): void {
    this.#widgetConfig[widget] = visible;
    this.#saveWidgetConfig();
    this.#applyWidgetConfig();
  }

  setWidgetVisibility(widget: keyof StatusBarWidgetsConfig, visible: boolean): void {
    this.setWidgetVisible(widget, visible);
  }

  #loadWidgetConfig(): StatusBarWidgetsConfig {
    try {
      const raw = localStorage.getItem(STATUS_BAR_STORAGE_KEY);
      if (raw) return { ...DEFAULT_STATUS_BAR_CONFIG, ...JSON.parse(raw) };
    } catch {}
    return { ...DEFAULT_STATUS_BAR_CONFIG };
  }

  #saveWidgetConfig(): void {
    try {
      localStorage.setItem(STATUS_BAR_STORAGE_KEY, JSON.stringify(this.#widgetConfig));
    } catch {}
  }

  #applyWidgetConfig(): void {
    if (this.pagesEl) this.pagesEl.style.display = this.#widgetConfig.pageNumber ? "" : "none";
    if (this.wordsEl) this.wordsEl.style.display = this.#widgetConfig.wordCount ? "" : "none";
    if (this.langBtn) this.langBtn.style.display = this.#widgetConfig.language ? "" : "none";
    if (this.extendSelEl)
      this.extendSelEl.style.display = this.#widgetConfig.extendSelection ? "" : "none";
    const capsEl = this.capsLockEl ?? this.shadowRoot?.querySelector<HTMLElement>(".caps-lock");
    if (capsEl) capsEl.style.display = "none";
    if (this.zoomEl) this.zoomEl.style.display = this.#widgetConfig.zoom ? "" : "none";
  }

  #openContextMenu(clientX: number, clientY: number): void {
    const menu =
      this.contextMenuEl ?? this.shadowRoot?.querySelector<HTMLElement>(".status-context-menu");
    if (!menu) return;
    menu.replaceChildren();

    const header = document.createElement("div");
    header.className = "status-menu-header";
    header.textContent = t("status.customize", this);
    menu.append(header);

    const WIDGETS: Array<{ key: keyof StatusBarWidgetsConfig; label: string }> = [
      { key: "pageNumber", label: t("status.pageNumber", this) },
      { key: "wordCount", label: t("status.wordCount", this) },
      { key: "language", label: t("status.language", this) },
      { key: "extendSelection", label: t("status.extendSelection", this) },
      { key: "capsLock", label: t("status.capsLock", this) },
      { key: "zoom", label: t("status.zoom", this) },
    ];

    for (const w of WIDGETS) {
      const row = document.createElement("div");
      row.className = "status-menu-item";
      row.dataset.widget = w.key;

      const check = document.createElement("span");
      check.className = "status-menu-check";
      check.textContent = this.#widgetConfig[w.key] ? "✓" : "";

      const label = document.createElement("span");
      label.textContent = w.label;

      row.append(check, label);
      row.addEventListener("click", (e) => {
        e.stopPropagation();
        this.setWidgetVisible(w.key, !this.#widgetConfig[w.key]);
        check.textContent = this.#widgetConfig[w.key] ? "✓" : "";
      });
      menu.append(row);
    }

    const menuHeight = 220;
    const menuWidth = 190;
    let top = clientY - menuHeight;
    if (top < 10) top = clientY + 10;
    let left = clientX;
    if (typeof window !== "undefined" && left + menuWidth > window.innerWidth) {
      left = window.innerWidth - menuWidth - 10;
    }

    menu.style.left = `${Math.round(left)}px`;
    menu.style.top = `${Math.round(top)}px`;
    menu.style.display = "block";

    (menu as unknown as { showPopover?: () => void }).showPopover?.();
  }

  extendChanged(): void {
    this.#renderExtend();
  }

  sectionChanged(): void {
    this.#renderSection();
  }

  pageChanged(): void {
    this.#renderPages();
  }
  totalChanged(): void {
    this.#renderPages();
  }
  pageLabelChanged(): void {
    this.#renderPages();
  }
  wordsChanged(): void {
    this.#renderWords();
  }
  zoomChanged(): void {
    this.#renderZoom();
  }
  viewChanged(): void {
    this.#syncViewPressed();
  }
  proofingChanged(): void {
    this.#syncSpellState();
  }
  languageChanged(): void {
    this.#renderLanguage();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#syncDir();
    this.#widgetConfig = this.#loadWidgetConfig();
    this.#renderAll();
    this.#applyWidgetConfig();
    this.addEventListener("contextmenu", this.#onContextMenu);
    this.#onWindowKey = (e: KeyboardEvent): void => {
      if (!this.isConnected) return;
      const caps = Boolean(e.getModifierState?.("CapsLock"));
      const el = this.capsLockEl ?? this.shadowRoot?.querySelector<HTMLElement>(".caps-lock");
      if (el) {
        el.textContent = caps ? t("status.capsLock", this) : "";
        el.style.display = caps && this.#widgetConfig.capsLock ? "inline-flex" : "none";
      }
    };
    window.addEventListener("keydown", this.#onWindowKey);
    window.addEventListener("keyup", this.#onWindowKey);
    if (this.extendSelEl) {
      this.extendSelEl.textContent = t("status.extendSelection", this);
    }
    // Slider drags live; the minus / plus buttons step by 10% (Word behavior).
    this.slider?.addEventListener("input", () => this.#emit(Number(this.slider?.value ?? 100)));
    this.outBtn?.addEventListener("click", () => this.#emit(Number(this.zoom ?? 100) - 10));
    this.inBtn?.addEventListener("click", () => this.#emit(Number(this.zoom ?? 100) + 10));
    // The language item opens the proofing-language dialog (Word) — it does
    // NOT switch the UI language (that lives in File → Options).
    this.langBtn?.addEventListener("click", () => this.#emitOpenLanguage());
    // The percent opens the Zoom dialog (Word), and the view shortcuts select
    // a document view — the host decides which view each name maps to.
    this.pctEl?.addEventListener("click", () => this.#emitOpenZoom());
    // The word count opens the Word Count dialog (Word).
    this.wordsEl?.addEventListener("click", () => this.#emitOpenWordCount());
    // The proofing book opens the Spelling pane (Word: click → Spelling).
    this.spellBtn?.addEventListener("click", () => this.#emitOpenSpelling());
    for (const btn of this.shadowRoot?.querySelectorAll<HTMLButtonElement>(".view-btn") ?? []) {
      btn.addEventListener("click", () =>
        this.dispatchEvent(
          new CustomEvent("view:select", {
            bubbles: true,
            composed: true,
            detail: { view: btn.dataset.view },
          }),
        ),
      );
    }
    this.#unsubscribe = observeLang(() => {
      this.#syncDir();
      this.#renderAll();
      this.#renderViewTitles();
    });
    this.#renderViewTitles();
    this.#syncViewPressed();
    this.#syncSpellState();
  }

  /** Localized tooltips for the view shortcuts (the same Word view names the
   *  View tab uses). */
  #renderViewTitles(): void {
    for (const btn of this.shadowRoot?.querySelectorAll<HTMLButtonElement>(".view-btn") ?? []) {
      const key =
        btn.dataset.view === "reading"
          ? "ribbon.cmd.read-mode"
          : btn.dataset.view === "web"
            ? "ribbon.cmd.web-layout"
            : "ribbon.cmd.print-layout";
      btn.title = t(key, this);
    }
  }

  /** The view buttons' pressed state mirrors the host's active view (none
   *  pressed in Draft — Word's status bar has no Draft button to light). An
   *  unstamped view hides the cluster (hosts without view surfaces). */
  #syncViewPressed(): void {
    if (this.viewsEl) this.viewsEl.style.display = this.view == null ? "none" : "";
    const active =
      this.view === "read"
        ? "reading"
        : this.view === "web"
          ? "web"
          : this.view === "print"
            ? "print"
            : null;
    for (const btn of this.shadowRoot?.querySelectorAll<HTMLButtonElement>(".view-btn") ?? []) {
      if (active) btn.setAttribute("aria-pressed", String(btn.dataset.view === active));
      else btn.removeAttribute("aria-pressed");
    }
  }

  #emitOpenZoom(): void {
    this.dispatchEvent(
      new CustomEvent("zoom:open", {
        bubbles: true,
        composed: true,
        detail: { zoom: Number(this.zoom ?? 100) },
      }),
    );
  }

  #emitOpenWordCount(): void {
    this.dispatchEvent(new CustomEvent("wordcount:open", { bubbles: true, composed: true }));
  }

  #emitOpenSpelling(): void {
    this.dispatchEvent(new CustomEvent("spellcheck:open", { bubbles: true, composed: true }));
  }

  #emitOpenLanguage(): void {
    this.dispatchEvent(new CustomEvent("language:open", { bubbles: true, composed: true }));
  }

  /** The proofing-language item mirrors the host-provided display name (the
   *  caret run's w:lang, like Word's status bar). */
  #renderLanguage(): void {
    if (this.langBtn) this.langBtn.textContent = this.language ?? "";
  }

  /** The book's check/cross face + localized tooltip mirror the host's
   *  proofing state; an unstamped state hides the book. */
  #syncSpellState(): void {
    if (!this.spellBtn) return;
    this.spellBtn.style.display = this.proofing == null ? "none" : "";
    this.spellBtn.dataset.state = this.proofing === "issues" ? "issues" : "ok";
    this.spellBtn.title = t("status.spelling", this);
  }

  readonly #onContextMenu = (e: MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    this.#openContextMenu(e.clientX, e.clientY);
  };

  disconnectedCallback(): void {
    this.removeEventListener("contextmenu", this.#onContextMenu);
    if (this.#onWindowKey) {
      window.removeEventListener("keydown", this.#onWindowKey);
      window.removeEventListener("keyup", this.#onWindowKey);
      this.#onWindowKey = undefined;
    }
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  #emit(zoom: number): void {
    this.dispatchEvent(
      new CustomEvent("zoom:change", {
        bubbles: true,
        composed: true,
        detail: { zoom: Math.max(10, Math.min(500, Math.round(zoom))) },
      }),
    );
  }

  #renderAll(): void {
    this.#renderSection();
    this.#renderPages();
    this.#renderWords();
    this.#renderZoom();
    this.#renderExtend();
  }

  #renderExtend(): void {
    if (!this.extendModeEl) return;
    this.extendModeEl.textContent = this.extend ? "EXT" : "";
    this.extendModeEl.style.display = this.extend ? "inline-block" : "none";
  }

  #renderSection(): void {
    // Hosts render exactly what they stamp — an unstamped attr collapses its
    // item (the presentation host has no section/word nouns).
    if (this.sectionEl)
      this.sectionEl.textContent =
        this.section == null
          ? ""
          : t("status.section", this).replace("{n}", String(Number(this.section)));
  }

  #renderPages(): void {
    if (!this.pagesEl) return;
    if (this.page == null) {
      this.pagesEl.textContent = "";
      return;
    }
    this.pagesEl.textContent = (this.pageLabel || t("status.page-of", this))
      .replace("{page}", String(Number(this.page)))
      .replace("{total}", String(Number(this.total || 1)));
  }

  #renderWords(): void {
    if (!this.wordsEl) return;
    if (this.words == null) {
      this.wordsEl.textContent = "";
      return;
    }
    if (this.words.includes("/")) {
      const parts = this.words.split("/").map((p) => p.trim());
      const sel = parts[0] ?? "0";
      const total = parts[1] ?? "0";
      const selFmt = t("status.words-selected", this);
      if (selFmt && selFmt !== "status.words-selected") {
        this.wordsEl.textContent = selFmt.replace("{sel}", sel).replace("{total}", total);
      } else {
        this.wordsEl.textContent =
          `${sel} / ${total} ${t("status.words", this).replace("{n}", "").trim()}`.trim();
      }
    } else {
      this.wordsEl.textContent = t("status.words", this).replace("{n}", String(Number(this.words)));
    }
  }

  #renderZoom(): void {
    const z = Number(this.zoom ?? 100);
    // Sync the slider without retriggering its own input handler — only write
    // when the value drifted (keyboard / ribbon zoom changed it out of band).
    if (this.slider && Number(this.slider.value) !== z) this.slider.value = String(z);
    if (this.pctEl) this.pctEl.textContent = `${z}%`;
  }
}

if (!customElements.get("docen-status-bar")) {
  customElements.define("docen-status-bar", DocenStatusBar);
}

export { DocenStatusBar };
export default DocenStatusBar;
