import type { StylesOptions } from "@docen/docx";
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

import { detectNodeHeadingLevel, moveHeadingSection } from "../../../document/commands/outline";
import { observeLang, resolveDir, t } from "../../i18n/localize";

export interface NavHeadingItem {
  id: string;
  title: string;
  level: number;
  pos: number;
  nodeIndex: number;
  collapsed?: boolean;
  hasChildren?: boolean;
  hasMatch?: boolean;
  partBefore?: string;
  partMatch?: string;
  partAfter?: string;
}

export interface NavPageItem {
  pageNumber: number;
  active?: boolean;
  thumbnailUrl?: string | null;
  snippet?: string;
}

export interface NavSearchResultItem {
  id: string;
  from: number;
  to: number;
  headingTitle: string;
  before: string;
  match: string;
  after: string;
}

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
    min-height: 0;
    box-sizing: border-box;
    font-family: var(
      --docen-font-family,
      -apple-system,
      BlinkMacSystemFont,
      "Segoe UI",
      Roboto,
      sans-serif
    );
    font-size: 12px;
    color: var(--docen-color-text, #242424);
    background: var(--docen-color-bg, #fbfbfb);
    user-select: none;
    overflow: hidden;
  }

  :host([dir="rtl"]),
  :host([data-dir="rtl"]),
  :host-context([dir="rtl"]) {
    direction: rtl;
  }

  :host([dir="rtl"]) .tree-view,
  :host([data-dir="rtl"]) .tree-view,
  :host-context([dir="rtl"]) .tree-view {
    direction: rtl;
  }

  :host([dir="rtl"]) .tree-item,
  :host([data-dir="rtl"]) .tree-item,
  :host-context([dir="rtl"]) .tree-item {
    border-left: none;
    border-right: 3px solid transparent;
    text-align: right;
  }

  :host([dir="rtl"]) .tree-item.active,
  :host([data-dir="rtl"]) .tree-item.active,
  :host-context([dir="rtl"]) .tree-item.active {
    border-left-color: transparent;
    border-right-color: #0f6cbd;
  }

  :host([dir="rtl"]) .chevron,
  :host([data-dir="rtl"]) .chevron,
  :host-context([dir="rtl"]) .chevron {
    margin-right: 0;
    margin-left: 4px;
    transform: scaleX(-1);
  }

  :host([dir="rtl"]) .chevron.collapsed,
  :host([data-dir="rtl"]) .chevron.collapsed,
  :host-context([dir="rtl"]) .chevron.collapsed {
    transform: scaleX(-1) rotate(-90deg);
  }

  :host([dir="rtl"]) .heading-title,
  :host([data-dir="rtl"]) .heading-title,
  :host-context([dir="rtl"]) .heading-title {
    text-align: right;
  }

  :host([dir="rtl"]) .search-box,
  :host([data-dir="rtl"]) .search-box,
  :host-context([dir="rtl"]) .search-box {
    direction: rtl;
  }

  :host([dir="rtl"]) .search-icon,
  :host([data-dir="rtl"]) .search-icon,
  :host-context([dir="rtl"]) .search-icon {
    margin-right: 0;
    margin-left: 6px;
  }

  :host([dir="rtl"]) .tab-strip,
  :host([data-dir="rtl"]) .tab-strip,
  :host-context([dir="rtl"]) .tab-strip {
    direction: rtl;
  }

  /* ── Search Bar ── */
  .search-bar {
    padding: 8px 10px;
    background: #f3f3f3;
    border-bottom: 1px solid var(--docen-color-divider, #e5e5e5);
    flex: 0 0 auto;
    box-sizing: border-box;
  }

  .search-box {
    display: flex;
    align-items: center;
    background: #ffffff;
    border: 1px solid #c8c8c8;
    border-radius: 4px;
    padding: 0 6px;
    height: 28px;
    box-sizing: border-box;
    transition:
      border-color 0.15s ease,
      box-shadow 0.15s ease;
  }

  .search-box:focus-within {
    border-color: #0f6cbd;
    box-shadow: 0 0 0 1px #0f6cbd;
  }

  .search-icon {
    width: 14px;
    height: 14px;
    color: #606060;
    flex-shrink: 0;
    margin-right: 6px;
  }

  .search-input {
    flex: 1;
    min-width: 0;
    border: none;
    outline: none;
    background: transparent;
    font-size: 12px;
    color: inherit;
    padding: 0;
  }

  .clear-btn {
    border: none;
    background: transparent;
    cursor: pointer;
    font-size: 12px;
    color: #808080;
    padding: 2px 4px;
    border-radius: 2px;
    display: flex;
    align-items: center;
    justify-content: center;
  }

  .clear-btn:hover {
    color: #242424;
    background: #e0e0e0;
  }

  /* ── Tab Strip ── */
  .tab-strip {
    display: flex;
    background: #f8f8f8;
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    flex: 0 0 auto;
    padding: 0 6px;
    gap: 2px;
  }

  .nav-tab {
    padding: 7px 12px;
    border: none;
    background: transparent;
    cursor: pointer;
    font-size: 12px;
    font-weight: 500;
    color: #555555;
    border-bottom: 2px solid transparent;
    transition:
      color 0.15s ease,
      border-color 0.15s ease;
  }

  .nav-tab:hover {
    color: #101010;
    background: rgba(0, 0, 0, 0.03);
  }

  .nav-tab.active {
    color: #0f6cbd;
    border-bottom-color: #0f6cbd;
    font-weight: 600;
  }

  /* ── Content Container ── */
  .pane-content {
    flex: 1;
    min-height: 0;
    overflow-y: auto;
    overflow-x: hidden;
    position: relative;
  }

  .tab-panel {
    display: flex;
    flex-direction: column;
    min-height: 100%;
  }

  .tab-panel[hidden] {
    display: none !important;
  }

  /* ── Headings Tree ── */
  .tree-view {
    padding: 6px 0;
    outline: none;
  }

  .tree-item {
    display: flex;
    align-items: center;
    min-height: 28px;
    padding: 3px 8px;
    cursor: pointer;
    position: relative;
    border-left: 3px solid transparent;
    transition: background 0.1s ease;
    box-sizing: border-box;
  }

  .tree-item:hover {
    background: #f0f0f0;
  }

  .tree-item.active {
    background: #e5f1fb;
    border-left-color: #0f6cbd;
    font-weight: 600;
  }

  .tree-item.focused {
    outline: 1px dotted #0f6cbd;
  }

  .tree-item.drop-before {
    border-top: 2px solid #0f6cbd !important;
  }

  .tree-item.drop-after {
    border-bottom: 2px solid #0f6cbd !important;
  }

  .chevron {
    width: 16px;
    height: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
    border-radius: 2px;
    flex-shrink: 0;
    margin-right: 4px;
    transition: transform 0.15s ease;
    color: #606060;
  }

  .chevron.collapsed {
    transform: rotate(-90deg);
  }

  .chevron.hidden {
    visibility: hidden;
    pointer-events: none;
  }

  .chevron:hover {
    background: #d8d8d8;
  }

  .heading-title {
    flex: 1;
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    font-size: 12px;
  }

  mark.nav-highlight {
    background: #fff3a8;
    color: #101010;
    font-weight: 600;
    padding: 0 1px;
    border-radius: 2px;
  }

  /* ── Pages List ── */
  .pages-list {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 14px;
    padding: 12px 8px;
  }

  .page-card {
    display: flex;
    flex-direction: column;
    align-items: center;
    cursor: pointer;
    border-radius: 4px;
    padding: 6px;
    transition: background 0.15s ease;
  }

  .page-card:hover {
    background: #f0f0f0;
  }

  .page-card.active {
    outline: 2px solid #0f6cbd;
    background: #eef6fc;
  }

  .page-thumb {
    width: 90px;
    height: 120px;
    background: #ffffff;
    border: 1px solid #d1d1d1;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
    display: flex;
    flex-direction: column;
    padding: 8px;
    box-sizing: border-box;
    overflow: hidden;
    font-size: 9px;
    color: #888888;
  }

  .page-thumb img {
    width: 100%;
    height: 100%;
    object-fit: cover;
  }

  .page-label {
    margin-top: 4px;
    font-size: 11px;
    color: #555555;
  }

  /* ── Results List ── */
  .results-summary {
    padding: 8px 12px;
    font-size: 11px;
    color: #606060;
    border-bottom: 1px solid #ececec;
    background: #f9f9f9;
  }

  .results-list {
    display: flex;
    flex-direction: column;
  }

  .result-card {
    padding: 8px 12px;
    cursor: pointer;
    border-bottom: 1px solid #f2f2f2;
    transition: background 0.1s ease;
  }

  .result-card:hover {
    background: #f0f6fc;
  }

  .result-heading {
    font-size: 11px;
    font-weight: 600;
    color: #0f6cbd;
    margin-bottom: 3px;
  }

  .result-snippet {
    font-size: 12px;
    color: #333333;
    line-height: 1.4;
  }

  /* ── Empty State ── */
  .empty-state {
    padding: 28px 16px;
    text-align: center;
    color: #888888;
    font-size: 12px;
  }
`;

const template = html<DocenNavPane>`
  <div class="search-bar" part="search">
    <div class="search-box">
      <svg class="search-icon" viewBox="0 0 16 16" fill="currentColor">
        <path
          d="M6.5 2a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9ZM1 6.5a5.5 5.5 0 0 1 9.9-3.26l3.4-3.4a.75.75 0 0 1 1.06 1.06l-3.4 3.4A5.5 5.5 0 0 1 6.5 12 5.5 5.5 0 0 1 1 6.5Z"
        />
      </svg>
      <input
        type="text"
        part="search-input"
        class="search-input"
        placeholder="${(x) => x.searchPlaceholder}"
        :value="${(x) => x.searchQuery}"
        @input="${(x, c) => x.onSearchInput(c.event as InputEvent)}"
        @keydown="${(x, c) => x.onSearchKeyDown(c.event as KeyboardEvent)}"
        ${ref("searchInput")}
      />
      <button
        class="clear-btn"
        ?hidden="${(x) => !x.searchQuery}"
        @click="${(x) => x.clearSearch()}"
        title="Clear search"
        aria-label="Clear search"
      >
        ✕
      </button>
    </div>
  </div>

  <div class="tab-strip" role="tablist" part="tabs" ${ref("tabs")}>
    <button
      role="tab"
      id="nav-headings"
      class="nav-tab ${(x) => (x.tab === "headings" ? "active" : "")}"
      aria-selected="${(x) => (x.tab === "headings" ? "true" : "false")}"
      @click="${(x) => x.setTab("headings")}"
    >
      ${(x) => x.tHeadings}
    </button>
    <button
      role="tab"
      id="nav-pages"
      class="nav-tab ${(x) => (x.tab === "pages" ? "active" : "")}"
      aria-selected="${(x) => (x.tab === "pages" ? "true" : "false")}"
      @click="${(x) => x.setTab("pages")}"
    >
      ${(x) => x.tPages}
    </button>
    <button
      role="tab"
      id="nav-results"
      class="nav-tab ${(x) => (x.tab === "results" ? "active" : "")}"
      aria-selected="${(x) => (x.tab === "results" ? "true" : "false")}"
      @click="${(x) => x.setTab("results")}"
    >
      ${(x) => x.tResults}
    </button>
  </div>

  <div class="pane-content" part="content">
    <!-- Headings Tab -->
    <div class="tab-panel tab-headings" ?hidden="${(x) => x.tab !== "headings"}">
      <slot name="headings"></slot>
      <div
        class="tree-view"
        role="tree"
        aria-label="Headings"
        tabindex="0"
        @keydown="${(x, c) => x.onTreeKeyDown(c.event as KeyboardEvent)}"
        ${ref("treeContainer")}
      ></div>
    </div>

    <!-- Pages Tab -->
    <div class="tab-panel tab-pages" ?hidden="${(x) => x.tab !== "pages"}">
      <slot name="pages"></slot>
      <div class="empty-state" ?hidden="${(x) => x.visiblePages.length > 0}">
        No pages available.
      </div>
      <div class="pages-list" role="list">
        ${repeat(
          (x) => x.visiblePages,
          html<NavPageItem, DocenNavPane>`
            <div
              class="page-card ${(p, c) =>
                p.pageNumber === c.parent.activePageNumber ? "active" : ""}"
              role="listitem"
              @click="${(p, c) => c.parent.onPageClick(p)}"
            >
              <div class="page-thumb">
                ${(p) =>
                  p.thumbnailUrl
                    ? html`<img src="${p.thumbnailUrl}" alt="" />`
                    : html`<div class="page-snippet">${p.snippet ?? ""}</div>`}
              </div>
              <span class="page-label">${(p, c) => c.parent.formatPageLabel(p.pageNumber)}</span>
            </div>
          `,
        )}
      </div>
    </div>

    <!-- Results Tab -->
    <div class="tab-panel tab-results" ?hidden="${(x) => x.tab !== "results"}">
      <slot name="results"></slot>
      <div class="empty-state" ?hidden="${(x) => !!x.searchQuery.trim()}">
        Type a search term to find occurrences.
      </div>
      <div class="results-summary" ?hidden="${(x) => !x.searchQuery.trim()}">
        ${(x) => x.resultsSummaryText}
      </div>
      <div
        class="empty-state"
        ?hidden="${(x) => !x.searchQuery.trim() || x.searchResults.length > 0}"
      >
        No results found.
      </div>
      <div
        class="results-list"
        role="list"
        ?hidden="${(x) => !x.searchQuery.trim() || x.searchResults.length === 0}"
      >
        ${repeat(
          (x) => x.searchResults,
          html<NavSearchResultItem, DocenNavPane>`
            <div
              class="result-card"
              role="listitem"
              @click="${(res, c) => c.parent.onResultClick(res)}"
            >
              <div class="result-heading">${(res) => res.headingTitle}</div>
              <div class="result-snippet">
                ${(res) => res.before}<mark class="nav-highlight">${(res) => res.match}</mark
                >${(res) => res.after}
              </div>
            </div>
          `,
        )}
      </div>
    </div>
  </div>
`;

/**
 * `<docen-nav-pane>` — Office-style full navigation pane:
 * - 3 tabs: Headings (hierarchical tree), Pages, Results
 * - Live search filter at top filtering headings, pages, and displaying results with snippet context
 * - Heading drag-and-drop reorders entire heading section in document
 * - Two-way scroll sync between canvas and heading tree
 * - Accessible keyboard navigation (Arrow keys, Enter, Space) with role="tree"
 */
@customElement({ name: "docen-nav-pane", template, styles })
export class DocenNavPane extends FASTElement {
  @attr tab: "headings" | "pages" | "results" = "headings";
  @attr dir: "ltr" | "rtl" = "ltr";

  @observable searchQuery = "";
  @observable headings: NavHeadingItem[] = [];
  @observable filteredHeadings: NavHeadingItem[] = [];
  @observable visibleHeadings: NavHeadingItem[] = [];
  @observable pages: NavPageItem[] = [];
  @observable visiblePages: NavPageItem[] = [];
  @observable searchResults: NavSearchResultItem[] = [];
  @observable activeHeadingId: string | null = null;
  @observable focusedHeadingId: string | null = null;
  @observable activePageNumber = 1;

  @observable dropTargetId: string | null = null;
  @observable dropPosition: "before" | "after" | null = null;

  @observable searchInput?: HTMLInputElement;
  @observable tabs?: HTMLElement;
  @observable treeContainer?: HTMLElement;

  draggedHeading: NavHeadingItem | null = null;

  #editor: Editor | null = null;
  #scrollContainer: HTMLElement | null = null;
  #isScrollingProgrammatically = false;
  #unobserveLang?: () => void;

  get isRtl(): boolean {
    const d = this.dir || this.getAttribute("dir");
    return d === "rtl" || resolveDir(this) === "rtl";
  }

  dirChanged(): void {
    this.renderHeadingsTree();
  }

  #syncDir = (): void => {
    const d = this.dir || this.getAttribute("dir");
    if (d === "rtl" || d === "ltr") return;
    const nextDir = resolveDir(this);
    if (this.getAttribute("dir") !== nextDir) {
      this.setAttribute("dir", nextDir);
    }
  };

  #onTransaction = (): void => {
    this.parseHeadingsFromDoc();
  };

  #onScroll = (): void => {
    if (this.#isScrollingProgrammatically) return;
    this.#syncActiveHeadingFromScroll();
  };

  override connectedCallback(): void {
    super.connectedCallback();
    this.#syncDir();
    this.tabChanged("", this.tab || "headings");
    this.parseHeadingsFromDoc();
    this.#unobserveLang = observeLang(() => {
      this.#syncDir();
      this.applyFilter();
    });
  }

  override disconnectedCallback(): void {
    if (this.#editor) {
      this.#editor.off("transaction", this.#onTransaction);
    }
    if (this.#scrollContainer) {
      this.#scrollContainer.removeEventListener("scroll", this.#onScroll);
    }
    this.#unobserveLang?.();
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

  tabChanged(_prev: string, next: string): void {
    if (!next) return;
    this.tab = next as "headings" | "pages" | "results";
    this.$emit("navigation:tab", { tab: this.tab, source: this });
  }

  setTab(tab: "headings" | "pages" | "results"): void {
    this.tab = tab;
    this.setAttribute("tab", tab);
    this.$emit("navigation:tab", { tab, source: this });
  }

  // ── Localization ──

  get searchPlaceholder(): string {
    return t("nav.search", this) || "Search";
  }

  get tHeadings(): string {
    return t("nav.headings", this) || "Headings";
  }

  get tPages(): string {
    return t("nav.pages", this) || "Pages";
  }

  get tResults(): string {
    return t("nav.results", this) || "Results";
  }

  get resultsSummaryText(): string {
    const count = this.searchResults.length;
    return `${count} result${count === 1 ? "" : "s"} found`;
  }

  formatPageLabel(page: number): string {
    const tmpl = t("nav.page", this) || "Page {page}";
    return tmpl.replace("{page}", String(page));
  }

  // ── Editor & Scroll Container Binding ──

  bindEditor(editor: Editor): void {
    if (this.#editor) {
      this.#editor.off("transaction", this.#onTransaction);
    }
    this.#editor = editor;
    editor.on("transaction", this.#onTransaction);
    this.parseHeadingsFromDoc();
  }

  setEditor(editor: Editor): void {
    this.bindEditor(editor);
  }

  bindScrollContainer(container: HTMLElement): void {
    if (this.#scrollContainer) {
      this.#scrollContainer.removeEventListener("scroll", this.#onScroll);
    }
    this.#scrollContainer = container;
    container.addEventListener("scroll", this.#onScroll, { passive: true });
  }

  // ── Document Parsing ──

  parseHeadingsFromDoc(): void {
    if (!this.#editor) return;
    const doc = this.#editor.state.doc;
    const styles = (doc.attrs as { styles?: StylesOptions }).styles;
    const list: NavHeadingItem[] = [];
    let pos = 1;

    for (let i = 0; i < doc.childCount; i++) {
      const node = doc.child(i);
      const level = detectNodeHeadingLevel(node, styles);
      if (level != null && level > 0) {
        const id = (node.attrs.id as string) || `heading-${i}-${pos}`;
        const title = node.textContent.trim() || "(Untitled)";
        list.push({
          id,
          title,
          level,
          pos,
          nodeIndex: i,
          collapsed: false,
          hasChildren: false,
          hasMatch: false,
          partBefore: title,
          partMatch: "",
          partAfter: "",
        });
      }
      pos += node.nodeSize;
    }

    for (let i = 0; i < list.length; i++) {
      const item = list[i]!;
      const next = list[i + 1];
      item.hasChildren = !!next && next.level > item.level;
    }

    // Retain previous collapsed states
    const prevCollapsed = new Set(this.headings.filter((h) => h.collapsed).map((h) => h.id));
    for (const h of list) {
      if (prevCollapsed.has(h.id)) h.collapsed = true;
    }

    this.headings = list;
    if (!this.activeHeadingId && list.length > 0) {
      this.activeHeadingId = list[0]!.id;
    }

    // Default pages from document if pages not supplied
    if (this.pages.length === 0) {
      const pageCount = Math.max(1, Math.ceil(doc.textContent.length / 2500));
      this.pages = Array.from({ length: pageCount }, (_, i) => ({
        pageNumber: i + 1,
        active: i === 0,
        snippet: doc.textContent.slice(i * 2500, i * 2500 + 80),
      }));
    }

    this.applyFilter();
  }

  // ── Search & Filter Logic ──

  onSearchInput(event: InputEvent): void {
    const target = event.target as HTMLInputElement | null;
    this.searchQuery = target?.value ?? "";
    if (this.searchQuery.trim() && this.tab !== "results") {
      this.setTab("results");
    }
    this.$emit("navigation:search", { query: this.searchQuery, source: this });
    this.applyFilter();
  }

  onSearchKeyDown(event: KeyboardEvent): void {
    if (event.key === "Enter") {
      event.preventDefault();
      this.$emit("navigation:find", {
        direction: event.shiftKey ? "prev" : "next",
        source: this,
      });
    }
  }

  clearSearch(): void {
    this.searchQuery = "";
    if (this.searchInput) this.searchInput.value = "";
    this.$emit("navigation:search", { query: "", source: this });
    this.applyFilter();
  }

  setSearchQuery(query: string): void {
    this.searchQuery = query;
    if (this.searchInput) this.searchInput.value = query;
    this.applyFilter();
  }

  applyFilter(): void {
    const query = this.searchQuery.trim().toLowerCase();
    if (!query) {
      this.filteredHeadings = this.headings.map((h) => ({
        ...h,
        hasMatch: false,
        partBefore: h.title,
        partMatch: "",
        partAfter: "",
      }));
      this.updateSearchResults();
      this.updateVisibleHeadings();
      this.updateVisiblePages();
      return;
    }

    const matches = new Set<string>();
    const withParts = this.headings.map((h) => {
      const idx = h.title.toLowerCase().indexOf(query);
      if (idx >= 0) {
        matches.add(h.id);
        return {
          ...h,
          hasMatch: true,
          partBefore: h.title.slice(0, idx),
          partMatch: h.title.slice(idx, idx + query.length),
          partAfter: h.title.slice(idx + query.length),
        };
      }
      return {
        ...h,
        hasMatch: false,
        partBefore: h.title,
        partMatch: "",
        partAfter: "",
      };
    });

    // Keep matching items and all their hierarchical ancestors
    const keepIds = new Set<string>(matches);
    for (let i = 0; i < withParts.length; i++) {
      if (matches.has(withParts[i]!.id)) {
        let curLevel = withParts[i]!.level;
        for (let j = i - 1; j >= 0; j--) {
          if (withParts[j]!.level < curLevel) {
            keepIds.add(withParts[j]!.id);
            curLevel = withParts[j]!.level;
          }
        }
      }
    }

    this.filteredHeadings = withParts.filter((h) => keepIds.has(h.id));
    this.updateSearchResults();
    this.updateVisibleHeadings();
    this.updateVisiblePages();
  }

  updateVisibleHeadings(): void {
    const list: NavHeadingItem[] = [];
    const collapsedStack: number[] = [];

    for (const h of this.filteredHeadings) {
      while (collapsedStack.length > 0 && collapsedStack[collapsedStack.length - 1]! >= h.level) {
        collapsedStack.pop();
      }
      if (collapsedStack.length === 0) {
        list.push(h);
      }
      if (h.collapsed && h.hasChildren) {
        collapsedStack.push(h.level);
      }
    }
    this.visibleHeadings = list;
    this.renderHeadingsTree();
  }

  renderHeadingsTree(): void {
    const container = this.treeContainer;
    if (!container) return;
    container.replaceChildren();

    const list = this.visibleHeadings;
    if (list.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty-state";
      empty.textContent = this.searchQuery
        ? "No headings match your search."
        : "No headings found in document.";
      container.append(empty);
      return;
    }

    const isRtl = this.isRtl;
    for (const h of list) {
      const itemEl = document.createElement("div");
      itemEl.setAttribute("role", "treeitem");
      itemEl.className = this.getHeadingItemClass(h);
      if (isRtl) {
        itemEl.style.paddingRight = `${(h.level - 1) * 16 + 8}px`;
        itemEl.style.paddingLeft = "8px";
      } else {
        itemEl.style.paddingLeft = `${(h.level - 1) * 16 + 8}px`;
        itemEl.style.paddingRight = "8px";
      }
      itemEl.setAttribute("aria-level", String(h.level));
      if (h.hasChildren) {
        itemEl.setAttribute("aria-expanded", !h.collapsed ? "true" : "false");
      }
      itemEl.setAttribute("aria-selected", h.id === this.activeHeadingId ? "true" : "false");
      itemEl.tabIndex = h.id === this.effectiveFocusedId ? 0 : -1;
      itemEl.draggable = true;
      itemEl.dataset.headingId = h.id;

      // Chevron
      const chevron = document.createElement("span");
      chevron.className = `chevron ${!h.hasChildren ? "hidden" : h.collapsed ? "collapsed" : ""}`;
      chevron.innerHTML = `<svg width="8" height="8" viewBox="0 0 8 8" fill="currentColor"><path d="M1.5 1.5 L6.5 4 L1.5 6.5 Z"/></svg>`;
      chevron.addEventListener("click", (e) => {
        e.stopPropagation();
        this.toggleHeadingCollapse(h.id);
      });
      itemEl.append(chevron);

      // Title
      const titleSpan = document.createElement("span");
      titleSpan.className = "heading-title";
      titleSpan.title = h.title;
      if (h.hasMatch && h.partMatch) {
        if (h.partBefore) titleSpan.append(document.createTextNode(h.partBefore));
        const mark = document.createElement("mark");
        mark.className = "nav-highlight";
        mark.textContent = h.partMatch;
        titleSpan.append(mark);
        if (h.partAfter) titleSpan.append(document.createTextNode(h.partAfter));
      } else {
        titleSpan.textContent = h.title;
      }
      itemEl.append(titleSpan);

      // Drag & Click events
      itemEl.addEventListener("click", (e) => {
        this.onHeadingClick(h, e);
      });
      itemEl.addEventListener("dragstart", (e) => this.onDragStart(h, e));
      itemEl.addEventListener("dragover", (e) => this.onDragOver(h, e));
      itemEl.addEventListener("dragleave", (e) => this.onDragLeave(h, e));
      itemEl.addEventListener("drop", (e) => this.onDrop(h, e));

      container.append(itemEl);
    }
  }

  updateVisiblePages(): void {
    if (!this.searchQuery.trim()) {
      this.visiblePages = this.pages;
      return;
    }
    const query = this.searchQuery.toLowerCase();
    this.visiblePages = this.pages.filter(
      (p) => !p.snippet || p.snippet.toLowerCase().includes(query),
    );
  }

  updateSearchResults(): void {
    const query = this.searchQuery.trim().toLowerCase();
    if (!query || !this.#editor) {
      this.searchResults = [];
      return;
    }

    const results: NavSearchResultItem[] = [];
    const doc = this.#editor.state.doc;
    let currentHeading = "Document";

    doc.descendants((node, pos) => {
      if (node.isText && node.text) {
        const text = node.text;
        const lower = text.toLowerCase();
        let idx = 0;
        while ((idx = lower.indexOf(query, idx)) !== -1) {
          const from = pos + idx;
          const to = from + query.length;
          const start = Math.max(0, idx - 25);
          const end = Math.min(text.length, idx + query.length + 25);
          const before = (start > 0 ? "…" : "") + text.slice(start, idx);
          const match = text.slice(idx, idx + query.length);
          const after = text.slice(idx + query.length, end) + (end < text.length ? "…" : "");

          results.push({
            id: `res-${from}-${to}`,
            from,
            to,
            headingTitle: currentHeading,
            before,
            match,
            after,
          });
          idx += query.length;
        }
      } else if (node.type.name === "heading" || (node.attrs as { heading?: string })?.heading) {
        currentHeading = node.textContent || currentHeading;
      }
    });

    this.searchResults = results;
  }

  // ── Visible Computed Collections ──

  get effectiveFocusedId(): string | null {
    if (this.focusedHeadingId) return this.focusedHeadingId;
    if (this.activeHeadingId) return this.activeHeadingId;
    return this.visibleHeadings[0]?.id ?? null;
  }

  getHeadingItemClass(h: NavHeadingItem): string {
    const classes = ["tree-item"];
    if (h.id === this.activeHeadingId) classes.push("active");
    if (h.id === this.effectiveFocusedId) classes.push("focused");
    if (this.dropTargetId === h.id && this.dropPosition === "before") {
      classes.push("drop-before");
    }
    if (this.dropTargetId === h.id && this.dropPosition === "after") {
      classes.push("drop-after");
    }
    return classes.join(" ");
  }

  // ── Heading Tree Interaction & Reordering ──

  onChevronClick(heading: NavHeadingItem, event: MouseEvent): void {
    event.stopPropagation();
    this.toggleHeadingCollapse(heading.id);
  }

  toggleHeadingCollapse(id: string): void {
    this.headings = this.headings.map((h) => (h.id === id ? { ...h, collapsed: !h.collapsed } : h));
    this.applyFilter();
  }

  onHeadingClick(heading: NavHeadingItem, _event: MouseEvent): void {
    this.selectHeading(heading);
  }

  selectHeading(heading: NavHeadingItem): void {
    this.activeHeadingId = heading.id;
    this.focusedHeadingId = heading.id;
    this.$emit("nav:select-heading", {
      id: heading.id,
      pos: heading.pos,
      nodeIndex: heading.nodeIndex,
    });
    this.$emit("outline:select", {
      id: heading.id,
      pos: heading.pos,
    });

    if (this.#editor) {
      this.#isScrollingProgrammatically = true;
      try {
        this.#editor.commands.setTextSelection(heading.pos);
        this.#editor.commands.focus();
        try {
          const dom = (this.#editor as any).view?.nodeDOM?.(heading.pos) as HTMLElement | null;
          dom?.scrollIntoView?.({ behavior: "smooth", block: "start" });
        } catch {
          // View may not be attached in headless environments
        }
      } finally {
        setTimeout(() => {
          this.#isScrollingProgrammatically = false;
        }, 300);
      }
    } else if (this.#scrollContainer) {
      this.#isScrollingProgrammatically = true;
      try {
        const targetEl = this.#scrollContainer.querySelector(
          `[data-heading-id="${heading.id}"], #${heading.id}`,
        );
        targetEl?.scrollIntoView({ behavior: "smooth", block: "start" });
      } finally {
        setTimeout(() => {
          this.#isScrollingProgrammatically = false;
        }, 300);
      }
    }
  }

  // ── Drag and Drop Reordering (W5.1 & W5.3) ──

  onDragStart(heading: NavHeadingItem, event: DragEvent): void {
    this.draggedHeading = heading;
    if (event.dataTransfer) {
      event.dataTransfer.setData("application/x-docen-heading", heading.id);
      event.dataTransfer.effectAllowed = "move";
    }
  }

  onDragOver(heading: NavHeadingItem, event: DragEvent): void {
    if (!this.draggedHeading || this.draggedHeading.id === heading.id) return;
    event.preventDefault();
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = "move";
    }
    const targetEl = event.currentTarget as HTMLElement | null;
    if (!targetEl) return;
    const rect = targetEl.getBoundingClientRect();
    const isTopHalf = event.clientY - rect.top < rect.height / 2;
    this.dropTargetId = heading.id;
    this.dropPosition = isTopHalf ? "before" : "after";
  }

  onDragLeave(heading: NavHeadingItem, event: DragEvent): void {
    const related = event.relatedTarget as HTMLElement | null;
    const current = event.currentTarget as HTMLElement | null;
    if (current && related && current.contains(related)) return;
    if (this.dropTargetId === heading.id) {
      this.dropTargetId = null;
      this.dropPosition = null;
    }
  }

  onDrop(heading: NavHeadingItem, event: DragEvent): void {
    event.preventDefault();
    event.stopPropagation();
    const source = this.draggedHeading;
    const position = this.dropPosition ?? "after";
    this.dropTargetId = null;
    this.dropPosition = null;
    this.draggedHeading = null;
    if (!source || source.id === heading.id) return;
    this.reorderHeading(source.id, heading.id, position);
  }

  reorderHeading(sourceId: string, targetId: string, position: "before" | "after"): boolean {
    const source = this.headings.find((h) => h.id === sourceId);
    const target = this.headings.find((h) => h.id === targetId);
    if (!source || !target || source.id === target.id) return false;

    if (this.#editor) {
      const ok = moveHeadingSection(this.#editor, source.nodeIndex, target.nodeIndex, position);
      if (ok) {
        this.parseHeadingsFromDoc();
      }
    } else {
      const srcIdx = this.headings.findIndex((h) => h.id === sourceId);
      const tgtIdx = this.headings.findIndex((h) => h.id === targetId);
      if (srcIdx >= 0 && tgtIdx >= 0) {
        const item = this.headings[srcIdx]!;
        const updated = [...this.headings];
        updated.splice(srcIdx, 1);
        const insIdx =
          position === "before"
            ? srcIdx < tgtIdx
              ? tgtIdx - 1
              : tgtIdx
            : srcIdx < tgtIdx
              ? tgtIdx
              : tgtIdx + 1;
        updated.splice(insIdx, 0, item);
        this.headings = updated;
        this.applyFilter();
      }
    }

    this.$emit("nav:reorder-heading", { sourceId, targetId, position });
    return true;
  }

  // ── Keyboard Navigation in Tree ──

  onTreeKeyDown(event: KeyboardEvent): void {
    const list = this.visibleHeadings;
    if (list.length === 0) return;

    const curIndex = list.findIndex((h) => h.id === this.effectiveFocusedId);
    const current = curIndex >= 0 ? list[curIndex]! : list[0]!;

    switch (event.key) {
      case "ArrowDown": {
        event.preventDefault();
        const nextIdx = Math.min(list.length - 1, (curIndex >= 0 ? curIndex : 0) + 1);
        this.focusedHeadingId = list[nextIdx]!.id;
        this.#focusItemElement(this.focusedHeadingId);
        break;
      }
      case "ArrowUp": {
        event.preventDefault();
        const prevIdx = Math.max(0, (curIndex >= 0 ? curIndex : 0) - 1);
        this.focusedHeadingId = list[prevIdx]!.id;
        this.#focusItemElement(this.focusedHeadingId);
        break;
      }
      case "ArrowRight": {
        event.preventDefault();
        if (current.hasChildren) {
          if (current.collapsed) {
            this.toggleHeadingCollapse(current.id);
          } else if (curIndex + 1 < list.length && list[curIndex + 1]!.level > current.level) {
            this.focusedHeadingId = list[curIndex + 1]!.id;
            this.#focusItemElement(this.focusedHeadingId);
          }
        }
        break;
      }
      case "ArrowLeft": {
        event.preventDefault();
        if (current.hasChildren && !current.collapsed) {
          this.toggleHeadingCollapse(current.id);
        } else {
          for (let i = curIndex - 1; i >= 0; i--) {
            if (list[i]!.level < current.level) {
              this.focusedHeadingId = list[i]!.id;
              this.#focusItemElement(this.focusedHeadingId);
              break;
            }
          }
        }
        break;
      }
      case "Home": {
        event.preventDefault();
        this.focusedHeadingId = list[0]!.id;
        this.#focusItemElement(this.focusedHeadingId);
        break;
      }
      case "End": {
        event.preventDefault();
        this.focusedHeadingId = list[list.length - 1]!.id;
        this.#focusItemElement(this.focusedHeadingId);
        break;
      }
      case "Enter": {
        event.preventDefault();
        this.selectHeading(current);
        break;
      }
      case " ": {
        event.preventDefault();
        if (current.hasChildren) {
          this.toggleHeadingCollapse(current.id);
        }
        break;
      }
    }
  }

  #focusItemElement(id: string): void {
    const el = this.shadowRoot?.querySelector(`[data-heading-id="${id}"]`) as HTMLElement | null;
    el?.focus();
    el?.scrollIntoView({ block: "nearest" });
  }

  // ── Two-Way Scroll Sync (Document Canvas -> Tree) ──

  #syncActiveHeadingFromScroll(): void {
    if (!this.#scrollContainer || this.headings.length === 0) return;
    const containerRect = this.#scrollContainer.getBoundingClientRect();
    const headingEls = this.#scrollContainer.querySelectorAll(
      "h1, h2, h3, h4, h5, h6, [data-heading]",
    );

    let closestHeadingId: string | null = null;
    let minDistance = Infinity;

    for (let i = 0; i < headingEls.length; i++) {
      const el = headingEls[i] as HTMLElement;
      const rect = el.getBoundingClientRect();
      const distance = rect.top - containerRect.top;
      if (distance >= -40 && distance < minDistance) {
        minDistance = distance;
        const id = el.getAttribute("data-heading-id") || el.id;
        if (id) closestHeadingId = id;
      }
    }

    if (closestHeadingId && closestHeadingId !== this.activeHeadingId) {
      this.activeHeadingId = closestHeadingId;
      const itemEl = this.shadowRoot?.querySelector(`[data-heading-id="${closestHeadingId}"]`);
      itemEl?.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
  }

  // ── Page & Result Interactions ──

  onPageClick(page: NavPageItem): void {
    this.activePageNumber = page.pageNumber;
    this.$emit("nav-pages:select", { page: page.pageNumber });
    this.$emit("nav:page-select", { page: page.pageNumber });
  }

  onResultClick(res: NavSearchResultItem): void {
    this.$emit("nav:result-select", { id: res.id, from: res.from, to: res.to });
    if (this.#editor) {
      this.#editor.commands.setTextSelection({ from: res.from, to: res.to });
      this.#editor.commands.focus();
      try {
        const dom = (this.#editor as any).view?.nodeDOM?.(res.from) as HTMLElement | null;
        dom?.scrollIntoView?.({ behavior: "smooth", block: "center" });
      } catch {
        // View may not be attached in headless environments
      }
    }
  }
}

export default DocenNavPane;
