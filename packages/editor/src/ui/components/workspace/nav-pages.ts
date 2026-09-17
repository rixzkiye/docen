import { FASTElement, css, customElement, html, observable, repeat } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface NavPageItem {
  pageNumber: number;
  active?: boolean;
  thumbnailUrl?: string | null;
}

const styles = css`
  :host {
    display: block;
    padding: 8px;
    box-sizing: border-box;
  }
  .pages-list {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 12px;
  }
  .page-card {
    display: flex;
    flex-direction: column;
    align-items: center;
    cursor: pointer;
    border-radius: 4px;
    padding: 4px;
    transition: background 0.15s ease;
  }
  .page-card:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.04));
  }
  .page-card.active {
    outline: 2px solid var(--docen-color-accent, #0f6cbd);
  }
  .page-thumb {
    width: 90px;
    height: 120px;
    background: #ffffff;
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
    display: flex;
    flex-direction: column;
    padding: 8px;
    box-sizing: border-box;
    gap: 4px;
    overflow: hidden;
  }
  .thumb-img {
    width: 100%;
    height: 100%;
    object-fit: contain;
    background: #ffffff;
    border-radius: 2px;
    display: block;
  }
  .thumb-line {
    height: 3px;
    background: #e0e0e0;
    border-radius: 1px;
    width: 100%;
  }
  .thumb-line.short {
    width: 60%;
  }
  .page-label {
    margin-top: 4px;
    font-size: 11px;
    color: var(--docen-color-foreground-secondary, #616161);
  }
`;

const template = html<DocenNavPages>`
  <div class="pages-list" part="pages-list">
    ${repeat(
      (x) => x.pages,
      html<NavPageItem, DocenNavPages>`
        <div
          class="page-card ${(x) => (x.active ? "active" : "")}"
          @click="${(x, c) => c.parent.onPageClick(x.pageNumber)}"
        >
          <div class="page-thumb">
            ${(x) =>
              x.thumbnailUrl
                ? html`<img class="thumb-img" src="${x.thumbnailUrl}" alt="Page thumbnail" />`
                : html`
                    <div class="thumb-line"></div>
                    <div class="thumb-line"></div>
                    <div class="thumb-line short"></div>
                    <div class="thumb-line"></div>
                    <div class="thumb-line short"></div>
                  `}
          </div>
          <span class="page-label">${(x, c) => c.parent.formatPageLabel(x.pageNumber)}</span>
        </div>
      `,
    )}
  </div>
`;

/**
 * `<docen-nav-pages>` — Navigation pane Pages tab displaying thumbnails for
 * document pages. Clicking a thumbnail emits `nav-pages:select { page }` to
 * scroll the viewport to that page.
 */
@customElement({ name: "docen-nav-pages", template, styles })
export class DocenNavPages extends FASTElement {
  @observable pages: NavPageItem[] = [{ pageNumber: 1, active: true }];
  @observable activePage = 1;

  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => {
      this.pages = [...this.pages];
    });
  }

  formatPageLabel(pageNumber: number): string {
    const tpl = t("nav.page", this);
    return tpl.includes("{page}")
      ? tpl.replace("{page}", String(pageNumber))
      : `${tpl} ${pageNumber}`;
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  setPageCount(count: number, current = 1, thumbnails?: (string | null)[]): void {
    const list: NavPageItem[] = [];
    const total = Math.max(1, count);
    for (let i = 1; i <= total; i++) {
      list.push({
        pageNumber: i,
        active: i === current,
        thumbnailUrl: thumbnails?.[i - 1] ?? this.pages[i - 1]?.thumbnailUrl,
      });
    }
    this.pages = list;
    this.activePage = current;
  }

  onPageClick(pageNumber: number): void {
    this.activePage = pageNumber;
    this.pages = this.pages.map((p) => ({
      ...p,
      active: p.pageNumber === pageNumber,
    }));
    this.$emit("nav-pages:select", { page: pageNumber });
  }
}
