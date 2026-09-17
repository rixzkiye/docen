import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface BookmarkItem {
  name: string;
  from: number;
  to: number;
  id?: number;
}

type FluentTextInput = HTMLElement & { value: string };

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(420px, 92vw);
  }
  .bm-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .field {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .field > label {
    font-weight: 500;
  }
  fluent-text-input {
    width: 100%;
  }
  .bm-list-container {
    height: 160px;
    border: 1px solid var(--colorNeutralStroke2, #d1d1d1);
    border-radius: 4px;
    overflow-y: auto;
    background: var(--colorNeutralBackground1, #fff);
    display: flex;
    flex-direction: column;
  }
  .bm-item {
    padding: 4px 8px;
    cursor: pointer;
    user-select: none;
    font-size: 13px;
  }
  .bm-item:hover {
    background: var(--colorNeutralBackground1Hover, #f0f0f0);
  }
  .bm-item.selected {
    background: var(--colorBrandBackground2, #e0eafc);
    font-weight: 600;
  }
  .options-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    padding-top: 4px;
  }
  .sort-group {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .radio-label,
  .check-label {
    display: flex;
    align-items: center;
    gap: 4px;
    cursor: pointer;
    font-size: 12px;
  }
`;

const template = html<DocenBookmarkDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="bm-body">
      <div class="field">
        <label ${ref("nameLabel")}></label>
        <fluent-text-input
          ${ref("nameInput")}
          type="text"
          @input="${(x) => x.onInput()}"
          @keydown="${(x, c) => x.onKeydown(c.event as KeyboardEvent)}"
        ></fluent-text-input>
      </div>

      <div class="bm-list-container" ${ref("listContainer")}></div>

      <div class="options-row">
        <div class="sort-group">
          <span ${ref("sortLabel")}></span>
          <label class="radio-label">
            <input
              type="radio"
              name="bm-sort"
              value="name"
              checked
              ${ref("sortNameRadio")}
              @change="${(x) => x.setSort("name")}"
            />
            <span ${ref("sortNameLabel")}></span>
          </label>
          <label class="radio-label">
            <input
              type="radio"
              name="bm-sort"
              value="location"
              ${ref("sortLocRadio")}
              @change="${(x) => x.setSort("location")}"
            />
            <span ${ref("sortLocLabel")}></span>
          </label>
        </div>
        <label class="check-label">
          <input type="checkbox" ${ref("hiddenCheck")} @change="${(x) => x.toggleHidden()}" />
          <span ${ref("hiddenLabel")}></span>
        </label>
      </div>
    </div>

    <div slot="action">
      <fluent-button
        appearance="accent"
        ${ref("addBtn")}
        @click="${(x) => x.onAdd()}"
      ></fluent-button>
      <fluent-button ${ref("deleteBtn")} @click="${(x) => x.onDelete()}"></fluent-button>
      <fluent-button ${ref("gotoBtn")} @click="${(x) => x.onGoTo()}"></fluent-button>
      <fluent-button ${ref("closeBtn")} @click="${(x) => x.hide()}"></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-bookmark-dialog>` — Word's Bookmark dialog.
 * Lists bookmarks with Name / Location sort, hidden bookmarks filter,
 * and Add / Delete / Go To actions.
 */
@customElement({ name: "docen-bookmark-dialog", template, styles })
export class DocenBookmarkDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable nameLabel?: HTMLElement;
  @observable nameInput?: FluentTextInput;
  @observable listContainer?: HTMLElement;
  @observable sortLabel?: HTMLElement;
  @observable sortNameRadio?: HTMLInputElement;
  @observable sortNameLabel?: HTMLElement;
  @observable sortLocRadio?: HTMLInputElement;
  @observable sortLocLabel?: HTMLElement;
  @observable hiddenCheck?: HTMLInputElement;
  @observable hiddenLabel?: HTMLElement;
  @observable addBtn?: HTMLElement & { disabled?: boolean };
  @observable deleteBtn?: HTMLElement & { disabled?: boolean };
  @observable gotoBtn?: HTMLElement & { disabled?: boolean };
  @observable closeBtn?: HTMLElement;

  @observable bookmarks: BookmarkItem[] = [];
  @observable selectedName = "";
  @observable sortMode: "name" | "location" = "name";
  @observable showHidden = false;

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

  show(bookmarks: BookmarkItem[] = []): void {
    this.bookmarks = [...bookmarks];
    this.selectedName = "";
    if (this.nameInput) this.nameInput.value = "";
    this.#updateButtons();
    this.#renderList();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  setSort(mode: "name" | "location"): void {
    this.sortMode = mode;
    this.#renderList();
  }

  toggleHidden(): void {
    this.showHidden = !!this.hiddenCheck?.checked;
    this.#renderList();
  }

  onInput(): void {
    this.selectedName = this.nameInput?.value.trim() ?? "";
    this.#updateButtons();
    this.#highlightSelectedInList();
  }

  onKeydown(e: KeyboardEvent): void {
    if (e.key === "Enter") {
      e.preventDefault();
      if (!this.addBtn?.disabled) this.onAdd();
    }
  }

  onSelect(bm: BookmarkItem): void {
    this.selectedName = bm.name;
    if (this.nameInput) this.nameInput.value = bm.name;
    this.#updateButtons();
    this.#highlightSelectedInList();
  }

  onAdd(): void {
    const name = this.nameInput?.value.trim() ?? "";
    if (!this.#isValidName(name)) return;
    this.$emit("bookmark:add", { name });
    this.hide();
  }

  onDelete(): void {
    if (!this.selectedName) return;
    this.$emit("bookmark:delete", { name: this.selectedName });
    this.bookmarks = this.bookmarks.filter((b) => b.name !== this.selectedName);
    this.selectedName = "";
    if (this.nameInput) this.nameInput.value = "";
    this.#updateButtons();
    this.#renderList();
  }

  onGoTo(): void {
    if (!this.selectedName) return;
    const bm = this.bookmarks.find((b) => b.name === this.selectedName);
    if (!bm) return;
    this.$emit("bookmark:goto", { name: bm.name, from: bm.from, to: bm.to });
  }

  #isValidName(name: string): boolean {
    return /^[A-Za-z一-鿿぀-ヿ_][^\s]*$/.test(name) && name.length <= 40;
  }

  #updateButtons(): void {
    const name = this.nameInput?.value.trim() ?? "";
    const isValid = this.#isValidName(name);
    const exists = this.bookmarks.some((b) => b.name === this.selectedName);

    if (this.addBtn) this.addBtn.disabled = !isValid;
    if (this.deleteBtn) this.deleteBtn.disabled = !exists;
    if (this.gotoBtn) this.gotoBtn.disabled = !exists;
  }

  #filteredBookmarks(): BookmarkItem[] {
    let list = this.bookmarks;
    if (!this.showHidden) {
      list = list.filter((b) => !b.name.startsWith("_"));
    }
    if (this.sortMode === "name") {
      return [...list].sort((a, b) => a.name.localeCompare(b.name));
    }
    return [...list].sort((a, b) => a.from - b.from);
  }

  #renderList(): void {
    if (!this.listContainer) return;
    this.listContainer.innerHTML = "";
    const items = this.#filteredBookmarks();
    for (const item of items) {
      const el = document.createElement("div");
      el.className = "bm-item" + (item.name === this.selectedName ? " selected" : "");
      el.textContent = item.name;
      el.addEventListener("click", () => this.onSelect(item));
      this.listContainer.appendChild(el);
    }
  }

  #highlightSelectedInList(): void {
    if (!this.listContainer) return;
    const children = this.listContainer.querySelectorAll<HTMLElement>(".bm-item");
    children.forEach((child) => {
      child.classList.toggle("selected", child.textContent === this.selectedName);
    });
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("bookmark.dialog-title", this);
    if (this.nameLabel) this.nameLabel.textContent = t("bookmark.name", this);
    if (this.sortLabel) this.sortLabel.textContent = t("bookmark.sort-by", this);
    if (this.sortNameLabel) this.sortNameLabel.textContent = t("bookmark.sort-name", this);
    if (this.sortLocLabel) this.sortLocLabel.textContent = t("bookmark.sort-location", this);
    if (this.hiddenLabel) this.hiddenLabel.textContent = t("bookmark.hidden", this);
    if (this.addBtn) this.addBtn.textContent = t("bookmark.add", this);
    if (this.deleteBtn) this.deleteBtn.textContent = t("bookmark.delete", this);
    if (this.gotoBtn) this.gotoBtn.textContent = t("bookmark.goto", this);
    if (this.closeBtn) this.closeBtn.textContent = t("bookmark.close", this);
  }
}
