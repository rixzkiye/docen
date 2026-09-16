import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { BuildingBlock } from "../../../document/building-blocks";
import { isDuplicateName, sortBlocks } from "../../../document/building-blocks";
import { observeLang, t } from "../../i18n/localize";

/** The seed the host passes before show(). */
export interface BuildingBlocksSeed {
  blocks: readonly BuildingBlock[];
  /** False in Viewing mode — rename/delete/insert are disabled. */
  editable: boolean;
}

interface FluentTextInputLike extends HTMLElement {
  value: string;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog {
    --dialog-width: min(620px, 94vw);
  }
  .bb-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    font-size: 13px;
  }
  .bb-head,
  .bb-row {
    display: grid;
    grid-template-columns: minmax(0, 1.6fr) minmax(0, 1fr) minmax(0, 1fr) auto auto;
    align-items: center;
    gap: 8px;
  }
  .bb-head {
    font-weight: 600;
    color: var(--docen-color-secondary, #595959);
    border-block-end: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    padding-block-end: 4px;
  }
  .bb-list {
    display: flex;
    flex-direction: column;
    gap: 4px;
    max-height: 300px;
    overflow-y: auto;
  }
  .bb-row .name {
    min-width: 0;
  }
  .bb-row.bb-editable .name:hover {
    border-color: var(--colorNeutralStroke1, #c7c7c7);
  }
  .bb-empty {
    padding: 12px;
    text-align: center;
    color: var(--docen-color-secondary, #595959);
  }
  .error {
    color: var(--colorPaletteRedForeground1, #c50f1f);
    min-height: 16px;
  }
  .error:empty {
    display: none;
  }
`;

const template = html<DocenBuildingBlocksDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="bb-body">
      <div class="bb-head">
        <span ${ref("nameHeadEl")}></span>
        <span ${ref("galleryHeadEl")}></span>
        <span ${ref("categoryHeadEl")}></span>
        <span></span>
        <span></span>
      </div>
      <div
        class="bb-list"
        ${ref("listEl")}
        @click="${(x, c) => x.onListClick(c.event)}"
        @change="${(x, c) => x.onListChange(c.event)}"
      ></div>
      <div class="error" ${ref("errorEl")} role="alert"></div>
    </div>
    <fluent-button
      slot="action"
      appearance="accent"
      ${ref("closeBtn")}
      @click="${(x) => x.hide()}"
    ></fluent-button>
  </docen-dialog>
`;

/**
 * `<docen-building-blocks-dialog>` — Word's Building Blocks Organizer: every
 * building block with its gallery/category, plus per-row Insert, inline Rename
 * (the name field) and Delete. The host seeds via `show({ blocks, editable })`
 * and applies the events — `building-blocks:insert { id }`,
 * `building-blocks:rename { id, name }`, `building-blocks:delete { id }`.
 * Rename duplicates are rejected inline; delete/rename ride the host's one
 * document transaction, so the document's undo restores them.
 */
@customElement({ name: "docen-building-blocks-dialog", template, styles })
class DocenBuildingBlocksDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable nameHeadEl?: HTMLElement;
  @observable galleryHeadEl?: HTMLElement;
  @observable categoryHeadEl?: HTMLElement;
  @observable listEl?: HTMLElement;
  @observable errorEl?: HTMLElement;
  @observable closeBtn?: HTMLElement;

  #blocks: BuildingBlock[] = [];
  #editable = true;
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => {
      this.#applyLabels();
      this.#render();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(seed: BuildingBlocksSeed): void {
    this.#blocks = seed.blocks.map((block) => ({ ...block }));
    this.#editable = seed.editable;
    if (this.errorEl) this.errorEl.textContent = "";
    this.#render();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Template-bound (public — the FAST template cannot reach a #private
   *  member): the list's delegated click — Insert / Delete buttons. */
  onListClick(event: Event): void {
    const button = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-role]");
    if (!button) return;
    const role = button.dataset.role;
    if (role === "insert") {
      const id = button.dataset.id;
      if (id) {
        this.dispatchEvent(
          new CustomEvent("building-blocks:insert", {
            bubbles: true,
            composed: true,
            detail: { id },
          }),
        );
      }
      return;
    }
    if (role === "delete" && this.#editable) {
      const id = button.dataset.id;
      if (!id) return;
      this.#blocks = this.#blocks.filter((block) => block.id !== id);
      this.#clearError();
      this.#render();
      this.dispatchEvent(
        new CustomEvent("building-blocks:delete", {
          bubbles: true,
          composed: true,
          detail: { id },
        }),
      );
    }
  }

  /** Template-bound: a name field committed (blur/Enter) — rename with an
   *  inline duplicate check (a rejected edit reverts to the stored name). */
  onListChange(event: Event): void {
    const input = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-role='name']");
    if (!input || !this.#editable) return;
    const id = input.dataset.id;
    const index = this.#blocks.findIndex((block) => block.id === id);
    if (!id || index < 0) return;
    const current = this.#blocks[index]!;
    const value = (input as FluentTextInputLike).value.trim();
    if (!value) {
      (input as FluentTextInputLike).value = current.name;
      this.#error(t("buildingBlocks.nameRequired", this));
      return;
    }
    if (value === current.name) return;
    const others = this.#blocks.filter((block) => block.id !== id).map((block) => block.name);
    if (isDuplicateName(others, value)) {
      (input as FluentTextInputLike).value = current.name;
      this.#error(t("buildingBlocks.duplicate", this));
      return;
    }
    current.name = value;
    this.#clearError();
    this.dispatchEvent(
      new CustomEvent("building-blocks:rename", {
        bubbles: true,
        composed: true,
        detail: { id, name: value },
      }),
    );
  }

  #error(message: string): void {
    if (this.errorEl) this.errorEl.textContent = message;
  }

  #clearError(): void {
    if (this.errorEl) this.errorEl.textContent = "";
  }

  #render(): void {
    const list = this.listEl;
    if (!list) return;
    const blocks = sortBlocks(this.#blocks);
    if (blocks.length === 0) {
      const empty = document.createElement("div");
      empty.className = "bb-empty";
      empty.textContent = t("buildingBlocks.empty", this);
      list.replaceChildren(empty);
      return;
    }
    const rows = blocks.map((block) => {
      const row = document.createElement("div");
      row.className = `bb-row${this.#editable ? " bb-editable" : ""}`;

      const name: HTMLElement = this.#editable
        ? document.createElement("fluent-text-input")
        : document.createElement("span");
      name.classList.add("name");
      name.setAttribute("data-role", "name");
      name.setAttribute("data-id", block.id);
      if (this.#editable) (name as FluentTextInputLike).value = block.name;
      else name.textContent = block.name;

      const gallery = document.createElement("span");
      gallery.textContent = t(`buildingBlock.gallery.${block.gallery}`, this);
      const category = document.createElement("span");
      category.textContent = block.category;

      const insert = document.createElement("fluent-button");
      insert.setAttribute("appearance", "neutral");
      insert.setAttribute("data-role", "insert");
      insert.setAttribute("data-id", block.id);
      insert.textContent = t("buildingBlocks.insert", this);
      if (!this.#editable) insert.setAttribute("disabled", "");

      const remove = document.createElement("fluent-button");
      remove.setAttribute("appearance", "stealth");
      remove.setAttribute("data-role", "delete");
      remove.setAttribute("data-id", block.id);
      remove.textContent = t("buildingBlocks.delete", this);
      if (!this.#editable) remove.setAttribute("disabled", "");

      row.append(name, gallery, category, insert, remove);
      return row;
    });
    list.replaceChildren(...rows);
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("buildingBlocks.title", this);
    if (this.nameHeadEl) this.nameHeadEl.textContent = t("buildingBlocks.name", this);
    if (this.galleryHeadEl) this.galleryHeadEl.textContent = t("buildingBlocks.gallery", this);
    if (this.categoryHeadEl) this.categoryHeadEl.textContent = t("buildingBlocks.category", this);
    if (this.closeBtn) this.closeBtn.textContent = t("buildingBlocks.close", this);
  }
}

export default DocenBuildingBlocksDialog;
