import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import {
  BLOCK_GALLERIES,
  DEFAULT_BLOCK_CATEGORY,
  DEFAULT_BLOCK_GALLERY,
  isDuplicateName,
} from "../../../document/building-blocks";
import { observeLang, t } from "../../i18n/localize";
import { listboxOf, opt, pick, pickedValue, type FluentDropdown } from "./fluent-combo";

/** The seed the host passes before show(): the captured selection preview +
 *  the current block names the duplicate check runs against. */
export interface QuickPartSeed {
  /** The selection's plain text (the dialog's preview line). */
  preview: string;
  /** Word's suggestion: the selection's own text. */
  suggestedName: string;
  /** Every existing block name — duplicates are rejected inline. */
  existingNames: readonly string[];
  gallery?: string;
  category?: string;
}

/** The dialog's OK detail — the host captures the content itself. */
export interface QuickPartValues {
  name: string;
  gallery: string;
  category: string;
  description: string;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog {
    --dialog-width: min(460px, 94vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .preview {
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    background: var(--colorNeutralBackground1, #fff);
    padding: 6px 10px;
    min-height: 18px;
    max-height: 72px;
    overflow: hidden;
    white-space: pre-wrap;
    color: var(--docen-color-secondary, #595959);
  }
  .row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .row > label {
    min-width: 88px;
  }
  .row fluent-dropdown,
  .row fluent-text-input {
    flex: 1;
    min-width: 0;
  }
  .row fluent-dropdown input {
    width: 100%;
    box-sizing: border-box;
  }
  .error {
    color: var(--colorPaletteRedForeground1, #c50f1f);
    min-height: 16px;
  }
  .error:empty {
    display: none;
  }
`;

const template = html<DocenQuickPartDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="preview" ${ref("previewEl")}></div>
      <div class="row">
        <label ${ref("nameLabel")}></label>
        <fluent-text-input ${ref("nameInput")} spellcheck="false"></fluent-text-input>
      </div>
      <div class="row">
        <label ${ref("galleryLabel")}></label>
        <fluent-dropdown type="combobox" appearance="outline" ${ref("gallerySel")}>
          <fluent-listbox popover="manual" tabindex="-1" ${ref("galleryList")}></fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            size="1"
            style="width:100%;box-sizing:border-box"
          />
        </fluent-dropdown>
      </div>
      <div class="row">
        <label ${ref("categoryLabel")}></label>
        <fluent-text-input ${ref("categoryInput")} spellcheck="false"></fluent-text-input>
      </div>
      <div class="row">
        <label ${ref("descriptionLabel")}></label>
        <fluent-text-input ${ref("descriptionInput")} spellcheck="false"></fluent-text-input>
      </div>
      <div class="error" ${ref("errorEl")} role="alert"></div>
    </div>
    <fluent-button
      slot="action"
      appearance="stealth"
      ${ref("cancelBtn")}
      @click="${(x) => x.hide()}"
    ></fluent-button>
    <fluent-button
      slot="action"
      appearance="accent"
      ${ref("okBtn")}
      @click="${(x) => x.onOk()}"
    ></fluent-button>
  </docen-dialog>
`;

/**
 * `<docen-quick-part-dialog>` — Word's "Create New Building Block" dialog
 * (Insert → Text → Quick Parts → Save Selection to Quick Part Gallery…).
 * Seeds from the captured selection preview plus the current block names;
 * duplicate/empty names are rejected inline with a message, and OK emits
 * `quick-part:save { name, gallery, category, description }` — the host owns
 * the content snapshot and persistence.
 */
@customElement({ name: "docen-quick-part-dialog", template, styles })
class DocenQuickPartDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable previewEl?: HTMLElement;
  @observable nameLabel?: HTMLElement;
  @observable nameInput?: HTMLElement & {
    value: string;
    focus?(): void;
    select?(): void;
  };
  @observable galleryLabel?: HTMLElement;
  @observable gallerySel?: FluentDropdown;
  @observable galleryList?: HTMLElement;
  @observable categoryLabel?: HTMLElement;
  @observable categoryInput?: HTMLInputElement & { value: string };
  @observable descriptionLabel?: HTMLElement;
  @observable descriptionInput?: HTMLInputElement & { value: string };
  @observable errorEl?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #existingNames: readonly string[] = [];
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => {
      this.#applyLabels();
      this.#buildGalleryOptions();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  /** Seed from the captured selection and open. */
  show(seed: QuickPartSeed): void {
    this.#existingNames = seed.existingNames;
    if (this.previewEl) this.previewEl.textContent = seed.preview;
    if (this.nameInput) this.nameInput.value = seed.suggestedName;
    if (this.categoryInput) this.categoryInput.value = seed.category ?? DEFAULT_BLOCK_CATEGORY;
    if (this.descriptionInput) this.descriptionInput.value = "";
    if (this.errorEl) this.errorEl.textContent = "";
    this.#buildGalleryOptions();
    pick(this.gallerySel, seed.gallery ?? DEFAULT_BLOCK_GALLERY);
    this.dialogEl?.show();
    requestAnimationFrame(() => {
      this.nameInput?.focus?.();
      this.nameInput?.select?.();
    });
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  readonly onOk = (): void => {
    const name = this.nameInput?.value.trim() ?? "";
    if (!name) {
      this.#error(t("quickPart.nameRequired", this));
      return;
    }
    if (isDuplicateName(this.#existingNames, name)) {
      this.#error(t("quickPart.duplicate", this));
      return;
    }
    const gallery = pickedValue(this.gallerySel) ?? DEFAULT_BLOCK_GALLERY;
    const category = this.categoryInput?.value.trim() || DEFAULT_BLOCK_CATEGORY;
    this.dispatchEvent(
      new CustomEvent("quick-part:save", {
        bubbles: true,
        composed: true,
        detail: {
          name,
          gallery,
          category,
          description: this.descriptionInput?.value.trim() ?? "",
        } satisfies QuickPartValues,
      }),
    );
    this.hide();
  };

  #error(message: string): void {
    if (this.errorEl) this.errorEl.textContent = message;
  }

  #buildGalleryOptions(): void {
    const list = listboxOf(this.gallerySel);
    if (!list) return;
    list.replaceChildren(
      ...BLOCK_GALLERIES.map((gallery) =>
        opt(t(`buildingBlock.gallery.${gallery}`, this), gallery),
      ),
    );
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("quickPart.title", this);
    if (this.nameLabel) this.nameLabel.textContent = t("quickPart.name", this);
    if (this.galleryLabel) this.galleryLabel.textContent = t("quickPart.gallery", this);
    if (this.categoryLabel) this.categoryLabel.textContent = t("quickPart.category", this);
    if (this.descriptionLabel) this.descriptionLabel.textContent = t("quickPart.description", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenQuickPartDialog;
