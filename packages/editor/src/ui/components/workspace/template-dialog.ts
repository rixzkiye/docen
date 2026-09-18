import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { BUILTIN_TEMPLATES } from "../../../document/templates";
import { observeLang, t } from "../../i18n/localize";
import { listValue, opt, type FluentListbox } from "./fluent-combo";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(460px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  fluent-listbox.list {
    width: 100%;
    min-height: 180px;
    border: 1px solid var(--colorNeutralStroke1, #d1d1d1);
    border-radius: 4px;
    padding: 2px;
    font-size: 13px;
    box-shadow: none;
  }
  .attach-section {
    padding-top: 10px;
    border-top: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .attach-row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .attach-filename {
    font-size: 12px;
    color: var(--colorNeutralForeground2, #616161);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    max-width: 250px;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 13px;
    cursor: pointer;
  }
  .attach-status {
    font-size: 12px;
    color: var(--colorPaletteGreenForeground1, #107c41);
    min-height: 16px;
  }
`;

const template = html<DocenTemplateDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <fluent-listbox class="list" ${ref("listSel")}></fluent-listbox>
      <input
        type="file"
        accept=".dotx,.docx"
        style="display:none"
        ${ref("fileInput")}
        @change="${(x, c) => x.onFileSelected(c.event)}"
      />
      <div class="attach-section">
        <div class="attach-row">
          <fluent-button
            appearance="neutral"
            ${ref("attachBtn")}
            @click="${(x) => x.onPickTemplate()}"
          ></fluent-button>
          <span class="attach-filename" ${ref("filenameEl")}></span>
        </div>
        <label class="check-field">
          <fluent-checkbox ${ref("autoUpdateBox")}></fluent-checkbox>
          <span ${ref("autoUpdateLabelEl")}></span>
        </label>
        <div class="attach-status" ${ref("statusEl")}></div>
      </div>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button
        appearance="accent"
        ${ref("createBtn")}
        @click="${(x) => x.commit()}"
      ></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-template-dialog>` — Word's "New from Template" gallery and template attachment:
 * provides the built-in templates (Blank, Report, Letter, Resume, Invoice) and
 * allows attaching external .dotx templates to import styles, numbering, docDefaults,
 * and DrawingML themes with "Automatically update document styles" (settings.linkStyles) support.
 */
@customElement({ name: "docen-template-dialog", template, styles })
class DocenTemplateDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable listSel?: FluentListbox;
  @observable fileInput?: HTMLInputElement;
  @observable attachBtn?: HTMLElement;
  @observable filenameEl?: HTMLElement;
  @observable autoUpdateBox?: HTMLElement & { checked?: boolean };
  @observable autoUpdateLabelEl?: HTMLElement;
  @observable statusEl?: HTMLElement;
  @observable createBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

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

  show(): void {
    this.#renderOptions();
    if (this.filenameEl) this.filenameEl.textContent = "";
    if (this.statusEl) this.statusEl.textContent = "";
    if (this.autoUpdateBox) this.autoUpdateBox.checked = true;
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  onPickTemplate(): void {
    this.fileInput?.click();
  }

  async onFileSelected(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    if (this.filenameEl) this.filenameEl.textContent = file.name;
    const arrayBuffer = await file.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const autoUpdateStyles = this.autoUpdateBox?.checked !== false;

    this.$emit("template:attach", {
      data,
      filename: file.name,
      autoUpdateStyles,
    });

    if (this.statusEl) {
      this.statusEl.textContent = t("templateDialog.attachedSuccess", this);
    }
    input.value = "";
  }

  /** Emit the picked template id and close (the host instantiates it). */
  commit(): void {
    const id = listValue(this.listSel);
    if (!id) return;
    this.$emit("template:create", { id });
    this.hide();
  }

  /** Build the option rows from the data module for the active locale */
  #renderOptions(): void {
    if (!this.listSel) return;
    const previous = listValue(this.listSel) ?? BUILTIN_TEMPLATES[0]?.id;
    this.listSel.replaceChildren(
      ...BUILTIN_TEMPLATES.map((tpl) => {
        const option = opt(t(tpl.nameKey, this), tpl.id);
        const description = document.createElement("span");
        description.slot = "description";
        description.textContent = t(tpl.descriptionKey, this);
        option.append(description);
        return option;
      }),
    );
    const selected = [
      ...this.listSel.querySelectorAll<HTMLElement & { selected?: boolean }>("fluent-option"),
    ].find((option) => option.getAttribute("value") === previous);
    if (selected) selected.selected = true;
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("templateDialog.title", this);
    if (this.createBtn) this.createBtn.textContent = t("templateDialog.create", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    if (this.attachBtn) this.attachBtn.textContent = t("templateDialog.attachTemplate", this);
    if (this.autoUpdateLabelEl) {
      this.autoUpdateLabelEl.textContent = t("templateDialog.autoUpdateStyles", this);
    }
    this.#renderOptions();
  }
}

export default DocenTemplateDialog;
