import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { BUILTIN_TEMPLATES } from "../../../document/templates";
import { observeLang, t } from "../../i18n/localize";
import { listValue, opt, type FluentListbox } from "./fluent-combo";

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(420px, 92vw);
  }
  .body {
    padding: 8px 4px 4px;
  }
  fluent-listbox.list {
    width: 100%;
    min-height: 156px;
    border: 1px solid var(--colorNeutralStroke1, #d1d1d1);
    border-radius: 4px;
    padding: 2px;
    font-size: 13px;
    box-shadow: none;
  }
`;

const template = html<DocenTemplateDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <fluent-listbox class="list" ${ref("listSel")}></fluent-listbox>
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
 * `<docen-template-dialog>` — Word's "New from Template" gallery: the built-in
 * templates from the data module (`document/templates.ts` — Blank, Report,
 * Letter) as a picker, each row carrying its own one-line description in the
 * option's description slot. Commit emits `template:create` `{ id }`; the host
 * builds the model JSON and instantiates it as a new document. The dialog never
 * builds documents itself.
 */
@customElement({ name: "docen-template-dialog", template, styles })
class DocenTemplateDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable listSel?: FluentListbox;
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
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Emit the picked template id and close (the host instantiates it). */
  commit(): void {
    const id = listValue(this.listSel);
    if (!id) return;
    this.$emit("template:create", { id });
    this.hide();
  }

  /** Build the option rows from the data module for the active locale (re-run
   *  on language change, keeping the picked id when it still exists). The
   *  description rides the option's `description` slot — the gallery reads
   *  self-contained rows, no selection-tracking side channel. */
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
    this.#renderOptions();
  }
}

export default DocenTemplateDialog;
