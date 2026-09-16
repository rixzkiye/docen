import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { AutocorrectReplacement } from "../../../document/settings";
import { observeLang, t } from "../../i18n/localize";

/** The dialog's seed/commit shape — the same rule toggles the store persists
 *  plus the user replacement table and exceptions list. */
export interface AutocorrectDialogValues {
  smartQuotes: boolean;
  emDash: boolean;
  ellipsis: boolean;
  hyperlinkAutoformat: boolean;
  capitalizeFirstLetter: boolean;
  ordinalSuperscript: boolean;
  replacements: readonly AutocorrectReplacement[];
  exceptions: readonly string[];
}

interface TextInputLike extends HTMLElement {
  value: string;
}

const styles = css`
  :host {
    display: contents;
  }
  /* --dialog-width is the docen-dialog sizing channel (::part width rules
     never reach FAST's fixed-positioned native surface). 540px fits the
     two-column replacement table with its delete buttons. */
  docen-dialog {
    --dialog-width: min(540px, 94vw);
  }
  .ac-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    font-size: 13px;
  }
  .ac-heading {
    font-weight: 600;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
  .replace-table {
    display: flex;
    flex-direction: column;
    gap: 4px;
    max-height: 216px;
    overflow-y: auto;
  }
  .replace-row {
    display: grid;
    grid-template-columns: 1fr auto 1fr auto;
    align-items: center;
    gap: 6px;
  }
  .replace-row .arrow {
    opacity: 0.6;
  }
  .replace-head {
    display: grid;
    grid-template-columns: 1fr auto 1fr auto;
    gap: 6px;
    opacity: 0.8;
  }
  textarea {
    width: 100%;
    box-sizing: border-box;
    min-height: 64px;
    resize: vertical;
    font: inherit;
  }
`;

const template = html<DocenAutocorrectDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="ac-body">
      <div class="ac-heading" ${ref("rulesHeadingEl")}></div>
      <label class="check-field">
        <fluent-checkbox ${ref("smartQuotesBox")}></fluent-checkbox>
        <span ${ref("smartQuotesLabelEl")}></span>
      </label>
      <label class="check-field">
        <fluent-checkbox ${ref("dashBox")}></fluent-checkbox>
        <span ${ref("dashLabelEl")}></span>
      </label>
      <label class="check-field">
        <fluent-checkbox ${ref("ellipsisBox")}></fluent-checkbox>
        <span ${ref("ellipsisLabelEl")}></span>
      </label>
      <label class="check-field">
        <fluent-checkbox ${ref("ordinalBox")}></fluent-checkbox>
        <span ${ref("ordinalLabelEl")}></span>
      </label>
      <label class="check-field">
        <fluent-checkbox ${ref("capitalizeBox")}></fluent-checkbox>
        <span ${ref("capitalizeLabelEl")}></span>
      </label>
      <label class="check-field">
        <fluent-checkbox ${ref("hyperlinkBox")}></fluent-checkbox>
        <span ${ref("hyperlinkLabelEl")}></span>
      </label>
      <div class="ac-heading" ${ref("replacementsHeadingEl")}></div>
      <div class="replace-head">
        <span ${ref("fromHeadEl")}></span>
        <span></span>
        <span ${ref("toHeadEl")}></span>
        <span></span>
      </div>
      <div
        class="replace-table"
        ${ref("tableEl")}
        @click="${(x, c) => x.onTableClick(c.event)}"
      ></div>
      <fluent-button
        appearance="neutral"
        ${ref("addBtn")}
        @click="${(x) => x.addRow()}"
      ></fluent-button>
      <div class="ac-heading" ${ref("exceptionsHeadingEl")}></div>
      <textarea ${ref("exceptionsEl")}></textarea>
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

/** A local editable row (the DOM owns the in-progress text; `#rows` is the
 *  collected snapshot). */
interface Row {
  from: string;
  to: string;
}

/**
 * `<docen-autocorrect-dialog>` — Word's AutoCorrect Options: the "Replace as
 * you type" rule toggles, the user replacement table (add/edit/delete) and
 * the exceptions list. The host (`<docen-document>`, opened from the Options
 * dialog's Proofing section) seeds the effective values via `show(values)`
 * and listens for `autocorrect:ok { …toggles, replacements, exceptions }`;
 * the dialog commits nothing itself. Cancel / Esc just close.
 */
@customElement({ name: "docen-autocorrect-dialog", template, styles })
class DocenAutocorrectDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable rulesHeadingEl?: HTMLElement;
  @observable smartQuotesBox?: HTMLElement & { checked?: boolean };
  @observable smartQuotesLabelEl?: HTMLElement;
  @observable dashBox?: HTMLElement & { checked?: boolean };
  @observable dashLabelEl?: HTMLElement;
  @observable ellipsisBox?: HTMLElement & { checked?: boolean };
  @observable ellipsisLabelEl?: HTMLElement;
  @observable ordinalBox?: HTMLElement & { checked?: boolean };
  @observable ordinalLabelEl?: HTMLElement;
  @observable capitalizeBox?: HTMLElement & { checked?: boolean };
  @observable capitalizeLabelEl?: HTMLElement;
  @observable hyperlinkBox?: HTMLElement & { checked?: boolean };
  @observable hyperlinkLabelEl?: HTMLElement;
  @observable replacementsHeadingEl?: HTMLElement;
  @observable fromHeadEl?: HTMLElement;
  @observable toHeadEl?: HTMLElement;
  @observable tableEl?: HTMLElement;
  @observable addBtn?: HTMLElement;
  @observable exceptionsHeadingEl?: HTMLElement;
  @observable exceptionsEl?: HTMLTextAreaElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #rows: Row[] = [];
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => {
      this.#collect();
      this.#applyLabels();
      this.#renderRows();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  /** Seed the dialog from the effective settings and open it. */
  show(values: AutocorrectDialogValues): void {
    if (this.smartQuotesBox) this.smartQuotesBox.checked = values.smartQuotes;
    if (this.dashBox) this.dashBox.checked = values.emDash;
    if (this.ellipsisBox) this.ellipsisBox.checked = values.ellipsis;
    if (this.ordinalBox) this.ordinalBox.checked = values.ordinalSuperscript;
    if (this.capitalizeBox) this.capitalizeBox.checked = values.capitalizeFirstLetter;
    if (this.hyperlinkBox) this.hyperlinkBox.checked = values.hyperlinkAutoformat;
    this.#rows = values.replacements.map(({ from, to }) => ({ from, to }));
    this.#renderRows();
    if (this.exceptionsEl) this.exceptionsEl.value = values.exceptions.join("\n");
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  readonly onOk = (): void => {
    this.#collect();
    this.dispatchEvent(
      new CustomEvent("autocorrect:ok", {
        bubbles: true,
        composed: true,
        detail: {
          smartQuotes: this.smartQuotesBox?.checked !== false,
          emDash: this.dashBox?.checked !== false,
          ellipsis: this.ellipsisBox?.checked !== false,
          ordinalSuperscript: this.ordinalBox?.checked === true,
          capitalizeFirstLetter: this.capitalizeBox?.checked !== false,
          hyperlinkAutoformat: this.hyperlinkBox?.checked !== false,
          replacements: this.#rows
            .map(({ from, to }) => ({ from: from.trim(), to: to.trim() }))
            .filter(({ from, to }) => from && to),
          exceptions: [
            ...new Set(
              (this.exceptionsEl?.value ?? "")
                .split(/[\n,，;；]+/u)
                .map((word) => word.trim())
                .filter(Boolean),
            ),
          ],
        } satisfies AutocorrectDialogValues,
      }),
    );
    this.hide();
  };

  /** Template-bound (public — the FAST template cannot reach a #private
   *  member): the table's delegated click — a remove button drops its row. */
  onTableClick(event: Event): void {
    const button = (event.target as HTMLElement | null)?.closest("[data-role='remove']");
    if (!button) return;
    const index = Number(button.getAttribute("data-index"));
    if (!Number.isInteger(index)) return;
    this.#collect();
    this.#rows.splice(index, 1);
    this.#renderRows();
  }

  addRow(): void {
    this.#collect();
    this.#rows.push({ from: "", to: "" });
    this.#renderRows();
  }

  /** Read the in-progress row text back into `#rows` (the DOM is the live
   *  buffer; inputs render from this snapshot when rows are added/removed). */
  #collect(): void {
    const nodes = this.tableEl?.querySelectorAll(".replace-row");
    if (!nodes) return;
    nodes.forEach((node, index) => {
      const row = this.#rows[index];
      if (!row) return;
      const from = node.querySelector<TextInputLike>("[data-role='from']");
      const to = node.querySelector<TextInputLike>("[data-role='to']");
      if (from) row.from = from.value;
      if (to) row.to = to.value;
    });
  }

  #renderRows(): void {
    const table = this.tableEl;
    if (!table) return;
    const rows: HTMLElement[] = [];
    for (const [index, row] of this.#rows.entries()) {
      const el = document.createElement("div");
      el.className = "replace-row";
      const from = document.createElement("fluent-text-input") as TextInputLike;
      from.setAttribute("data-role", "from");
      from.setAttribute("aria-label", t("autocorrect.find", this));
      from.value = row.from;
      const arrow = document.createElement("span");
      arrow.className = "arrow";
      arrow.textContent = "→";
      const to = document.createElement("fluent-text-input") as TextInputLike;
      to.setAttribute("data-role", "to");
      to.setAttribute("aria-label", t("autocorrect.replaceWith", this));
      to.value = row.to;
      const remove = document.createElement("fluent-button");
      remove.setAttribute("data-role", "remove");
      remove.setAttribute("data-index", String(index));
      remove.setAttribute("appearance", "stealth");
      remove.setAttribute("icon-only", "");
      remove.setAttribute("title", t("autocorrect.remove", this));
      remove.setAttribute("aria-label", t("autocorrect.remove", this));
      remove.textContent = "✕";
      el.append(from, arrow, to, remove);
      rows.push(el);
    }
    table.replaceChildren(...rows);
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("autocorrect.title", this);
    if (this.rulesHeadingEl)
      this.rulesHeadingEl.textContent = t("autocorrect.replaceAsYouType", this);
    if (this.smartQuotesLabelEl)
      this.smartQuotesLabelEl.textContent = t("autocorrect.smartQuotes", this);
    if (this.dashLabelEl) this.dashLabelEl.textContent = t("autocorrect.hyphens", this);
    if (this.ellipsisLabelEl) this.ellipsisLabelEl.textContent = t("autocorrect.ellipsis", this);
    if (this.ordinalLabelEl) this.ordinalLabelEl.textContent = t("autocorrect.ordinals", this);
    if (this.capitalizeLabelEl)
      this.capitalizeLabelEl.textContent = t("autocorrect.capitalize", this);
    if (this.hyperlinkLabelEl)
      this.hyperlinkLabelEl.textContent = t("autocorrect.hyperlinks", this);
    if (this.replacementsHeadingEl)
      this.replacementsHeadingEl.textContent = t("autocorrect.replacements", this);
    if (this.fromHeadEl) this.fromHeadEl.textContent = t("autocorrect.find", this);
    if (this.toHeadEl) this.toHeadEl.textContent = t("autocorrect.replaceWith", this);
    if (this.addBtn) this.addBtn.textContent = t("autocorrect.add", this);
    if (this.exceptionsHeadingEl)
      this.exceptionsHeadingEl.textContent = t("autocorrect.exceptions", this);
    if (this.exceptionsEl) this.exceptionsEl.placeholder = t("autocorrect.exceptionsHint", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
  }
}

export default DocenAutocorrectDialog;
