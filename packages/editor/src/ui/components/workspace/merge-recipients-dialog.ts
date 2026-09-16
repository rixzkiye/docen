import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { MergeRecipients } from "../../../document/commands/mail-merge";
import { observeLang, t } from "../../i18n/localize";

export interface RecipientFilterState {
  searchQuery: string;
  sortColumn: number;
  sortAscending: boolean;
  selectedIndices: Set<number>;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(640px, 94vw);
  }
  .recipients-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .toolbar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 8px;
  }
  .toolbar input[type="text"] {
    padding: 4px 8px;
    font-size: 12px;
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 4px;
    flex: 1;
    max-width: 240px;
  }
  .button-group {
    display: flex;
    gap: 6px;
  }
  .button-group button {
    padding: 4px 8px;
    font-size: 12px;
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 3px;
    background: transparent;
    cursor: pointer;
  }
  .button-group button:hover {
    background: var(--colorNeutralBackground1Hover, #f5f5f5);
  }
  .table-container {
    max-height: 280px;
    overflow: auto;
    border: 1px solid var(--docen-color-divider, #e0e0e0);
    border-radius: 4px;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
  }
  th,
  td {
    padding: 6px 8px;
    border-bottom: 1px solid var(--docen-color-divider, #f0f0f0);
    text-align: left;
    white-space: nowrap;
  }
  th {
    background: var(--docen-color-neutral-background, #f7f7f7);
    font-weight: 600;
    cursor: pointer;
    user-select: none;
    position: sticky;
    top: 0;
    z-index: 1;
  }
  th:hover {
    background: #eaeaea;
  }
  th.checkbox-col,
  td.checkbox-col {
    width: 30px;
    text-align: center;
    cursor: default;
  }
  tr.excluded {
    opacity: 0.5;
    background: #fafafa;
  }
  .count-bar {
    font-size: 11px;
    color: var(--docen-color-foreground-secondary, #616161);
  }
  .footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 8px;
  }
`;

const template = html<DocenMergeRecipientsDialog>`
  <docen-dialog ${ref("dialog")} modal heading="${(x) => t("recipients.editTitle", x)}">
    <div class="recipients-body" slot="body">
      <div class="toolbar">
        <input
          type="text"
          placeholder="${(x) => t("recipients.filterPlaceholder", x)}"
          ${ref("filterInput")}
          @input="${(x) => x.onFilterChange()}"
        />
        <div class="button-group">
          <button type="button" @click="${(x) => x.selectAll()}">
            ${(x) => t("recipients.selectAll", x)}
          </button>
          <button type="button" @click="${(x) => x.clearAll()}">
            ${(x) => t("recipients.clearAll", x)}
          </button>
        </div>
      </div>

      <div class="table-container">
        <table ${ref("tableEl")}>
          <thead>
            <tr ${ref("headerRowEl")}></tr>
          </thead>
          <tbody ${ref("tbodyEl")}></tbody>
        </table>
      </div>

      <div class="count-bar" ${ref("countBarEl")}></div>
    </div>
    <div class="footer" slot="footer">
      <fluent-button appearance="primary" @click="${(x) => x.onOk()}"
        >${(x) => t("dialog.ok", x)}</fluent-button
      >
      <fluent-button @click="${(x) => x.close()}">${(x) => t("dialog.cancel", x)}</fluent-button>
    </div>
  </docen-dialog>
`;

@customElement({ name: "docen-merge-recipients-dialog", template, styles })
export class DocenMergeRecipientsDialog extends FASTElement {
  @observable dialog?: HTMLElement & { open: boolean };
  @observable filterInput?: HTMLInputElement;
  @observable tableEl?: HTMLTableElement;
  @observable headerRowEl?: HTMLTableRowElement;
  @observable tbodyEl?: HTMLTableSectionElement;
  @observable countBarEl?: HTMLElement;

  #recipients: MergeRecipients | null = null;
  #selectedIndices = new Set<number>();
  #sortCol = 0;
  #sortAsc = true;
  #query = "";
  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => {});
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  show(recipients: MergeRecipients | null): void {
    this.#recipients = recipients;
    this.#selectedIndices = new Set(
      recipients ? Array.from({ length: recipients.rows.length }, (_, i) => i) : [],
    );
    this.#query = "";
    if (this.filterInput) this.filterInput.value = "";
    this.#render();
    if (this.dialog) this.dialog.open = true;
  }

  close(): void {
    if (this.dialog) this.dialog.open = false;
  }

  selectAll(): void {
    if (!this.#recipients) return;
    this.#selectedIndices = new Set(this.#recipients.rows.map((_, i) => i));
    this.#render();
  }

  clearAll(): void {
    this.#selectedIndices.clear();
    this.#render();
  }

  onFilterChange(): void {
    this.#query = this.filterInput?.value.toLowerCase().trim() ?? "";
    this.#renderTableBody();
  }

  sortBy(colIdx: number): void {
    if (this.#sortCol === colIdx) {
      this.#sortAsc = !this.#sortAsc;
    } else {
      this.#sortCol = colIdx;
      this.#sortAsc = true;
    }
    this.#render();
  }

  #render(): void {
    this.#renderHeader();
    this.#renderTableBody();
  }

  #renderHeader(): void {
    if (!this.headerRowEl || !this.#recipients) return;
    this.headerRowEl.replaceChildren();

    const cbTh = document.createElement("th");
    cbTh.className = "checkbox-col";
    const masterCb = document.createElement("input");
    masterCb.type = "checkbox";
    masterCb.checked =
      this.#selectedIndices.size === this.#recipients.rows.length &&
      this.#recipients.rows.length > 0;
    masterCb.addEventListener("change", (e) => {
      const target = e.target as HTMLInputElement;
      if (target.checked) this.selectAll();
      else this.clearAll();
    });
    cbTh.appendChild(masterCb);
    this.headerRowEl.appendChild(cbTh);

    this.#recipients.headers.forEach((hdr, idx) => {
      const th = document.createElement("th");
      th.textContent = hdr + (this.#sortCol === idx ? (this.#sortAsc ? " ▲" : " ▼") : "");
      th.addEventListener("click", () => this.sortBy(idx));
      this.headerRowEl?.appendChild(th);
    });
  }

  #renderTableBody(): void {
    if (!this.tbodyEl || !this.#recipients) return;
    this.tbodyEl.replaceChildren();

    // Collect indexed rows
    let indexed = this.#recipients.rows.map((r, i) => ({ row: r, index: i }));

    // Filter
    if (this.#query) {
      indexed = indexed.filter((item) =>
        item.row.some((cell) => cell.toLowerCase().includes(this.#query)),
      );
    }

    // Sort
    indexed.sort((a, b) => {
      const valA = a.row[this.#sortCol] ?? "";
      const valB = b.row[this.#sortCol] ?? "";
      const cmp = valA.localeCompare(valB, undefined, { numeric: true });
      return this.#sortAsc ? cmp : -cmp;
    });

    for (const item of indexed) {
      const tr = document.createElement("tr");
      const isChecked = this.#selectedIndices.has(item.index);
      if (!isChecked) tr.classList.add("excluded");

      const tdCb = document.createElement("td");
      tdCb.className = "checkbox-col";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = isChecked;
      cb.addEventListener("change", (e) => {
        const checked = (e.target as HTMLInputElement).checked;
        if (checked) this.#selectedIndices.add(item.index);
        else this.#selectedIndices.delete(item.index);
        tr.classList.toggle("excluded", !checked);
        this.#updateCountBar();
      });
      tdCb.appendChild(cb);
      tr.appendChild(tdCb);

      for (let c = 0; c < this.#recipients.headers.length; c++) {
        const td = document.createElement("td");
        td.textContent = item.row[c] ?? "";
        tr.appendChild(td);
      }

      this.tbodyEl.appendChild(tr);
    }

    this.#updateCountBar();
  }

  #updateCountBar(): void {
    if (!this.countBarEl || !this.#recipients) return;
    const total = this.#recipients.rows.length;
    const sel = this.#selectedIndices.size;
    const pattern = t("recipients.selectedCount", this);
    this.countBarEl.textContent = pattern.replace("{0}", String(sel)).replace("{1}", String(total));
  }

  onOk(): void {
    if (!this.#recipients) {
      this.close();
      return;
    }
    const filteredRows = this.#recipients.rows.filter((_, idx) => this.#selectedIndices.has(idx));
    const updated: MergeRecipients = {
      headers: [...this.#recipients.headers],
      rows: filteredRows,
    };
    this.$emit("recipients:updated", { recipients: updated });
    this.close();
  }
}
