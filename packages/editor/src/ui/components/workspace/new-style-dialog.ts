import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import type { ModifyStylePatch, NewStyleDefinition } from "../../../document/extensions/commands";
import {
  CJK_FONT_NAMES,
  FONT_NAMES,
  FONT_SIZES_CN,
  FONT_SIZES_PT,
} from "../../../document/font-lists";
import { observeLang, resolveLang, t } from "../../i18n/localize";
import { renderIcon } from "../ribbon/command-helpers";
import { opt, pick, pickLadder, pickedValue, type FluentDropdown } from "./fluent-combo";
import { ALIGN_GROUP, LINE_GROUP, type StyleChoice } from "./modify-style-dialog";

export interface NewStyleState {
  choices: StyleChoice[];
  defaultName?: string;
  type?: "paragraph" | "character";
  basedOn?: string | null;
  next?: string | null;
}

export type { StyleChoice };

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog {
    --dialog-width: min(552px, 94vw);
  }
  .body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .section-label {
    font-weight: 600;
    opacity: 0.8;
  }
  .props {
    display: grid;
    grid-template-columns: max-content 1fr;
    gap: 8px 12px;
    align-items: center;
  }
  .props > label {
    display: flex;
    white-space: nowrap;
  }
  .props > label .fill {
    flex: 1;
    text-align: justify;
    text-align-last: justify;
  }
  .props > fluent-dropdown,
  .props > fluent-text-input {
    width: 100%;
    min-width: 0;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .row > fluent-dropdown {
    min-width: 0;
    flex: none;
  }
  fluent-dropdown.font-sel {
    width: 120px;
  }
  fluent-dropdown.size-sel {
    width: 112px;
  }
  fluent-dropdown.scope-sel {
    flex: none;
    width: 88px;
    margin-inline-start: auto;
  }
  fluent-toggle-button.mini {
    width: 28px;
    height: 28px;
    min-width: 28px;
    padding: 0;
  }
  fluent-toggle-button.mini[data-active="true"] {
    background: var(--neutral-fill-secondary-rest, #e5e5e5);
  }
  .sep {
    width: 1px;
    height: 16px;
    background: var(--neutral-stroke-rest, #d1d1d1);
    margin: 0 2px;
    flex: none;
  }
  .preview {
    border: 1px solid var(--neutral-stroke-rest, #d1d1d1);
    border-radius: 4px;
    padding: 6px 10px 8px;
    min-height: 48px;
  }
  .preview > label {
    font-size: 12px;
    opacity: 0.7;
  }
  .preview-text {
    margin-top: 2px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .checks {
    display: flex;
    gap: 16px;
    flex-wrap: wrap;
  }
  .check-field {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }
  .actions {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
  }
  .actions > .spacer {
    flex: 1;
  }
  .hidden {
    display: none !important;
  }
`;

type FluentCheckbox = HTMLElement & { checked?: boolean };

const template = html<DocenNewStyleDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="body">
      <div class="section-label" ${ref("propsLabel")}></div>
      <div class="props">
        <label ${ref("nameLabel")}></label>
        <fluent-text-input appearance="outline" ${ref("nameInput")}></fluent-text-input>

        <label ${ref("typeLabel")}></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("typeSel")}
          @change="${(x) => x.onTypeChange()}"
        >
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            size="1"
            style="width:100%;box-sizing:border-box"
          />
        </fluent-dropdown>

        <label ${ref("basedOnLabel")}></label>
        <fluent-dropdown type="combobox" appearance="outline" ${ref("basedOnSel")}>
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            size="1"
            style="width:100%;box-sizing:border-box"
          />
        </fluent-dropdown>

        <label
          ${ref("nextLabel")}
          class="${(x) => (x.styleType === "character" ? "hidden" : "")}"
        ></label>
        <fluent-dropdown
          type="combobox"
          appearance="outline"
          ${ref("nextSel")}
          class="${(x) => (x.styleType === "character" ? "hidden" : "")}"
        >
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
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

      <div class="section-label" ${ref("formatLabel")}></div>
      <div class="row">
        <fluent-dropdown type="combobox" appearance="outline" class="font-sel" ${ref("fontSel")}>
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            size="1"
            style="width:100%;box-sizing:border-box"
          />
        </fluent-dropdown>
        <fluent-dropdown type="combobox" appearance="outline" class="size-sel" ${ref("sizeSel")}>
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
          <input
            slot="control"
            role="combobox"
            aria-haspopup="listbox"
            type="combobox"
            size="1"
            style="width:100%;box-sizing:border-box"
          />
        </fluent-dropdown>
        <fluent-toggle-button
          class="mini"
          appearance="subtle"
          ${ref("boldBtn")}
          @click="${(x) => x.pressMini("bold")}"
        ></fluent-toggle-button>
        <fluent-toggle-button
          class="mini"
          appearance="subtle"
          ${ref("italicBtn")}
          @click="${(x) => x.pressMini("italic")}"
        ></fluent-toggle-button>
        <fluent-toggle-button
          class="mini"
          appearance="subtle"
          ${ref("underlineBtn")}
          @click="${(x) => x.pressMini("underline")}"
        ></fluent-toggle-button>
        <docen-color-picker
          icon="font-color"
          event="new-style-color"
          no-split
          ${ref("colorPick")}
        ></docen-color-picker>
        <fluent-dropdown type="combobox" appearance="outline" class="scope-sel" ${ref("scopeSel")}>
          <fluent-listbox popover="manual" tabindex="-1"></fluent-listbox>
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

      <div
        class="row para-row ${(x) => (x.styleType === "character" ? "hidden" : "")}"
        ${ref("paraRow")}
      ></div>

      <div class="preview">
        <label ${ref("previewLabel")}></label>
        <div class="preview-text" ${ref("previewText")}></div>
      </div>

      <div class="checks">
        <label class="check-field">
          <fluent-checkbox ${ref("quickCheck")}></fluent-checkbox>
          <span ${ref("quickCheckLabel")}></span>
        </label>
        <label class="check-field">
          <fluent-checkbox ${ref("autoCheck")}></fluent-checkbox>
          <span ${ref("autoCheckLabel")}></span>
        </label>
      </div>

      <div class="actions">
        <div class="spacer"></div>
        <fluent-button
          appearance="neutral"
          ${ref("cancelBtn")}
          @click="${(x) => x.hide()}"
        ></fluent-button>
        <fluent-button
          appearance="accent"
          ${ref("okBtn")}
          @click="${(x) => x.applyPatch()}"
        ></fluent-button>
      </div>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-new-style-dialog>` — the Word "Create New Style from Formatting"
 * dialog.
 */
@customElement({ name: "docen-new-style-dialog", template, styles })
export default class DocenNewStyleDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable propsLabel?: HTMLElement;
  @observable nameLabel?: HTMLElement;
  @observable nameInput?: HTMLElement & { value?: string };
  @observable typeLabel?: HTMLElement;
  @observable typeSel?: FluentDropdown;
  @observable basedOnLabel?: HTMLElement;
  @observable basedOnSel?: FluentDropdown;
  @observable nextLabel?: HTMLElement;
  @observable nextSel?: FluentDropdown;
  @observable formatLabel?: HTMLElement;
  @observable fontSel?: FluentDropdown;
  @observable sizeSel?: FluentDropdown;
  @observable boldBtn?: HTMLElement;
  @observable italicBtn?: HTMLElement;
  @observable underlineBtn?: HTMLElement;
  @observable colorPick?: HTMLElement;
  @observable scopeSel?: FluentDropdown;
  @observable paraRow?: HTMLElement;
  @observable previewLabel?: HTMLElement;
  @observable previewText?: HTMLElement;
  @observable quickCheck?: FluentCheckbox;
  @observable quickCheckLabel?: HTMLElement;
  @observable autoCheck?: FluentCheckbox;
  @observable autoCheckLabel?: HTMLElement;
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  @observable styleType: "paragraph" | "character" = "paragraph";

  #bold = false;
  #italic = false;
  #underline = false;
  #color: string | null = null;
  #alignment: string | null = null;
  #lineSpacing: number | null = null;
  #fontScope = "all";
  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    for (const btn of [this.boldBtn, this.italicBtn, this.underlineBtn]) {
      if (btn) (btn as unknown as { press?: () => void }).press = () => {};
    }
    this.#bindColorPicker();
    this.scopeSel?.addEventListener("change", () => {
      this.#fontScope = pickedValue(this.scopeSel) || "all";
      this.#fillFontCombo();
      this.#updatePreview();
    });
    this.fontSel?.addEventListener("change", () => this.#updatePreview());
    this.sizeSel?.addEventListener("change", () => this.#updatePreview());
    this.nameInput?.addEventListener("input", () => this.#updatePreview());
    this.#unobserveLang = observeLang(() => {
      this.#applyLabels();
      this.#buildParaRow();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(state: NewStyleState): void {
    this.styleType = state.type ?? "paragraph";
    this.#fontScope = "all";
    this.#bold = false;
    this.#italic = false;
    this.#underline = false;
    this.#color = null;
    this.#alignment = null;
    this.#lineSpacing = null;

    const defaultName = state.defaultName || "Style 1";
    if (this.nameInput) this.nameInput.value = defaultName;

    this.#fillType();
    this.#fillChoices(this.basedOnSel, state.choices, state.basedOn ?? "Normal", true);
    this.#fillChoices(this.nextSel, state.choices, state.next ?? "Normal", true);
    this.#fillScopeCombo();
    this.#fillCombos();

    this.#setPressed(this.boldBtn, false);
    this.#setPressed(this.italicBtn, false);
    this.#setPressed(this.underlineBtn, false);

    if (this.colorPick) this.colorPick.setAttribute("default-color", "000000");
    this.#buildParaRow();

    if (this.quickCheck) this.quickCheck.checked = true;
    if (this.autoCheck) this.autoCheck.checked = false;

    this.#updatePreview();
    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  onTypeChange(): void {
    const val = pickedValue(this.typeSel);
    this.styleType = val === "character" ? "character" : "paragraph";
    this.#updatePreview();
  }

  applyPatch(): void {
    const name = this.nameInput?.value?.trim() || "Style 1";
    const patch: ModifyStylePatch = {
      id: name.replace(/[^a-zA-Z0-9]/g, "") || "Style1",
      name,
      basedOn: pickedValue(this.basedOnSel) || null,
      next: this.styleType === "paragraph" ? pickedValue(this.nextSel) || null : null,
      font: pickedValue(this.fontSel) || null,
      size: pickedValue(this.sizeSel) ? Number(pickedValue(this.sizeSel)) : null,
      bold: this.#bold,
      italic: this.#italic,
      underline: this.#underline,
      color: this.#color,
      alignment: this.styleType === "paragraph" ? this.#alignment : undefined,
      lineSpacing: this.styleType === "paragraph" ? this.#lineSpacing : undefined,
      quickFormat: this.quickCheck?.checked ?? true,
      autoRedefine: this.autoCheck?.checked ?? false,
    };

    const def: NewStyleDefinition = {
      id: patch.id,
      name,
      type: this.styleType,
      basedOn: patch.basedOn ?? undefined,
      next: patch.next ?? undefined,
      quickFormat: patch.quickFormat,
      autoRedefine: patch.autoRedefine,
      patch,
    };

    this.$emit("new-style:ok", def);
    this.hide();
  }

  pressMini(which: "bold" | "italic" | "underline"): void {
    const btn =
      which === "bold" ? this.boldBtn : which === "italic" ? this.italicBtn : this.underlineBtn;
    const on = !((btn as unknown as { pressed?: boolean } | undefined)?.pressed === true);
    this.#setPressed(btn, on);
    if (which === "bold") this.#bold = on;
    else if (which === "italic") this.#italic = on;
    else this.#underline = on;
    this.#updatePreview();
  }

  readonly #onColor = (event: Event): void => {
    const v = (event as CustomEvent).detail?.value as string | { val?: string } | undefined;
    this.#color = v == null || v === "none" ? null : typeof v === "string" ? v : (v.val ?? null);
    this.#updatePreview();
  };

  #bindColorPicker(): void {
    if (!this.colorPick || this.colorPick.dataset.bound) return;
    this.colorPick.dataset.bound = "1";
    this.colorPick.addEventListener("command", this.#onColor);
  }

  #fillType(): void {
    const listbox = this.typeSel?.querySelector("fluent-listbox");
    if (!listbox) return;
    listbox.replaceChildren(
      opt(t("modifyStyleDialog.typeParagraph", this), "paragraph"),
      opt(t("newStyleDialog.typeCharacter", this), "character"),
    );
    pick(this.typeSel, this.styleType);
  }

  #fillFontCombo(): void {
    const listbox = this.fontSel?.querySelector("fluent-listbox");
    if (!listbox) return;
    const names =
      this.#fontScope === "cjk"
        ? CJK_FONT_NAMES
        : this.#fontScope === "latin"
          ? FONT_NAMES
          : [...new Set([...CJK_FONT_NAMES, ...FONT_NAMES])];
    listbox.replaceChildren(...names.map((name) => opt(name, name)));
    pickLadder(this.fontSel, "Calibri");
  }

  #fillScopeCombo(): void {
    const listbox = this.scopeSel?.querySelector("fluent-listbox");
    if (!listbox) return;
    listbox.replaceChildren(
      opt(t("modifyStyleDialog.allFonts", this), "all"),
      opt(t("modifyStyleDialog.cjkFonts", this), "cjk"),
      opt(t("modifyStyleDialog.latinFonts", this), "latin"),
    );
    pick(this.scopeSel, this.#fontScope);
  }

  #fillCombos(): void {
    const listBoxes = [this.fontSel, this.sizeSel].map((s) => s?.querySelector("fluent-listbox"));
    if (listBoxes.some((b) => !b)) return;
    const [, sizeBox] = listBoxes;
    this.#fillFontCombo();
    const zh = resolveLang(this).toLowerCase().startsWith("zh");
    const ladder = [
      ...(zh ? FONT_SIZES_CN.map(([name, pt]) => opt(`${name} (${pt})`, String(pt))) : []),
      ...FONT_SIZES_PT.map((pt) => opt(String(pt), String(pt))),
    ];
    sizeBox!.replaceChildren(...ladder);
    pickLadder(this.sizeSel, "11");
  }

  #fillChoices(
    sel: FluentDropdown | undefined,
    choices: StyleChoice[],
    picked: string,
    blank: boolean,
  ): void {
    const listbox = sel?.querySelector("fluent-listbox");
    if (!listbox) return;
    const options: HTMLElement[] = [];
    if (blank) options.push(opt(t("modifyStyleDialog.inherit", this), ""));
    for (const c of choices) options.push(opt(c.name, c.id));
    if (picked && !options.some((o) => o.getAttribute("value") === picked))
      options.splice(blank ? 1 : 0, 0, opt(picked, picked));
    listbox.replaceChildren(...options);
    pick(sel, picked);
  }

  #setPressed(el: HTMLElement | null | undefined, on: boolean): void {
    if (el) (el as unknown as { pressed?: boolean }).pressed = on;
  }

  #iconToggle(icon: string, title: string, bars?: HTMLElement): HTMLElement {
    const b = document.createElement("fluent-toggle-button");
    b.className = "mini";
    b.setAttribute("appearance", "subtle");
    b.title = title;
    (b as unknown as { press?: () => void }).press = () => {};
    if (bars) b.append(bars);
    else renderIcon(b, icon);
    return b;
  }

  #barsSvg(path: string): HTMLElement {
    const span = document.createElement("span");
    span.innerHTML =
      `<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true">` +
      `<path d="${path}" fill="none" stroke="currentColor" stroke-width="1.3" ` +
      `stroke-linecap="round"/></svg>`;
    return span;
  }

  #buildParaRow(): void {
    const row = this.paraRow;
    if (!row) return;
    row.replaceChildren();
    for (const [value, icon, titleKey] of ALIGN_GROUP) {
      const b = this.#iconToggle(icon, t(titleKey, this));
      b.dataset.value = value;
      this.#setPressed(b, this.#alignment === value);
      b.addEventListener("click", () => {
        const on = !((b as unknown as { pressed?: boolean }).pressed === true);
        this.#setPressed(b, on);
        for (const other of ALIGN_GROUP)
          if (other[0] !== value)
            this.#setPressed(row.querySelector(`[data-value="${other[0]}"]`), false);
        this.#alignment = on ? value : null;
        this.#updatePreview();
      });
      row.append(b);
    }
    row.append(this.#separator());
    for (const [value, path, titleKey] of LINE_GROUP) {
      const b = this.#iconToggle("", t(titleKey, this), this.#barsSvg(path));
      b.dataset.line = String(value);
      this.#setPressed(b, this.#lineSpacing === value);
      b.addEventListener("click", () => {
        const on = !((b as unknown as { pressed?: boolean }).pressed === true);
        this.#setPressed(b, on);
        for (const other of LINE_GROUP)
          if (other[0] !== value)
            this.#setPressed(row.querySelector(`[data-line="${other[0]}"]`), false);
        this.#lineSpacing = on ? value : null;
      });
      row.append(b);
    }
  }

  #separator(): HTMLElement {
    const sep = document.createElement("div");
    sep.className = "sep";
    return sep;
  }

  #propsLabel(el: HTMLElement | undefined, key: string): void {
    if (!el) return;
    el.innerHTML = `<span class="fill">${t(key, this)}</span><span>:</span>`;
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("newStyleDialog.title", this);
    if (this.propsLabel) this.propsLabel.textContent = t("modifyStyleDialog.properties", this);
    this.#propsLabel(this.nameLabel, "modifyStyleDialog.name");
    this.#propsLabel(this.typeLabel, "modifyStyleDialog.styleType");
    this.#propsLabel(this.basedOnLabel, "modifyStyleDialog.basedOn");
    this.#propsLabel(this.nextLabel, "modifyStyleDialog.next");
    if (this.formatLabel) this.formatLabel.textContent = t("modifyStyleDialog.formatting", this);
    if (this.previewLabel) this.previewLabel.textContent = t("modifyStyleDialog.preview", this);
    if (this.quickCheckLabel)
      this.quickCheckLabel.textContent = t("modifyStyleDialog.quickFormat", this);
    if (this.autoCheckLabel)
      this.autoCheckLabel.textContent = t("modifyStyleDialog.autoRedefine", this);
    if (this.okBtn) this.okBtn.textContent = t("common.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("common.cancel", this);
    if (this.boldBtn) this.boldBtn.title = t("ribbon.cmd.bold", this);
    if (this.italicBtn) this.italicBtn.title = t("ribbon.cmd.italic", this);
    if (this.underlineBtn) this.underlineBtn.title = t("ribbon.cmd.underline", this);
    this.#fillType();
    this.#fillScopeCombo();
  }

  #updatePreview(): void {
    if (!this.previewText) return;
    const text = this.nameInput?.value?.trim() || "Style 1";
    this.previewText.textContent = text;
    const font = pickedValue(this.fontSel);
    const size = pickedValue(this.sizeSel);
    const cssParts: string[] = [];
    if (font) cssParts.push(`font-family:${font}`);
    if (size) cssParts.push(`font-size:${size}pt`);
    if (this.#bold) cssParts.push("font-weight:bold");
    if (this.#italic) cssParts.push("font-style:italic");
    if (this.#underline) cssParts.push("text-decoration:underline");
    if (this.#color) cssParts.push(`color:#${this.#color}`);
    if (this.styleType === "paragraph" && this.#alignment) {
      cssParts.push(`text-align:${this.#alignment === "both" ? "justify" : this.#alignment}`);
    }
    this.previewText.style.cssText = cssParts.join(";");
  }
}
