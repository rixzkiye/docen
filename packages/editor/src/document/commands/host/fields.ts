import type { HostCommandDomain } from "./registry";

/** The field domain's view of the host — only what its command bodies touch. */
export interface FieldsHostView {
  /** Open the Field dialog (commit arrives via field:ok). */
  fieldInsert(): void;
  /** Re-resolve the field atom under the caret (Word's F9). */
  fieldUpdateAtSelection(): void;
  updateAllFields(): void;
  /** Edit the field code under the caret. */
  fieldEditAtSelection(): void;
  /** Toggle a checkbox field's state. */
  fieldToggleCheckboxAtSelection(): void;
  /** Word's field-code projection toggle (Alt+F9). */
  toggleFieldCodes(): void;
  /** Insert one placeholder math template at the caret. */
  insertEquation(template: string): void;
  /** Insert one equation-context symbol character. */
  insertEquationSymbol(char: string): void;
}

/**
 * Field and equation commands split out of the host element: the Field dialog
 * entries, update/edit/checkbox atoms, the field-codes toggle, and the
 * equation templates + symbols of the Insert tab and Equation Tools tab.
 */
export class FieldsHostCommands implements HostCommandDomain {
  constructor(private readonly host: FieldsHostView) {}

  readonly chrome: readonly string[] = [];

  readonly editor: readonly string[] = [
    "insert-field",
    "update-field",
    "update-all-fields",
    "toggle-field-codes",
    "edit-field",
    "toggle-field-checkbox",
    "equation",
    "insert-symbol",
  ];

  run(event: string, value?: string): boolean {
    // Field — open the Field dialog (Insert → Text → Explore Quick Parts →
    // Field); the commit arrives via field:ok. The context-menu entries act
    // on the field atom under the caret (update = Word's F9).
    if (event === "insert-field") {
      this.host.fieldInsert();
      return true;
    }
    if (event === "update-field") {
      this.host.fieldUpdateAtSelection();
      return true;
    }
    if (event === "update-all-fields") {
      this.host.updateAllFields();
      return true;
    }
    if (event === "toggle-field-codes") {
      this.host.toggleFieldCodes();
      return true;
    }
    if (event === "edit-field") {
      this.host.fieldEditAtSelection();
      return true;
    }
    if (event === "toggle-field-checkbox") {
      this.host.fieldToggleCheckboxAtSelection();
      return true;
    }
    // Equation — insert one placeholder math template at the caret (Word's
    // Insert → Symbols → Equation gallery).
    if (event === "equation") {
      this.host.insertEquation(value ? String(value) : "fraction");
      return true;
    }
    if (event === "insert-symbol") {
      this.host.insertEquationSymbol(value ? String(value) : "");
      return true;
    }
    return false;
  }
}
