import type { RibbonTab } from "../../ui";
import { btn, col, grid, group, tabNode } from "./shared";

/**
 * Developer tab in the Fluent Ribbon.
 * Hosts Code (Macros / Visual Basic placeholder), Controls (Word's Content Controls:
 * Rich Text, Plain Text, Picture, Checkbox, Combo Box, Drop-Down List, Date Picker,
 * Building Blocks, Design Mode, Properties), and Protect (Restrict Editing).
 */
export const developerTab = (): RibbonTab =>
  tabNode("developer", [
    group("code", [
      btn("group-objects", "view-macros", { size: "large" }),
      btn("edit", "record-macro", { size: "large" }),
    ]),
    group("controls", [
      grid([
        btn("edit", "sdt-rich-text"),
        btn("text", "sdt-plain-text"),
        btn("image", "sdt-picture"),
        btn("checkbox", "sdt-checkbox"),
        btn("menu", "sdt-combo-box"),
        btn("menu", "sdt-dropdown"),
        btn("calendar", "sdt-date"),
        btn("copy", "sdt-building-block"),
      ]),
      col([btn("grid", "toggle-design-mode"), btn("settings", "sdt-properties")]),
    ]),
    group("protect", [btn("protect", "restrict-editing", { size: "large" })]),
  ]);
