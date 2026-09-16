import type { HostCommandDomain } from "./registry";

/** The sections/page-setup domain's view of the host — only what its command
 *  bodies touch. */
export interface SectionsHostView {
  /** Open the Page Setup dialog for the current section. */
  openPageSetup(): void;
  /** Page size presets (w:pgSz via the section command). */
  setPageSize(value?: string): void;
  setOrientation(value?: string): void;
  setMargins(value?: string): void;
  /** Open the Columns dialog prefilled from the current section. */
  openColumnsDialog(): void;
  setColumnCount(count: number): void;
  /** Open the Line Numbering Options dialog. */
  openLineNumbersOptions(): void;
  /** w:lnNumType's restart mode on the current section. */
  setLineNumbers(mode: "none" | "continuous" | "newPage" | "newSection"): void;
  /** Open Word's Borders and Shading dialog on the given tab. */
  openBordersDialog(tab: "border" | "page"): void;
  setPageBorders(preset?: string): void;
  /** Insert the cover block at the document start. */
  insertCoverPage(): void;
  /** Insert two page breaks (Word's Blank Page). */
  insertBlankPage(): void;
  /** Set hyphenation mode (none/auto/manual). */
  setHyphenation(mode: "none" | "auto" | "manual"): void;
  /** Open Hyphenation Options dialog. */
  openHyphenationOptions(): void;
  /** Insert discretionary soft hyphen. */
  insertSoftHyphen(): void;
}

/**
 * Page-setup and Insert → Pages commands split out of the host element: the
 * Page Setup dialog entries (size/orientation/margins/columns/line numbers),
 * the Borders and Shading dialog entries, and cover/blank pages.
 */
export class SectionsHostCommands implements HostCommandDomain {
  constructor(private readonly host: SectionsHostView) {}

  readonly chrome: readonly string[] = [
    "page-setup-dialog",
    "page-size",
    "orientation",
    "margins",
    "columns",
    "line-numbers",
    "hyphenation",
    "insert-soft-hyphen",
  ];

  readonly editor: readonly string[] = ["border", "page-border", "cover-page", "blank-page"];

  run(event: string, value?: string): boolean {
    // Page setup actions write sectionProperties; the transaction re-renders.
    // "more"/"custom" open the Page Setup dialog instead of a preset; the
    // Page Setup group's dialog-box launcher opens the same dialog.
    if (event === "page-setup-dialog") {
      this.host.openPageSetup();
      return true;
    }
    if (event === "page-size") {
      if (value === "more") this.host.openPageSetup();
      else this.host.setPageSize(value);
      return true;
    }
    if (event === "orientation") {
      this.host.setOrientation(value);
      return true;
    }
    if (event === "margins") {
      if (value === "custom") this.host.openPageSetup();
      else this.host.setMargins(value);
      return true;
    }
    // Columns presets (the Layout tab's Columns menu: one/two/three);
    // More Columns opens the dialog prefilled from the current section.
    if (event === "columns") {
      if (value === "more") this.host.openColumnsDialog();
      else {
        const count = Number(value);
        if (count >= 1 && count <= 9) this.host.setColumnCount(count);
      }
      return true;
    }
    // Line Numbers menu (the Layout tab): the mode writes w:lnNumType's
    // restart on the current section; "none" clears the numbering; the
    // options entry opens the Line Numbering Options dialog.
    if (event === "line-numbers") {
      if (value === "options") this.host.openLineNumbersOptions();
      else if (
        value === "none" ||
        value === "continuous" ||
        value === "newPage" ||
        value === "newSection"
      )
        this.host.setLineNumbers(value);
      return true;
    }
    if (event === "hyphenation") {
      if (value === "options") this.host.openHyphenationOptions();
      else if (value === "none" || value === "auto" || value === "manual")
        this.host.setHyphenation(value);
      return true;
    }
    if (event === "insert-soft-hyphen") {
      this.host.insertSoftHyphen();
      return true;
    }
    // The Borders and Shading dialog entries — the border split's and the
    // page-border split's last item carry the dialog value; the source split
    // picks the tab (the remaining preset values fall through below).
    if (value === "borders-shading" && (event === "border" || event === "page-border")) {
      this.host.openBordersDialog(event === "page-border" ? "page" : "border");
      return true;
    }
    if (event === "page-border") {
      this.host.setPageBorders(value);
      return true;
    }
    // Insert → Pages menu: a cover block at the document start, or two page
    // breaks (Word's Blank Page).
    if (event === "cover-page") {
      this.host.insertCoverPage();
      return true;
    }
    if (event === "blank-page") {
      this.host.insertBlankPage();
      return true;
    }
    return false;
  }
}
