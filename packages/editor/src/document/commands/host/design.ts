import type { HostCommandDomain } from "./registry";

/** The Design-tab domain's view of the host — only what its command bodies
 *  touch. */
export interface DesignHostView {
  /** Write/clear the doc-level page background (w:background). */
  setPageColor(
    value?: string | { themeColor: string; val: string; themeTint?: string; themeShade?: string },
  ): void;
  /** Stamp the styles' docDefaults paragraph spacing preset. */
  setParagraphSpacing(preset?: string): void;
  /** Open the custom-watermark dialog. */
  openWatermarkDialog(): void;
  /** Stamp/strip the header watermark shape. */
  setWatermark(preset?: string): void;
  /** Open the page background's Fill Effects dialog. */
  openFillEffectsDialog(): void;
  /** Restore the styles model the document opened with. */
  restoreStylesSnapshot(): void;
}

/**
 * Design-tab commands split out of the host element: Page Color, Paragraph
 * Spacing, Watermark, Fill Effects, and the style-set gallery's "document
 * default" restore.
 */
export class DesignHostCommands implements HostCommandDomain {
  constructor(private readonly host: DesignHostView) {}

  readonly chrome: readonly string[] = [];

  readonly editor: readonly string[] = [
    "page-color",
    "paragraph-spacing",
    "watermark",
    "fill-effects",
    "style-set",
  ];

  run(event: string, value?: string): boolean {
    // Page Color — write/clear the doc-level w:background from the palette
    // (Word's Design → Page Color).
    if (event === "page-color") {
      this.host.setPageColor(
        value as
          | string
          | { themeColor: string; val: string; themeTint?: string; themeShade?: string },
      );
      return true;
    }
    // Paragraph Spacing presets — stamp the styles' docDefaults paragraph
    // spacing (Word's Design → Paragraph Spacing; the document-level default
    // every paragraph without explicit spacing inherits).
    if (event === "paragraph-spacing") {
      this.host.setParagraphSpacing(typeof value === "string" ? value : undefined);
      return true;
    }
    // Watermark gallery — a preset id stamps the header shape, "remove"
    // strips it; the custom entry opens Word's watermark dialog (Word's
    // Design → Watermark split button).
    if (event === "watermark") {
      if (value === "custom") {
        this.host.openWatermarkDialog();
        return true;
      }
      this.host.setWatermark(typeof value === "string" ? value : undefined);
      return true;
    }
    // Fill Effects — the page background's picture fill (Word's Design →
    // Page Color → Fill Effects; the dialog commits via fill-effects:ok).
    if (event === "fill-effects") {
      this.host.openFillEffectsDialog();
      return true;
    }
    // The style-set gallery's "document default" entry restores the styles
    // model the document opened with (the preset values are written by the
    // style-set command; only this entry needs the snapshot).
    if (event === "style-set" && value === "default") {
      this.host.restoreStylesSnapshot();
      return true;
    }
    return false;
  }
}
