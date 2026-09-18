// The element's static chrome: the shadow-root stylesheet and template, plus
// the HTML escaping the filename header interpolation needs.

import { css, html } from "@microsoft/fast-element";

/** Escape a host-supplied string for safe interpolation into innerHTML. The
 *  `filename` attribute comes from a user-selected File.name at openDOCX, which
 *  can contain markup — without escaping it flows into #renderHeader's template
 *  and executes. */
export const escapeHtml = (s: string): string =>
  s.replace(/[&<>"']/g, (c) =>
    c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : c === '"' ? "&quot;" : "&#39;",
  );

export const documentStyles = css`
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
    /* Anchors the input layer (see the template comment there). */
    position: relative;
  }
  docen-workspace {
    flex: 1 1 auto;
    min-height: 0;
    height: 100%;
  }
  .input-layer {
    position: absolute;
    inset: 0;
    /* The layer itself must never intercept pointer input — only the bridge's
       programmatic textarea focus uses it. */
    pointer-events: none;
  }
  /* Office ribbon group layout helpers — a large button beside stacked rows of
       small icon-only buttons. Applied to light-DOM wrappers in the ribbon. */
  .rb-col {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .rb-row {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 2px;
    flex-wrap: wrap;
  }
  /* Small icon-only buttons as a 3-row column-flow grid: buttons stack into
       columns of ≤3 (Word's compact group layout), not a flat single row. */
  .rb-grid {
    display: grid;
    grid-template-rows: repeat(3, auto);
    grid-auto-flow: column;
    gap: 2px;
    align-content: start;
  }
  /* data-columns — an N-per-row grid whose rows share column tracks, so the
       second column of a 2×2 group starts at one x regardless of how each
       row's own controls measure. Controls keep their natural width. */
  .rb-grid[data-columns] {
    grid-template-rows: none;
    grid-template-columns: repeat(var(--rb-grid-cols, 2), minmax(0, max-content));
    grid-auto-flow: row;
  }
  .rb-grid[data-columns] > * {
    justify-self: start;
  }
  .rb-vsep {
    width: 1px;
    align-self: stretch;
    background: var(--docen-color-divider, #e1e1e1);
    margin: 0 2px;
  }
  .avatar {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 20px;
    height: 20px;
    border-radius: 50%;
    background: var(--docen-color-brand, #0078d4);
    color: #fff;
    font-size: 10px;
    font-weight: 600;
    margin-inline-end: 4px;
  }
  .avatar-img {
    object-fit: cover;
    background: none;
  }
  /* QAT undo/redo history: an empty stack shows no flyout caret (Word's QAT
       drops it for a fresh document) — #updateStatus stamps the flag. */
  docen-ribbon-split-button[data-history-empty]::part(caret) {
    display: none;
  }
  /* QAT customize — the trigger is Fluent's native icon-only menu-button
       (its own chevron, no injected glyph: injected content duplicates the
       built-in end-slot caret). Sized to the row like the Ribbon Display
       Options button; Fluent's 32px clamps would inflate it into a pill. */
  .qat-customize {
    min-width: 0;
    width: 20px;
    max-width: none;
    min-height: 26px;
    padding-inline: 0;
  }
  /* The canvas surface — the scroll container sits one level up (the
       document-area); this wrapper just anchors the edit bridge's textarea and
       caret overlays (position:relative). cursor:text is the editing surface's
       I-beam, like Word's page area. */
  .docen-canvas {
    position: relative;
    width: fit-content;
    margin: 0 auto;
    padding: 32px 0;
    cursor: text;
  }
  /* Open-progress veil over the canvas (Word centers its opening spinner in
       the document area too): label + Fluent progress bar, centered on a
       translucent white wash so the not-yet-laid-out document doesn't flash
       behind it. The area is the scroll container, so the veil is its FIRST
       child, sticky at top, and one area-height tall (= the visible region,
       so the center lands mid-viewport). While shown, the veil adds its
       height to the scroll range — so freeze the scroller for the duration
       (nothing behind it is worth scrolling to). */
  docen-document-area:has(> .load-veil:not([hidden])) {
    overflow: hidden;
  }
  .load-veil {
    position: sticky;
    top: 0;
    /* 100% = the area's content box (its 24px paddings stay uncovered — under
       the translucent wash the not-yet-replaced document shows as a hairline
       edge, invisible against a blank first load). */
    height: 100%;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 12px;
    background: rgba(255, 255, 255, 0.82);
    z-index: 40;
  }
  .load-veil[hidden] {
    display: none;
  }
  .load-veil fluent-progress-bar {
    width: 240px;
  }
  .load-veil .load-label {
    font-size: 13px;
    color: var(--docen-color-text-2, #424242);
  }
  /* Unsupported-content warning: a Word-style yellow message bar, sticky at
       the top of the scrolling document area so it stays visible while the
       user reads. Shown only when the loaded document carries content the
       editor can preserve but not edit (altChunk, subDoc, SmartArt, OLE,
       raw/custom XML, content parts). */
  .content-warning {
    position: sticky;
    top: 0;
    z-index: 30;
    display: flex;
    align-items: center;
    gap: 8px;
    box-sizing: border-box;
    padding: 6px 12px;
    font-size: 12.5px;
    line-height: 1.45;
    background: #fff4ce;
    color: #3b3b3b;
    border-bottom: 1px solid #f2dc9b;
  }
  .content-warning[hidden] {
    display: none;
  }
  .content-warning-icon {
    flex: none;
  }
  .content-warning-text {
    flex: 1;
    min-width: 0;
  }
  .content-warning-close {
    flex: none;
    border: 0;
    background: transparent;
    color: inherit;
    font: inherit;
    cursor: pointer;
    padding: 2px 6px;
    border-radius: 4px;
  }
  .content-warning-close:hover {
    background: rgba(0, 0, 0, 0.08);
  }
  /* Find Results — Office-style match list: each hit rendered with surrounding
       context and a data-from/to for click-to-jump. Padding keeps items off the
       pane edge (the previous "N matches" text butted right against it). */
  .search-results {
    padding: 6px 8px;
    box-sizing: border-box;
  }
  .search-results .result-count {
    font-size: 12px;
    color: var(--docen-color-marks, #6e6e6e);
    padding: 2px 4px 8px;
  }
  .search-results .result-item {
    display: block;
    width: 100%;
    text-align: start;
    border: none;
    background: transparent;
    padding: 5px 8px;
    margin-block-end: 2px;
    border-radius: 4px;
    font-family: inherit;
    font-size: 12px;
    line-height: 1.45;
    color: #3b3b3b;
    cursor: pointer;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .search-results .result-item:hover {
    background: var(--docen-color-hover, rgba(0, 0, 0, 0.06));
  }
  .search-results .result-item mark {
    background: rgba(255, 235, 59, 0.85);
    color: inherit;
    font-weight: 600;
  }
`;

export const documentTemplate = html`
  <docen-workspace>
    <docen-title-bar slot="header" part="header"></docen-title-bar>
    <docen-ribbon slot="ribbon" part="ribbon"></docen-ribbon>
    <docen-task-pane slot="task-pane-start" position="start" part="nav-pane">
      <docen-navigation-pane>
        <docen-outline slot="headings"></docen-outline>
        <docen-nav-pages slot="pages"></docen-nav-pages>
        <div class="search-results" slot="results" part="search-results"></div>
      </docen-navigation-pane>
    </docen-task-pane>
    <docen-document-area>
      <div class="load-veil" part="load-veil" hidden>
        <fluent-progress-bar></fluent-progress-bar>
        <span class="load-label"></span>
      </div>
      <div class="content-warning" part="content-warning" role="status" hidden>
        <span class="content-warning-icon" aria-hidden="true">⚠️</span>
        <span class="content-warning-text"></span>
        <button class="content-warning-close" type="button">✕</button>
      </div>
      <docen-context-menu part="context-menu">
        <div class="docen-canvas" part="page"></div>
      </docen-context-menu>
    </docen-document-area>
    <docen-task-pane slot="task-pane-end" position="end" part="props-pane">
      <slot name="properties">
        <docen-format-pane></docen-format-pane>
      </slot>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="comments-pane" title="Comments">
      <docen-comments-pane></docen-comments-pane>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="revisions-pane" title="Revisions">
      <docen-revisions-pane></docen-revisions-pane>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="clipboard-pane" title="Clipboard">
      <docen-clipboard-pane></docen-clipboard-pane>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="proofing-pane" title="Spelling">
      <docen-spelling-pane></docen-spelling-pane>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="thesaurus-pane" title="Thesaurus">
      <docen-thesaurus-pane></docen-thesaurus-pane>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="translate-pane" title="Translate">
      <docen-translate-pane></docen-translate-pane>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="styles-pane" title="Styles">
      <docen-styles-pane></docen-styles-pane>
    </docen-task-pane>
    <docen-task-pane
      slot="task-pane-end"
      position="end"
      part="reveal-pane"
      title="Reveal Formatting"
    >
      <docen-reveal-formatting-pane></docen-reveal-formatting-pane>
    </docen-task-pane>
    <docen-task-pane
      slot="task-pane-end"
      position="end"
      part="restrict-pane"
      title="Restrict Editing"
    >
      <docen-restrict-editing-pane></docen-restrict-editing-pane>
    </docen-task-pane>
    <docen-task-pane
      slot="task-pane-end"
      position="end"
      part="a11y-pane"
      title="Accessibility Checker"
    >
      <docen-a11y-checker-pane></docen-a11y-checker-pane>
    </docen-task-pane>
    <docen-task-pane slot="task-pane-end" position="end" part="alt-text-pane" title="Alt Text">
      <docen-alt-text-pane></docen-alt-text-pane>
    </docen-task-pane>
    <docen-status-bar slot="status" part="status"></docen-status-bar>
  </docen-workspace>
  <!-- The edit bridge's textarea lives here, at the shadow root: inside the
       workspace it would sit under docen-context-menu, whose fluent-menu
       treats Space/Enter as menu keys and preventDefaults them — killing the
       textarea's beforeinput (spaces and Enter silently dropped). -->
  <div class="input-layer" part="input-layer"></div>
  <docen-mini-toolbar part="mini-toolbar"></docen-mini-toolbar>
  <docen-key-tips part="key-tips"></docen-key-tips>
  <docen-options-dialog part="options"></docen-options-dialog>
  <docen-autocorrect-dialog part="autocorrect"></docen-autocorrect-dialog>
  <docen-quick-part-dialog part="quick-part"></docen-quick-part-dialog>
  <docen-building-blocks-dialog part="building-blocks"></docen-building-blocks-dialog>
  <docen-word-count-dialog part="word-count"></docen-word-count-dialog>
  <docen-symbol-dialog part="symbol"></docen-symbol-dialog>
  <docen-paragraph-dialog part="paragraph"></docen-paragraph-dialog>
  <docen-page-setup-dialog part="page-setup"></docen-page-setup-dialog>
  <docen-table-dialog part="table"></docen-table-dialog>
  <docen-columns-dialog part="columns"></docen-columns-dialog>
  <docen-link-dialog part="link"></docen-link-dialog>
  <docen-zoom-dialog part="zoom"></docen-zoom-dialog>
  <docen-paste-special-dialog part="paste-special"></docen-paste-special-dialog>
  <docen-font-dialog part="font"></docen-font-dialog>
  <docen-language-dialog part="language"></docen-language-dialog>
  <docen-date-time-dialog part="date-time"></docen-date-time-dialog>
  <docen-toc-dialog part="toc"></docen-toc-dialog>
  <docen-phonetic-dialog part="phonetic"></docen-phonetic-dialog>
  <docen-two-in-one-dialog part="two-in-one"></docen-two-in-one-dialog>
  <docen-define-list-dialog part="define-list"></docen-define-list-dialog>
  <docen-caption-dialog part="caption"></docen-caption-dialog>
  <docen-note-dialog part="note"></docen-note-dialog>
  <docen-note-settings-dialog part="note-settings"></docen-note-settings-dialog>
  <docen-line-numbers-dialog part="line-numbers"></docen-line-numbers-dialog>
  <docen-page-number-format-dialog part="page-number-format"></docen-page-number-format-dialog>
  <docen-field-dialog part="field"></docen-field-dialog>
  <docen-chart-data-dialog part="chart-data"></docen-chart-data-dialog>
  <docen-chart-type-dialog part="chart-type"></docen-chart-type-dialog>
  <docen-compress-pictures-dialog part="compress-pictures"></docen-compress-pictures-dialog>
  <docen-cross-reference-dialog part="cross-reference"></docen-cross-reference-dialog>
  <docen-sources-dialog part="sources"></docen-sources-dialog>
  <docen-recipients-dialog part="recipients"></docen-recipients-dialog>
  <docen-merge-field-dialog part="merge-field"></docen-merge-field-dialog>
  <docen-table-properties-dialog part="table-properties"></docen-table-properties-dialog>
  <docen-drawing-properties-dialog part="drawing-properties"></docen-drawing-properties-dialog>
  <docen-distribute-dialog part="distribute-dialog"></docen-distribute-dialog>
  <docen-borders-shading-dialog part="borders-shading"></docen-borders-shading-dialog>
  <docen-watermark-dialog part="watermark-dialog"></docen-watermark-dialog>
  <docen-fill-effects-dialog part="fill-effects"></docen-fill-effects-dialog>
  <docen-online-pictures-dialog part="online-pictures"></docen-online-pictures-dialog>
  <docen-text-effects-dialog part="text-effects"></docen-text-effects-dialog>
  <docen-inspect-dialog part="inspect"></docen-inspect-dialog>
  <docen-template-dialog part="template"></docen-template-dialog>
  <docen-print-preview part="print-preview"></docen-print-preview>
  <docen-modify-style-dialog part="modify-style"></docen-modify-style-dialog>
  <docen-new-style-dialog part="new-style"></docen-new-style-dialog>
  <docen-hyphenation-dialog part="hyphenation"></docen-hyphenation-dialog>
  <docen-tabs-dialog part="tabs"></docen-tabs-dialog>
  <docen-find-replace-dialog></docen-find-replace-dialog>
  <docen-bookmark-dialog part="bookmark"></docen-bookmark-dialog>
  <docen-go-to-dialog part="go-to"></docen-go-to-dialog>
  <docen-properties-dialog part="properties"></docen-properties-dialog>
  <docen-sdt-dialog part="sdt"></docen-sdt-dialog>
  <docen-compare-dialog part="compare"></docen-compare-dialog>
  <docen-signature-line-dialog part="signature-line"></docen-signature-line-dialog>
  <docen-dropcap-dialog part="dropcap"></docen-dropcap-dialog>
  <docen-merge-recipients-dialog part="merge-recipients"></docen-merge-recipients-dialog>
  <docen-version-history-dialog part="version-history"></docen-version-history-dialog>
  <input type="file" id="file-input" accept=".docx,.docm,.dotx,.dotm,.md,.markdown,.xml" hidden />
  <input type="file" id="image-input" accept="image/*" hidden />
  <input type="file" id="picture-input" accept="image/*" hidden />
  <input type="file" id="text-input" accept=".txt,.md,.markdown,text/plain" hidden />
`;
