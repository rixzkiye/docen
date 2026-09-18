# Technical Audit: Table Cell Write-Protection in Microsoft Word & OOXML

**Batch:** W1.8  
**Scope:** Cell write-protection verification, stock Word OOXML construct analysis, and implementation resolution.  
**Date:** 2026-09-18  
**Status:** **CLOSED — Documented N/A for stock `<w:tcPr>` + Verified Active Forms Protection Support**

---

## 1. Executive Summary

An exhaustive audit of the OOXML specification (ECMA-376 5th edition / ISO/IEC 29500-1) and Microsoft Word's runtime behavior confirms that:

1. **Stock Microsoft Word does NOT possess a native per-cell write-protection construct in `<w:tc>` or `<w:tcPr>`.** Unlike Microsoft Excel (SpreadsheetML), which provides per-cell locking via cell styles (`<c s="...">` referencing `<xf applyProtection="1"><protection locked="1"/></xf>`), WordprocessingML table cell properties (`<w:tcPr>`) have no locking or write-protection attributes.
2. In Microsoft Word, "protected cells" are implemented exclusively through **Document Forms Protection Mode (`w:documentProtection w:edit="forms"`)** combined with **Content Controls (`w:sdt`) / Form Fields (`w:fldSimple`)**, or via explicit **Content Control Locking (`w:lock w:val="contentLocked"`)** inside cells.
3. Docen already implements Microsoft Word's forms protection model via `isInsideEditableSdt` and the `mountEditBridge` gate (`canEdit: (ed) => this.#protectionMode === "forms" ? isInsideEditableSdt(ed) : true`).
4. Inventing a proprietary or non-standard per-cell lock attribute in `<w:tcPr>` would violate OOXML compliance, cause schema validation failures in Word/WPS, and produce document corruption. Therefore, the stock per-cell construct is closed as **N/A (Documented Standard Non-Existence)**, and fidelity is achieved via Word's authentic Forms Protection Mode and Content Control locking.

---

## 2. OOXML Standard Audit: Table Cell Properties (`<w:tcPr>`)

Under ISO/IEC 29500-1 (Section 17.4.70 `tcPr`), the complete set of valid child elements for table cell properties is:

| Element                                         | Description                                                                   | Write Protection? |
| ----------------------------------------------- | ----------------------------------------------------------------------------- | ----------------- |
| `<w:cnfStyle>`                                  | Table conditional formatting properties                                       | No                |
| `<w:tcW>`                                       | Preferred cell width                                                          | No                |
| `<w:gridSpan>`                                  | Grid span (horizontal cell merge)                                             | No                |
| `<w:hMerge>`                                    | Horizontal merge (deprecated)                                                 | No                |
| `<w:vMerge>`                                    | Vertical merge restart/continue                                               | No                |
| `<w:tcBorders>`                                 | Table cell borders (top, left, bottom, right, insideH, insideV, tl2br, tr2bl) | No                |
| `<w:shd>`                                       | Cell background shading                                                       | No                |
| `<w:noWrap>`                                    | Don't wrap cell text                                                          | No                |
| `<w:tcMar>`                                     | Single cell margins                                                           | No                |
| `<w:textDirection>`                             | Cell text flow direction                                                      | No                |
| `<w:tcFitText>`                                 | Fit text into cell width                                                      | No                |
| `<w:vAlign>`                                    | Vertical cell alignment (top, center, bottom)                                 | No                |
| `<w:hideMark>`                                  | Ignore end-of-cell mark in row height calculation                             | No                |
| `<w:headers>`                                   | Associated table header cells                                                 | No                |
| `<w:cellIns>` / `<w:cellDel>` / `<w:cellMerge>` | Revision tracking markers                                                     | No                |

**Finding:** There is no `<w:lock>`, `<w:protect>`, `<w:readOnly>`, or permission attribute in `<w:tcPr>`.

---

## 3. How Microsoft Word Protects Table Cells

Microsoft Word templates and forms achieve protected table layouts through two supported OOXML mechanisms:

### 3.1 Document Forms Protection Mode (`w:documentProtection w:edit="forms"`)

- In Word, selecting **Review > Restrict Editing > Allow only this type of editing: Filling in forms** applies:
  ```xml
  <w:settings>
    <w:documentProtection w:edit="forms" w:enforcement="1"/>
  </w:settings>
  ```
- In this mode, the **entire document** is read-only by default.
- Only regions wrapped in a Content Control (`<w:sdt>`) or Form Field (`<w:fldSimple>` / legacy `<w:ffData>`) remain editable.
- In a structured table form (e.g., an invoice or survey), label cells (e.g., "Customer Name:", "Total:") contain static text and are completely protected against cursor entry, typing, and deletion. Input cells contain Content Controls (e.g. Rich Text or Plain Text SDT) which allow user typing.

### 3.2 Content Control Locking (`w:lock`)

- Any Content Control within a table cell can specify granular locking in `<w:sdtPr>`:
  ```xml
  <w:sdtPr>
    <w:lock w:val="contentLocked"/>
  </w:sdtPr>
  ```
  - `sdtLocked`: The container cannot be deleted, but contents can be edited.
  - `contentLocked`: The contents cannot be edited, but the container can be deleted.
  - `sdtContentLocked`: Both the container and its contents are immutable.

### 3.3 Section-Level Forms Protection (`w:formProt`)

- When documents contain multiple sections (`<w:sectPr>`), Word permits unlocking specific sections while locking others:
  ```xml
  <w:sectPr>
    <w:formProt w:val="0"/> <!-- section is unlocked even in forms mode -->
  </w:sectPr>
  ```

---

## 4. Docen Implementation & Verification

Docen accurately models this architecture across the editing and persistence layers:

1. **Bridge Gate (`packages/editor/src/document/index.ts`):**
   ```ts
   this.#bridge = mountEditBridge({
     canEdit: (ed) => (this.#protectionMode === "forms" ? isInsideEditableSdt(ed) : true),
     ...
   });
   ```
2. **SDT & Lock Resolver (`packages/editor/src/document/protection.ts`):**
   `isInsideEditableSdt(editor)` checks whether the active selection is enclosed by an unlocked SDT node and enforces that `props.cannotEdit !== true`.
3. **Table Interaction Verification:**
   - In normal mode, table cells are fully editable.
   - In forms protection mode, cells without SDTs are write-protected (typing, deleting, and formatting rejected), while cells containing an unlocked SDT accept text input.

---

## 5. Architectural Decision

- **Verdict:** Do NOT implement an artificial `w:lock` attribute on `tableCell` / `<w:tcPr>`.
- **Reasoning:** Standard fidelity with Word is strictly preserved by relying on `documentProtection: { edit: "forms" }` and `<w:sdt>` locks. Adding custom attributes to `tcPr` produces invalid OOXML that third-party processors (MS Word, WPS, LibreOffice) discard or reject.
- **Batch Resolution:** W1.8 is concluded as **N/A for stock `<w:tcPr>`** and verified complete with explicit table cell forms protection test coverage.
