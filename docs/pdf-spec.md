# PDF Specification & Parity Decisions (R12-W2)

This document formalizes the mapping, standards conformance decisions, and explicit exclusions for PDF generation in Docen, specifically regarding page labels, numbering schemes, and viewer preferences.

## 1. Page Labels (`/PageLabels`) & Numbering Format (`w:numFmt`)

### Standards Reference

- **ISO 32000-1:2008 §12.4.2** (Page Labels) defines the entries in a page label dictionary (`Table 159`).
- Permissible numbering styles (`/S` key):
  - `/D`: Decimal arabic numerals (1, 2, 3, …)
  - `/R`: Uppercase roman numerals (I, II, III, …)
  - `/r`: Lowercase roman numerals (i, ii, iii, …)
  - `/A`: Uppercase letters (A, B, C, …)
  - `/a`: Lowercase letters (a, b, c, …)
  - _(Absent `/S`)_: No numeric numbering; only the label prefix string (`/P`) is displayed.

### OOXML Mapping (`w:pgNumType`)

OOXML section properties (`w:sectPr/w:pgNumType`) define page numbering style, start offsets, and chapter prefix attachments:

- `w:fmt`: maps to PDF `/S` style codes:
  - `decimal` -> `/D`
  - `upperRoman` -> `/R`
  - `lowerRoman` -> `/r`
  - `upperLetter` -> `/A`
  - `lowerLetter` -> `/a`
  - `none` -> omits `/S` (displays prefix only or unnumbered page)
- `w:start`: maps to `/St` (first page number in range, omitted when 1).
- `w:prefix`: explicit prefix string emitted as `/P (...)`.
- `w:chapStyle` + `w:chapSep` (or `separator`):
  - `w:chapSep` enums: `hyphen` ("-"), `period` ("."), `colon` (":"), `emDash` ("—"), `enDash` ("–"), or verbatim characters.
  - Carried via `ProjectedPageNumbering` (`chapterStyle`, `separator`, `prefix`).

### Exotic `numFmt` Fallback Decision

- **Problem**: OOXML supports many exotic numbering schemes (e.g. `chineseCounting`, `chineseLegalSimplified`, `taiwaneseCounting`, `ideographLegalTraditional`, `ordinal`, `cardinalText`, `aiueo`, `iroha`, `katakana`, `ganada`, `numberInDash`, etc.).
- **Constraint**: ISO 32000-1 §12.4.2 provides no mechanism or algorithm for arbitrary glyph-run label numbering generators in `/PageLabels`. PDF viewers (Adobe Acrobat, Chromium PDFium, Apple Preview, PDF.js) only implement the 5 standard ISO sequences for `/S`.
- **Decision**: All exotic/unsupported numbering formats degrade gracefully to `decimal` (`/D`) (or `none` when explicitly unnumbered). If chapter prefixes or custom label strings are specified, they are preserved via `/P`.

---

## 2. Viewer Preferences (`/ViewerPreferences`)

### Standards Reference

- **ISO 32000-1:2008 §12.2** (Viewer Preferences) defines window configuration flags for PDF interactive viewing:
  - `HideToolbar`: Hide viewer application toolbars.
  - `HideMenubar`: Hide viewer application menu bars.
  - `HideWindowUI`: Hide viewer UI chrome.
  - `FitWindow`: Resize document window to fit page size.
  - `CenterWindow`: Position window at center of screen.
  - `DisplayDocTitle`: Display document title from `/Info` / `dc:title` instead of filename in window title bar.

### Evaluation of OOXML (WordprocessingML) Sources

An audit of ISO/IEC 29500-1 / ECMA-376 (WordprocessingML `settings.xml`, `document.xml`, `webSettings.xml`) shows:

- Word stores editing and document views via `w:view` (`none`, `print`, `outline`, `masterPages`, `normal`, `web`) and `w:zoom`.
- Word does **not** store or manage PDF viewer window chrome preferences (`HideToolbar`, `HideMenubar`, `FitWindow`, `CenterWindow`).

### Written Exclusion Decision

- **Decision**: Because Word/DOCX contains no document source or authoring intent for PDF window chrome preferences, they are **not wired from DOCX documents** (explicit written exclusion).
- **Low-level API availability**: `options.viewerPreferences` in `packages/editor/src/document/export-pdf.ts` remains available for programmatic and backend callers.
- **PDF/UA-1 Requirement**: `/ViewerPreferences << /DisplayDocTitle true >>` is mandatory for PDF/UA-1 accessibility compliance (ISO 14289-1 §7.1). This flag will be systematically wired in Lane **Q1** when `--pdfa ua` is enabled.
