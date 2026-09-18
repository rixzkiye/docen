# DOCX oracle harness (footnotes, headers/footers, editability fuzz)

Independent evidence that generated DOCX files are editable and stable across
the constructs the R8 audit flagged: footnote/endnote numbering and layout,
headers/footers/page numbers, and the general round-trip margin.

Two phases, one command:

```bash
pnpm editability                     # JS invariants + python/LO/pdftotext oracles
pnpm exec vp test run oracle fuzz    # JS half only (fast, no external tools)
node packages/docx/tests/oracle/run.mjs --skip-fuzz --only hf-logo --verbose
```

## Fixtures (items 6 and 7)

`fixtures/<name>/` holds either `model.json` (DocumentOptions, resolved via
`resolveDocument`) or `input.docx` (an independently authored package parsed
with `parseDOCX`), plus `checks.json`:

```jsonc
{
  "xml":    [{ "part": "word/footnotes.xml", "contains": [...], "order": [[...]], "regex": [...] }],
  "python": [{ "path": "document.sections[0].footnotePr.numFmt", "equals": "upperRoman" }],
  "pdf":    [{ "page": 1, "contains": ["FIRST HEADER"], "regex": ["PAGE\\s+i"] }],
  "pdfStable": true,     // compare rendered page text across generations (default)
  "resaveStable": true   // compare text after LibreOffice resave + docen re-export (default)
}
```

`fixtures.spec.ts` runs each fixture through:

```
fixture ──> gen1 ──parse──> gen2 ──parse──> gen3
```

and asserts, for every fixture:

- **stability** — `gen2` and `gen3` are byte-identical, part for part
  (`roundTrip(roundTrip(x)) == roundTrip(x)` at the stable point);
- **determinism** — generating the same JSON twice yields the same bytes;
- **model equality** — `parse(gen2)` and `parse(gen3)` agree after volatile
  fields (ids, rsids) are normalized;
- **XML expectations** — every `xml` check holds on all three generations;
- no exception anywhere.

`run.mjs` then applies the external oracles over the generated files:

- **python-docx 1.2.0** opens every package; `oracle.py` (lxml) reports
  structural signals — no XML parse errors, no dangling relationships, no
  content-type gaps, then evaluates the fixture `python` checks;
- **LibreOffice headless** renders gen1/gen2/gen3 to PDF (**pdftotext** page
  text must match, and the `pdf` checks must pass);
- **LibreOffice resave** → `parseDOCX` → `generateDOCX` → LibreOffice render:
  an independently written package must round-trip through docen and render
  the same text.

Current matrix (12 fixtures):

| Fixture                  | Covers                                                                                                                        |
| ------------------------ | ----------------------------------------------------------------------------------------------------------------------------- |
| `notes-format-start`     | section `footnotePr`/`endnotePr` numFmt + numStart, separator/continuation-separator ids, multi-paragraph note, table in note |
| `notes-restart-sections` | per-section format/start/restart (`eachSect`, `continuous`) on all three sectPrs                                              |
| `notes-separators`       | custom separator, continuation separator and continuation notice (footnotes + endnotes)                                       |
| `notes-positions`        | `beneathText`/`pageBottom`, `sectEnd`/`docEnd`, explicit `noEndnote` false                                                    |
| `notes-settings-sparse`  | settings-level note defaults + section override, sparse/duplicate ids, table+image inside a note                              |
| `hf-first-page`          | `titlePg` + first/default header and footer                                                                                   |
| `hf-even-odd`            | `evenAndOddHeaders` + even/default headers and footers across 3 pages                                                         |
| `hf-page-numbers`        | `pgNumType` lowerRoman→decimal with start 5                                                                                   |
| `hf-linked-previous`     | absent slots inherit; an explicit blank header stays blank                                                                    |
| `hf-logo`                | header/footer images, media rels, inherited by a linked section                                                               |
| `hf-multi-section`       | three sections, titlePg/even/odd inheritance, per-section page numbering                                                      |
| `hf-foreign`             | python-docx authored input (`titlePg`, even header, `pgNumType` + chapStyle, two sections)                                    |

`fixtures/hf-foreign/input.docx` is produced by
`python3 packages/docx/tests/oracle/author-foreign.py` (python-docx 1.2.0) and
checked in so the harness runs without python-docx.

### LibreOffice limitations observed

LO applies per-section `w:footnotePr` only from the final `sectPr` and ignores
`numRestart` entirely (no `eachPage` restarts); Word honors both. The harness
therefore proves per-section numbering structurally (python/lxml) and proves
format/start values visually on the final section (PDF text) — the LO
behaviour is an oracle limitation, not a docen defect.

LO also exits nonzero (status 1) while still writing a complete PDF for a
2-column section that carries both a footnote reference and an endnote
reference. The minimal repro survives LO's own resave and reproduces with a
plain office-open export (no docen code), so it is an LO layout bug on valid
OOXML: a 2-column section + one paragraph with a footnote reference + one with
an endnote reference. `run.mjs` still fails when the PDF is missing or empty
and records the nonzero-with-PDF case as a warning (see the `loWarnings`
array in `report.json`).

## Fuzz campaign (item 15)

`../fuzz/generator.ts` builds bounded random documents from checked-in seeds
(`../fuzz/seeds.json`) across marks, styles, tables with `columnSpan`/
`verticalMerge`, nested lists, fields (PAGE/NUMPAGES/CREATEDATE), images,
hyperlinks, notes, multi-section layouts, headers/footers and page-numbering
properties. `../fuzz/fuzz.spec.ts` runs the same stability invariants as the
fixtures and bounds each run by `DOCEN_FUZZ_BUDGET_MS` (default 12 s) over
`DOCEN_FUZZ_SAMPLES` (default 24 of 64). Override for a full campaign:

```bash
DOCEN_FUZZ_SAMPLES=64 DOCEN_FUZZ_BUDGET_MS=120000 pnpm exec vp test run fuzz
pnpm exec vp test run oracle fuzz && node packages/docx/tests/oracle/run.mjs --fuzz-limit 64
```

Failures are delta-debugged by a bounded shrinker and written to
`tests/.temp/fuzz/failures/seed-<n>.json` with the minimal reproducing
document, the failure message and the seed — that file becomes the regression
fixture.

`run.mjs` options: `--only a,b`, `--skip-fixtures`, `--skip-fuzz`,
`--fuzz-limit N`, `--fuzz-pdf N` (generations rendered, default 2),
`--fuzz-resave-every N`, `--require` (fail when `soffice`/`pdftotext`/`python3`
are missing instead of skipping), `--verbose`.

Generated artifacts and the LibreOffice profile live in `packages/docx/tests/.temp/`
(gitignored).
