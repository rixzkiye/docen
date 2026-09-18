# Bundled production fonts

Metric-compatible open fonts for the Word default families, consumed by
`registerDefaultFonts()` (`src/text/default-fonts.ts`). Shaping is on by
default, so these faces give the layout engine deterministic advances without
depending on machine-installed fonts.

| Word family     | Bundled face     | Upstream                                                           |
| --------------- | ---------------- | ------------------------------------------------------------------ |
| Calibri         | Carlito          | https://github.com/googlefonts/carlito                             |
| Calibri Light   | Carlito          | (no metric-compatible Light face exists; regular metrics are used) |
| Cambria         | Caladea          | https://github.com/huertatipografica/Caladea                       |
| Arial           | Liberation Sans  | https://github.com/liberationfonts/liberation-fonts                |
| Times New Roman | Liberation Serif | https://github.com/liberationfonts/liberation-fonts                |

Each family ships Regular/Bold/Italic/BoldItalic; `ShapedMeasurer` selects the
weight/slant slot and falls back to the regular face when a styled slot was not
registered. All four faces are licensed under the SIL Open Font License 1.1
(`OFL.txt` in each family directory).

## Integrity

`test/font-metrics-golden.json` records the SHA-256 of every `.ttf` here plus
the shaped-run hashes computed from them. Verify on any machine with:

```bash
pnpm exec vp test run packages/layout/src/text/default-fonts.spec.ts
# regenerate after an intentional font change:
DOCEN_FONT_GOLDEN_UPDATE=1 pnpm exec vp test run packages/layout/src/text/default-fonts.spec.ts
```

The browser fetch path and the Node `fs` path resolve the same files; passing
`baseUrl` to `registerDefaultFonts()` points a bundler-managed copy at the
same bytes (the hashes must match the golden).
