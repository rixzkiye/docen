# Deterministic DOCX generation

`generateDOCX`, `generateDOCXSync` and `generateDOCXStream` are byte-reproducible:
the same input produces the same bytes in the same process and in a fresh one.
This document records the seam that makes that true, so upstream upgrades don't
silently reintroduce nondeterminism.

## Nondeterminism sources and the fix

| Source                                                           | Fix                                                                                                                                                          |
| ---------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `docProps/core.xml` `created`/`modified` default to `Date.now()` | `generateDOCX*` fills missing dates from the generation clock (`DOCX_EPOCH` by default, `options.date` to override, `null` to omit). Source dates still win. |
| Comment `w:date` defaults to `Date.now()`                        | The patched `stringifyComment` reads `scope.now`.                                                                                                            |
| `wp:docPr/@id` counter (`_docPropsIdGen`, module-global)         | The patched `stringifyDocPr` reads `scope.nextDocPr()`.                                                                                                      |
| VML `_x0000_s####` counter (`nextSpid`, module-global)           | The patched `nextVmlShapeId` reads `scope.nextVml()`.                                                                                                        |
| `crypto.randomUUID()` for font keys and SmartArt GUIDs           | The patched `uniqueUuid` / smartart `uuid()` read `scope.nextUuid()`.                                                                                        |
| `crypto.getRandomValues()` for altChunk/subDoc part names        | The patched `uniqueId` reads `scope.nextToken()`.                                                                                                            |
| ZIP entry mtime (`fflate` uses `Date.now()` per entry)           | The patched ZIP writers pass a fixed DOS-epoch timestamp (`1980-01-01` local, i.e. time `0`, date `0x21`) on every path.                                     |
| ZIP entry order                                                  | The compiler emits `Object.keys(files)` order; the native writer and fflate both preserve it.                                                                |

The native Node/Bun ZIP writer already wrote zeroed timestamps; the patch also
normalizes the browser/Deno `fflate` fallback (and the incremental ZIP stream)
so bytes are stable on every runtime. A reproducibility test can force the
pure-JS path on Node with
`globalThis[Symbol.for("docen.ooxml.force-js-deflate")] = true`.

## The scope contract

`options.date` opens a scope (in `converters/determinism.ts`) under
`globalThis[Symbol.for("docen.ooxml.determinism")]` for the synchronous compile
phase of one generation, then restores the previous scope:

```ts
interface GenerationScope {
  now: string | undefined; // ISO-8601 date default (undefined = omit dates)
  nextDocPr(): number; // wp:docPr/@id, starting at 1
  nextVml(): number; // VML shape id, starting at 1025
  nextUuid(): string; // RFC 4122 v4-shaped UUID
  nextToken(): string; // 21-char lowercase token (part names)
}
```

The upstream packages are patched (`patches/@office-open__core@0.14.5.patch`,
`patches/@office-open__docx@0.14.5.patch`) to read this scope; outside a scope
their original counter/random/time behavior is unchanged. Because upstream
materializes every package part before its first `await`, no two concurrent
generations can interleave inside the scope — `Promise.all` exports stay
isolated (see `converters/concurrency.spec.ts`).

## Verification

```bash
pnpm --filter @docen/docx build
pnpm exec vp test run packages/docx/src/converters/determinism.spec.ts \
  packages/docx/src/converters/determinism-process.spec.ts
```

- `determinism.spec.ts` — same-process async/sync/stream byte equality, fixed
  dates, per-generation drawing ids, input non-mutation.
- `determinism-process.spec.ts` — two fresh child processes
  (`tests/determinism-child.mjs`) hash the same bytes, on both the native and
  the forced pure-JS ZIP path; the JS path's ZIP timestamps are asserted to be
  the fixed epoch and its entry order is asserted stable.
