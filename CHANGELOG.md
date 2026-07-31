# Changelog

All notable changes to this package are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [semantic versioning](https://semver.org/spec/v2.0.0.html).

## 1.3.1 — 2026-07-31

Documentation and verification only. `src/` has not changed since 1.3.0, so the
generated workbook is byte for byte what 1.3.0 produced — which the golden
fixtures assert rather than merely claim.

### Documentation

- **The `0.2.7` entry below has been corrected.** It described a release that
  was never published: the corruption fixes were backported into that version
  before it reached the registry, so "no API or output changes" and the
  `engines.node` floor it announced were both untrue of the package on npm. The
  entry now records what actually shipped, and says that it was corrected.
- Everything below `0.2.7` is deprecated on npm, worded to make clear it is a
  data-correctness problem rather than a security one. Installing those versions
  never exposed a caller to a vulnerable dependency — the `jszip` range floats,
  so a fresh install resolved the fixed version regardless.

### Verification

- Generated workbooks are now checked against the Open XML schema on every CI
  run, using Microsoft's own validator from the Open XML SDK. This is a stricter
  question than the existing LibreOffice check, which is deliberately tolerant:
  it answers whether the file conforms to the format as specified, rather than
  whether one reader accepts it. Six workbooks pass against the Office2007
  schema, the oldest an `.xlsx` can target.

## 1.3.0 — 2026-07-31

Found by a second adversarial review pass, plus documentation the earlier
releases should have carried.

### Changed

- **Reading `sharedStringsXmlStream` before the worksheet now raises.** The
  table is filled as the sheet is walked, so reading it early produced a part
  declaring a count of zero and listing nothing — valid-looking and quietly
  wrong. Draining `sheetXmlStream` first, which is the documented order, is
  unaffected; so is building a workbook, since the archive writer consumes
  entries in order.
- `styleIdFunc` rejections name the type as well as the value. A boxed `Number`
  or a numeric string prints like a perfectly good index, which made the bare
  value baffling.

### Documentation

- **A 0.2.x → 1.x migration guide**, `docs/migrating-to-1.x.md`, written to be
  followed step by step: what needs no change, what is a find-and-replace, and
  which behaviour changes need a human or an agent to look at the cells. It
  ships inside the package, so it is readable from `node_modules` without the
  repository.
- The release runbook covers the 0.2 maintenance line: why a backport must be a
  `0.2.z` rather than a `0.3.0`, what may and may not be backported, and how the
  `legacy` dist-tag keeps a maintenance release from becoming the default
  install.
- The README's plans list no longer mixes finished work with notes from 2018.

### Also on npm

`0.2.7` ships the corruption fixes to the 0.2 line under the `legacy` dist-tag,
for projects on `^0.2.6` that cannot take the 1.x requirements. Its own entry
below records what changed.

## 1.2.0 — 2026-07-31

Found by an adversarial review pass over the preceding phases.

### Fixed

- **A cell holding the string `"constructor"` corrupted the whole workbook.**
  The shared-string table was a plain object, so values that name a member of
  `Object.prototype` — `constructor`, `toString`, `valueOf`, `hasOwnProperty`,
  `__proto__` — found the inherited member instead of missing. The sheet then
  referenced `<v>function Object() { [native code] }</v>` where an index
  belonged, and the string table came out empty. This affects **every
  previously published version, 0.2.6 included.**
- The same collision applied to style `format` and `fill` values, producing
  `numFmtId="function Object() { [native code] }"`.
- Reading `sheetXmlStream` and then calling `getFile()` produced a **silently
  empty workbook**: the stream consumes the rows, and rows can only be walked
  once. Both now raise instead.

### Changed

- `styleIdFunc` return values are validated. A non-integer, a negative number,
  or an index past the declared styles now raises with the offending cell
  address. Previously the value was interpolated straight into the `s`
  attribute, so a string could close the attribute and inject another, and an
  out-of-range index produced a file Excel offers to repair.

### Added

- Direct tests for the ZIP64 central-directory and end-of-archive records. Those
  branches trigger past 4 GiB and past 65 535 entries, which no fixture can
  reach in reasonable time, so the record builders are exercised directly.
- A test that concurrent writers stay isolated.

## 1.1.0 — 2026-07-31

The package is written in TypeScript and ships type declarations. No behaviour
changes: the generated workbook is byte-for-byte identical to 1.0.0, which the
golden tests assert.

### Added

- Type declarations for the whole public surface, so no `@types` package is
  needed. `CellValue`, `Row`, `CellStyle` and `XlsxStreamWriterOptions` are
  exported alongside the class.
- Source maps and the TypeScript sources ship with the package, so stack traces
  and go-to-definition land on real code.

### Changed

- Sources moved from `src/*.js` to `src/*.ts`; the published entry point is now
  `dist/index.js`. `require("xlsx-stream-writer")` is unaffected.
- Tests run against the compiled `dist/`, so they exercise the artifact that
  actually ships rather than the sources it was built from.
- The build is `tsc` alone — no bundler. `typescript` and `@types/node` are the
  only additions, both development-only.

## 1.0.0 — 2026-07-31

**The package now has no runtime dependencies.** `jszip` and
`stream-browserify` are gone, and with them the 15 packages they resolved to.
The ZIP container is written here, compressing through `node:zlib` on the server
and `CompressionStream` in the browser, selected by the `browser` field.

The 1.0.0 marks a supported API, not a rewrite of the interface: `addRows` and
`getFile` behave as before.

### Added

- `getStream()` returns the archive as a `ReadableStream` of bytes, so a
  workbook larger than memory can be piped straight to disk or to a response.
  `getFile()` buffers the whole archive, which capped the size the package could
  produce at available memory — the one thing a "stream writer" should not do.
- ZIP64 records on parts that need them. Above 4 GiB the previous output was
  silently corrupt; a 500 000-row export already produces 80 MB of worksheet XML,
  so the ceiling was reachable.
- `compressionLevel` option, 0-9. Node only; browsers expose no level control.
- Archives are reproducible: entry timestamps are fixed rather than "now", so
  the same rows always produce the same bytes.

### Changed

- **Breaking.** `sheetXmlStream` and `sharedStringsXmlStream` are now web
  `ReadableStream`s rather than `stream-browserify` readables. `addRows` accepts
  the same range of inputs as before, native Node streams included.
- **Breaking.** `getFile()` and `getStream()` reject rather than throw when
  called before `addRows()` or a second time.
- **Breaking.** `helpers.wrapRowsInStream` and `helpers.toRowsStream` are gone;
  `addRows` accepts arrays, streams and iterables directly.
- The archive no longer contains directory entries. They are optional in ZIP and
  absent from the workbooks Excel itself writes; JSZip's habit of emitting them
  was the only reason they were there.
- Small workbooks now use plain ZIP entries with exact sizes in the local
  header, instead of the data descriptors JSZip emitted for everything. Only
  parts above 8 MiB — where the final size cannot be known in advance — use the
  streamed form.
- Shared strings are now generated after the sheet by contract rather than by
  accident. The previous ordering depended on JSZip consuming entries in
  insertion order, an undocumented detail of a third-party library.

### Verification

Output is checked against three readers with no shared lineage with the writer:
the test suite's own central-directory parser, JSZip, and the system `unzip`.

## 0.3.0 — 2026-07-31

Correctness and robustness. No dependency changes. Output for the documented
happy path is unchanged; the differences below all concern inputs that
previously produced a corrupt workbook, a wrong one, or a hang.

### Fixed

- **Styles leaked between writers.** `getStyles` appended to module-level
  arrays, so the second workbook built in a process inherited the first one's
  fills and cell formats, and every style id after the first pointed at the
  wrong entry.
- **Options leaked between writers.** The constructor merged into the shared
  defaults object, so `new XlsxStreamWriter()` could inherit an earlier
  writer's `inlineStrings` and `styles`.
- **Native Node streams were rejected.** `addRows` tested `instanceof` against
  `stream-browserify`'s class, so passing a `node:stream` Readable — the stream
  type Node callers actually have — threw "Argument must be an array of arrays".
  It now also accepts web `ReadableStream`s, iterables and async iterables.
- **Source-stream errors hung forever.** `.pipe()` does not forward errors, so a
  failing source left `getFile()` pending indefinitely. It now rejects.
- **Illegal XML characters corrupted the file.** Control characters and unpaired
  surrogates were written raw, producing a workbook Excel refuses to open. They
  are now removed; tab, newline and carriage return are kept.
- **An empty row set produced a malformed sheet.** With no rows, the worksheet
  part was written without its header — a bare closing tag.
- **Blank cells were written as broken shared-string references.** `null`,
  `undefined` and `NaN` produced `t="s"` with an empty `<v>`, a reference to
  nothing. They are now genuinely blank cells.
- **`Infinity` was written literally**, as `<v>Infinity</v>`, which Excel treats
  as damaged content. Non-finite numbers are now blank.
- `sst/@count` reported the distinct string count; it now reports the number of
  cells referencing the table, with `uniqueCount` reporting distinct values.

### Changed

- `boolean` values become real boolean cells (`t="b"`), so Excel shows `TRUE` and
  `FALSE` rather than the text "true" and "false".
- `Date` values become Excel date serials instead of
  `"Thu Jan 01 1970 00:00:00 GMT+0000 (Coordinated Universal Time)"`. Apply a
  date `format` style to display them as dates.
- `bigint` values are written as numbers rather than strings.
- Values whose only `toString` is `Object.prototype`'s now raise instead of
  writing `[object Object]` into the sheet. Objects that define their own
  `toString` — decimal and date libraries — keep working.
- Rows beyond 1 048 576, and rows wider than 16 384 cells, raise with the
  offending row number instead of producing a file Excel will not open.
- `addRows()` and `getFile()` each raise if called twice on the same writer.
  Calling them twice previously discarded work or produced a corrupt result
  silently.
- Invalid `options.styles` and `options.styleIdFunc` are rejected at
  construction rather than failing later.

## 0.2.7 — 2026-07-31

Published from the `0.2.x` branch under the `legacy` dist-tag, for projects on
`^0.2.6` that cannot take the 1.x requirements. It carries the corruption fixes
backported from 1.x as well as the security and tooling work below.

**This entry has been corrected.** It first read "Security and tooling only. No
API or output changes: the generated workbook is byte-for-byte identical to
0.2.6." The backport landed in the same version before 0.2.7 was published, so
that never described the release actually on the registry: every fix below
changes the output of the input that was hitting it. The `engines.node` floor
the entry announced went the same way — the published 0.2.7 declares none.

### Fixed

- **A cell holding the string `"constructor"` corrupted the whole workbook.**
  The shared-string table was a plain object, so values naming a member of
  `Object.prototype` — `constructor`, `toString`, `valueOf`, `hasOwnProperty`,
  `__proto__` — found the inherited member instead of missing. The sheet
  referenced `<v>function Object() { [native code] }</v>` where an index
  belonged, and the string table came out empty. Present in every release up to
  and including 0.2.6.
- The same collision applied to style `format` and `fill` values, producing a
  `numFmtId` or `fillId` that was a stringified function.
- **Styles leaked between writers.** `getStyles` appended to module-level
  arrays, so the second workbook built in a process inherited the first one's
  fills and every style id after the first pointed at the wrong entry.
- **Options leaked between writers.** The constructor merged into the shared
  defaults object, so a later `new XlsxStreamWriter()` could inherit an earlier
  writer's `inlineStrings` and `styles`.
- **Source-stream errors hung forever.** `.pipe()` does not forward errors, so a
  failing source left `getFile()` pending indefinitely. It now rejects.
- **Illegal XML characters corrupted the file.** Control characters were
  written raw, producing a workbook Excel refuses to open. They are now
  removed; tab, newline and carriage return are kept.
- **Blank cells were written as broken shared-string references.** `null`,
  `undefined` and `NaN` produced `t="s"` with an empty `<v>`, a reference to
  nothing. They are now genuinely blank cells.
- **`Infinity` was written literally**, which Excel treats as damaged content.
  Non-finite numbers are now blank.
- **An empty row set produced a malformed sheet** — the worksheet part was
  written without its header.
- Rows beyond 1 048 576, and rows wider than 16 384 cells, now raise with the
  offending row number instead of producing a file Excel will not open.
- `sst/@count` reported the distinct string count; it now reports the number of
  cells referencing the table.

### Security

- Cleared all 59 known advisories (11 critical, 19 high) reported by
  `npm audit` against this repository's checkout of 0.2.6. Almost all of them
  came in through `jest@25`, and the lockfile is what pinned them. **Installing
  0.2.6 never exposed a caller to those advisories**: `^3.1.5` floats, so a
  fresh install resolves the fixed `jszip`, and npm ignores a dependency's
  lockfile. Nothing below 0.2.7 is deprecated for a security reason — see
  `SECURITY.md`.
- `jszip` moved from `^3.1.5` to `^3.10.1`, so the range no longer *reaches*
  the prototype-pollution and path-traversal advisories against 3.1.5, and the
  lockfile stops pinning a version that has them.
- `stream-browserify` moved from `^2.0.2` to `^3.0.0`.
- Removed `jest@25` and the unused `crc` devDependency, taking the development
  dependency tree from 576 packages to zero. Tests now run on the built-in
  `node:test` runner.
- Regenerated `package-lock.json` at lockfile v3; the previous v1 lockfile was
  why the resolved tree stayed four years stale.

### Added

- Continuous integration across Node 20.19, 22 and 24, with `npm audit` at
  `--audit-level=low` and a full-history secret scan.
- Release pipeline publishing through npm OIDC trusted publishing with SLSA
  provenance, including registry integrity and provenance verification after
  publish. See `docs/RELEASING.md`.
- Test suite now inspects the generated `.xlsx` archive itself — entry order,
  per-part bytes, XML well-formedness, CRCs and streaming-mode flags — rather
  than only in-memory streams.
- `LICENSE` file. The package has always declared MIT but shipped no license
  text.

### Changed

- `addRows` also accepts native `node:stream` readables, web `ReadableStream`s,
  iterables and async iterables. An `instanceof` check against
  `stream-browserify`'s class had rejected the stream type Node callers
  actually have. Purely additive.
- `npm test` runs once and exits, so it works in CI. Use `npm run test:watch`
  for the previous watching behaviour.
- The published tarball now contains only `index.js`, `src/`, `README.md`,
  `CHANGELOG.md` and `LICENSE`. Previous releases also shipped `tests/`,
  `examples/` and the lockfile.

### Removed

- `examples/crc_speed_test.js`, a scratch benchmark that depended on the
  now-removed `crc` package and had a latent bug of its own.

### Unchanged on purpose

Booleans, dates, bigints and objects still go through `String()`. Writing them
as typed cells is what 1.x does, but it changes the output of input that
already worked, which a patch release must not do. The once-only `addRows` and
`getFile` guards stayed on 1.x for the same reason.

## 0.2.6 and earlier

See the git history.
