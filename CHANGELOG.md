# Changelog

All notable changes to this package are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [semantic versioning](https://semver.org/spec/v2.0.0.html).

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

Security and tooling only. No API or output changes: the generated workbook is
byte-for-byte identical to 0.2.6.

### Security

- Cleared all 59 known advisories (11 critical, 19 high) reported by
  `npm audit` against the 0.2.6 dependency tree.
- `jszip` moved from `^3.1.5` to `^3.10.1`, picking up fixes for the
  prototype-pollution and path-traversal advisories affecting the pinned 3.1.5.
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

- `npm test` runs once and exits, so it works in CI. Use `npm run test:watch`
  for the previous watching behaviour.
- The published tarball now contains only `index.js`, `src/`, `README.md`,
  `CHANGELOG.md` and `LICENSE`. Previous releases also shipped `tests/`,
  `examples/` and the lockfile.
- Declared `engines.node` as `>=20.19.0`.

### Removed

- `examples/crc_speed_test.js`, a scratch benchmark that depended on the
  now-removed `crc` package and had a latent bug of its own.

## 0.2.6 and earlier

See the git history.
