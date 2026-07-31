# Changelog

All notable changes to this package are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [semantic versioning](https://semver.org/spec/v2.0.0.html).

## 0.2.7 — 2026-07-31

Maintenance release for the 0.2 line. Every fix here corrects input that was
already producing a wrong or unopenable file; no input that previously worked
changes its output, so this is safe to take on `^0.2.6` without any migration.

The 1.x line has the same fixes plus a rewritten ZIP layer, typed cells for
dates and booleans, TypeScript declarations and no runtime dependencies. It
requires Node 20.19+ and a modern browser target, which is why this exists.

### Fixed

- **A cell holding the string `"constructor"` corrupted the whole workbook.**
  The shared-string table was a plain object, so values naming a member of
  `Object.prototype` — `constructor`, `toString`, `valueOf`, `hasOwnProperty`,
  `__proto__` — found the inherited member instead of missing. The sheet
  referenced `<v>function Object() { [native code] }</v>` where an index
  belonged, and the string vanished from the table. Present in every release up
  to and including 0.2.6.
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
- **Illegal XML characters corrupted the file.** Control characters and unpaired
  surrogates were written raw, producing a workbook Excel refuses to open. They
  are now removed; tab, newline and carriage return are kept.
- **Blank cells were written as broken shared-string references.** `null`,
  `undefined` and `NaN` produced `t="s"` with an empty `<v>`, a reference to
  nothing.
- **`Infinity` was written literally**, which Excel treats as damaged content.
  Non-finite numbers are now blank.
- **An empty row set produced a malformed sheet** — the worksheet part was
  written without its header.
- Rows beyond Excel's grid limits now raise with the offending row number
  instead of producing a file that will not open.
- `sst/@count` reported the distinct string count rather than the number of
  cells referencing the table.

### Changed

- `addRows` accepts native `node:stream` readables, web `ReadableStream`s,
  iterables and async iterables in addition to arrays. Previously an
  `instanceof` check against `stream-browserify`'s class rejected the stream
  type Node callers actually have. Purely additive.
- `jszip` floor raised to `^3.10.1`, `stream-browserify` to `^3.0.0`.
- Tests run on the built-in `node:test` runner; the development dependency tree
  is empty.

### Unchanged on purpose

Booleans, dates, bigints and objects still go through `String()`. Writing them
as typed cells is more correct and is what 1.x does, but it changes the output
of input that already worked, which a patch release must not do.

## 0.2.6 and earlier

See the git history.
