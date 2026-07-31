# Migrating from 0.2.x to 1.x

Written to be followed mechanically. Work top to bottom; each section says
whether it is a find-and-replace or a judgement call.

## 0. Decide whether to migrate at all

1.x requires:

- **Node.js 20.19 or newer.**
- **In the browser:** `CompressionStream` (Chrome 80+, Firefox 113+, Safari
  16.4+) and a bundler that honours the `browser` field in `package.json`.

If either is out of reach, stay on the 0.2 line:

```sh
npm install xlsx-stream-writer@legacy
```

`0.2.7` carries the same corruption fixes as 1.x — including the one where a
cell containing the string `constructor` corrupted the whole workbook — and
changes nothing else. It is a safe upgrade from any earlier 0.2.

Everything below applies only if you are moving to 1.x.

## 1. What does not change

No action needed for any of these:

- `new XlsxStreamWriter(options)`, and the `inlineStrings`, `styles` and
  `styleIdFunc` options.
- `addRows(arrayOfArrays)`.
- `getFile()`, still resolving to a `Buffer` in Node and a `Blob` in the browser.
- The generated worksheet for string and finite-number cells, byte for byte.

If your code only uses those, upgrading is `npm install xlsx-stream-writer@latest`
and nothing else.

## 2. Mechanical changes

### 2.1 Deep imports no longer resolve

1.x declares an `exports` map, so only the package root is importable.

```js
// Before — no longer resolves
const { wrapRowsInStream } = require("xlsx-stream-writer/src/helpers");
const rows = wrapRowsInStream(myRows);
xlsx.addRows(rows);

// After — pass the rows directly; addRows accepts arrays, Node streams,
// web ReadableStreams, iterables and async iterables
xlsx.addRows(myRows);
```

`wrapRowsInStream` existed to bridge arrays into streams. `addRows` now does
that itself, so delete the helper and the wrapping call.

### 2.2 The XML stream properties are web streams

`sheetXmlStream` and `sharedStringsXmlStream` were `stream-browserify`
readables. They are now `ReadableStream`s. Only touch this if you read them
directly — most code does not.

```js
// Before
xlsx.sheetXmlStream.pipe(destination);
xlsx.sheetXmlStream.on("data", chunk => { /* ... */ });

// After — Node
const { Readable } = require("node:stream");
Readable.fromWeb(xlsx.sheetXmlStream).pipe(destination);

// After — anywhere
for await (const chunk of xlsx.sheetXmlStream) { /* ... */ }
```

**Order is now enforced.** The shared-string table is built while the worksheet
is walked, so reading `sharedStringsXmlStream` first used to return a table
declaring zero strings. That now raises. Drain the sheet first:

```js
await drain(xlsx.sheetXmlStream);
const table = await drain(xlsx.sharedStringsXmlStream); // only valid now
```

### 2.3 Drop `stream-browserify`

1.x has no runtime dependencies. If `stream-browserify` is in your
`package.json` only because this package needed it, remove it. Likewise any
bundler alias or Node-stream polyfill added for it.

## 3. Behaviour changes — review, do not find-and-replace

These change the *content* of generated cells. Each one is a fix for something
that was wrong before, but if you depended on the old output you must act.

| Cell value | 0.2.x wrote | 1.x writes | Action |
| --- | --- | --- | --- |
| `Date` | text, e.g. `Thu Jan 01 1970 00:00:00 GMT+0000` | Excel date serial | **Add a date format**, or cells display as numbers. See 3.1. |
| `boolean` | text `true` / `false` | boolean cell — Excel shows `TRUE` / `FALSE` | None, unless you relied on the lowercase text. Then map with `String(v)` before adding. |
| `bigint` | text | number | None normally. |
| object without a meaningful `toString` | `[object Object]` | **raises** | Convert the value before adding the row. Objects that define their own `toString` (decimal, date libraries) are unaffected. |
| `null`, `undefined`, `NaN` | a shared-string reference to nothing | blank cell | None — the old output was corrupt. |
| `Infinity`, `-Infinity` | literal `Infinity` | blank cell | None — the old output was corrupt. |
| control characters, unpaired surrogates | written raw; Excel refuses the file | removed | None — the old output was corrupt. |

Additional cases that now raise rather than producing a broken file:

| Situation | 0.2.x | 1.x |
| --- | --- | --- |
| `styleIdFunc` returns a non-integer, a negative, or an index past `styles` | written into the attribute as-is | raises, naming the cell |
| More than 1 048 576 rows or 16 384 columns | a file Excel cannot open | raises, naming the row |
| `addRows()` or `getFile()` called twice on one writer | silent corruption | raises |
| Reading `sheetXmlStream`, then calling `getFile()` | a silently empty workbook | raises |

For the last one: a writer walks its rows exactly once. Use the XML streams *or*
build a workbook, not both, and construct a new writer per workbook.

### 3.1 Formatting dates after the change

A date serial with no number format displays as a number. Declare a format and
point the cell at it:

```js
const xlsx = new XlsxStreamWriter({
  styles: [{ format: "dd.mm.yyyy" }],           // style 1; style 0 is the default
  styleIdFunc: value => (value instanceof Date ? 1 : 0),
});
xlsx.addRows([["When"], [new Date()]]);
```

## 4. Worth adopting once migrated

- **`getStream()`** returns the archive as a `ReadableStream` of bytes, so a
  workbook larger than memory can go straight to disk or to an HTTP response.
  `getFile()` buffers the whole archive.

  ```js
  const { Readable } = require("node:stream");
  const { pipeline } = require("node:stream/promises");
  await pipeline(Readable.fromWeb(xlsx.getStream()), fs.createWriteStream("out.xlsx"));
  ```

- **`compressionLevel`** (0–9, Node only) trades CPU against file size.
- **TypeScript declarations** ship with the package; no `@types` needed.

## 5. Verify

Run your export against a fixture that exercises the changed cases — a `Date`, a
`boolean`, a `null`, and the literal string `constructor` — then:

1. Confirm the file opens without a repair prompt.
2. `unzip -t <file>` reports no errors.
3. Spot-check that date cells display as dates, not as five-digit numbers.

If a cell shows a number where you expected a date, you are missing the format
from 3.1. If the export now throws, read the message: every new exception names
the offending cell or row.
