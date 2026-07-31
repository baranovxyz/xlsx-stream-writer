# xlsx-stream-writer

Create `.xlsx` files in streaming mode, in the browser and in Node.js.

Built for one job: writing very large spreadsheets with simple formatting,
without holding all the rows in memory.

**No dependencies.** The ZIP container is written by this package, compressing
through `node:zlib` on the server and `CompressionStream` in the browser.

Rewritten from the CoffeeScript
[node-xlsx-writer](https://github.com/rubenv/node-xlsx-writer) and changed to run
in both environments. The API is completely different from that implementation.

## Install

```sh
npm install xlsx-stream-writer
```

Requires Node.js 20.19 or newer. In the browser it needs `CompressionStream`
(Chrome 80+, Firefox 113+, Safari 16.4+) and a bundler that honours the
`browser` field, which is all of them.

## Rows from an array

```javascript
const XlsxStreamWriter = require("xlsx-stream-writer");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const xlsx = new XlsxStreamWriter();
xlsx.addRows(rows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
```

`getFile()` resolves to a `Buffer` in Node.js and a `Blob` in the browser.

## Rows from a stream

`addRows` accepts an array, a Node.js readable stream, a web `ReadableStream`, or
any iterable or async iterable of rows — so you can feed it a database cursor or
a generator directly:

```javascript
const XlsxStreamWriter = require("xlsx-stream-writer");
const fs = require("fs");

async function* fetchRows() {
  yield ["Name", "Location"];
  for await (const record of database.stream("SELECT name, location FROM city")) {
    yield [record.name, record.location];
  }
}

const xlsx = new XlsxStreamWriter();
xlsx.addRows(fetchRows());

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
```

If the source fails part-way through, `getFile()` rejects with that error.

## Streaming the workbook out

`getFile()` builds the whole archive in memory. For workbooks too large for
that, `getStream()` returns a `ReadableStream` of the archive bytes, so rows go
in while bytes come out:

```javascript
const { Readable } = require("node:stream");
const { pipeline } = require("node:stream/promises");
const fs = require("node:fs");

const xlsx = new XlsxStreamWriter();
xlsx.addRows(generateRows());

await pipeline(
  Readable.fromWeb(xlsx.getStream()),
  fs.createWriteStream("large.xlsx"),
);
```

Writing 500 000 rows × 4 columns this way produces a 12 MB file — 80 MB of
worksheet XML — in about 30 seconds, with peak memory around 240 MB. Most of
that is the shared-string table; with `inlineStrings: true` the same export peaks
near 120 MB and takes half the time, at the cost of a slightly larger file. If
your strings are mostly distinct, prefer inline strings.

## Cell values

| Value                        | Written as                                        |
| ---------------------------- | ------------------------------------------------- |
| `string`                     | shared string, or an inline string                 |
| `number` (finite), `bigint`  | number                                             |
| `boolean`                    | boolean — Excel shows `TRUE` / `FALSE`             |
| `Date`                       | Excel date serial; apply a date `format` style to display it as a date |
| `null`, `undefined`, `NaN`   | blank cell                                         |
| `Infinity`, `-Infinity`      | blank cell — Excel has no representation for them  |
| anything else                | its `toString()`; objects that have no meaningful one raise an error |

Characters that XML 1.0 cannot represent — most control characters, unpaired
surrogates — are removed, since leaving them in produces a file Excel refuses to
open.

A worksheet is limited to 1 048 576 rows and 16 384 columns. Exceeding either
raises rather than producing a file that will not open.

## Options

```javascript
const xlsx = new XlsxStreamWriter({
  // Write strings into the sheet directly instead of into a shared-string
  // table. Larger output, but no string table to hold in memory.
  inlineStrings: false,

  // Cell formats, referenced by index from styleIdFunc. Index 0 is the
  // implicit default, so the first entry here is style 1.
  styles: [{ fill: "FFFF0000" }, { format: "dd.mm.yyyy" }],

  // Choose a style per cell.
  styleIdFunc: (value, columnIndex, rowIndex) => (rowIndex === 0 ? 1 : 0),

  // Deflate level, 0-9. Node only; browsers expose no level control.
  compressionLevel: 4,
});
```

`fill` is an ARGB colour like `FFFF0000`. `format` is an Excel number format
string like `0.00` or `dd.mm.yyyy`.

## Lifecycle

A writer builds one workbook: call `addRows` once, then `getFile` or `getStream`
once. Calling either again raises, because the row stream has been consumed.
Create a new `XlsxStreamWriter` for the next workbook.

Archives are reproducible — entry timestamps are fixed rather than "now", so the
same rows always produce the same bytes.

## Plans

- [ ] improve api
- [x] add tests
- [x] replace JSZip with a built-in zip writer, so the package has no runtime dependencies
- [ ] ship TypeScript sources and type declarations
- [ ] make browser build, put on some cdn
- [ ] optimize shared string stuff
- [ ] maybe use web workers to build xlsx in browser
- [ ] maybe implement some specifis for nodejs

## Security

Reporting and threat model: [SECURITY.md](SECURITY.md).

## License

MIT
