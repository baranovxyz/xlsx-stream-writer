// Compile-only check that the published declarations are usable. This file is
// never executed; `npm run test:types` type-checks it against dist/index.d.ts,
// which is what consumers actually see.

import XlsxStreamWriter = require("../../dist");

const styles: XlsxStreamWriter.CellStyle[] = [
  { fill: "FFFF0000" },
  { format: "dd.mm.yyyy" },
];

const writer = new XlsxStreamWriter({
  inlineStrings: false,
  styles,
  styleIdFunc: (value: XlsxStreamWriter.CellValue, columnId: number, rowId: number) =>
    rowId === 0 ? 1 : 0,
  compressionLevel: 4,
});

const rows: XlsxStreamWriter.Row[] = [
  ["Name", "Location"],
  ["Alpha", 1, true, new Date(), null, undefined, 10n],
];

writer.addRows(rows);

// Each accepted input shape should type-check.
declare const webStream: ReadableStream<XlsxStreamWriter.Row>;
declare function generate(): AsyncGenerator<XlsxStreamWriter.Row>;
new XlsxStreamWriter().addRows(webStream);
new XlsxStreamWriter().addRows(generate());
new XlsxStreamWriter().addRows(new Set<XlsxStreamWriter.Row>());

async function build(): Promise<void> {
  const file: Buffer | Blob = await writer.getFile();
  void file;

  const stream: ReadableStream<Uint8Array> = new XlsxStreamWriter().getStream();
  void stream;
}

void build;

// Options are optional.
new XlsxStreamWriter();
