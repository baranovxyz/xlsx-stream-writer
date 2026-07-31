const test = require("node:test");
const assert = require("node:assert/strict");

const JSZip = require("jszip");
const XlsxStreamWriter = require("../../index");
const { readZip } = require("../support/unzip");
const { BUFFER_LIMIT } = require("../../src/zip/writer");

// Enough rows to push the worksheet part past the writer's buffering limit, so
// the genuinely streamed code path runs — not the injectable one the unit tests
// use.
const ROW_COUNT = 120000;

function* generateRows() {
  yield ["Name", "Location", "Amount", "Active"];
  for (let i = 1; i < ROW_COUNT; i++) {
    yield [`Customer number ${i}`, `City of origin ${i % 997}`, i * 1.5, i % 2 === 0];
  }
}

test("a workbook past the buffering limit is still a valid archive", { timeout: 120000 }, async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(generateRows());
  const buffer = await xlsx.getFile();

  const zip = readZip(buffer);
  const sheet = zip.entry("xl/worksheets/sheet1.xml");

  assert.ok(
    sheet.uncompressedSize > BUFFER_LIMIT,
    `sheet should exceed the ${BUFFER_LIMIT} byte limit, was ${sheet.uncompressedSize}`,
  );
  assert.equal(sheet.usesDataDescriptor, true, "an oversized part must defer its sizes");

  // readZip verifies the CRC and length of every entry as it parses, so this
  // reaching the assertions at all is most of the check.
  assert.match(zip.text("xl/worksheets/sheet1.xml"), /<\/sheetData><\/worksheet>$/);
  assert.match(zip.text("xl/sharedStrings.xml"), /<\/sst>$/);

  // And an implementation with no shared lineage agrees.
  const viaJsZip = await JSZip.loadAsync(buffer);
  const sheetXml = await viaJsZip.file("xl/worksheets/sheet1.xml").async("string");
  assert.match(sheetXml, new RegExp(`<row r="${ROW_COUNT}">`));
});

test("getStream() emits bytes before the rows run out", { timeout: 120000 }, async () => {
  let rowsProduced = 0;
  function* counted() {
    for (const row of generateRows()) {
      rowsProduced++;
      yield row;
    }
  }

  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(counted());

  const reader = xlsx.getStream().getReader();
  const first = await reader.read();
  const producedWhenFirstChunkArrived = rowsProduced;

  assert.equal(first.done, false);
  assert.ok(first.value.length > 0);
  // The point of the package: bytes come out while rows are still going in,
  // rather than after the whole workbook has been assembled in memory.
  assert.ok(
    producedWhenFirstChunkArrived < ROW_COUNT,
    `expected output before all ${ROW_COUNT} rows were read, but ${producedWhenFirstChunkArrived} had been consumed`,
  );

  // Drain the rest so the workbook is complete and internally consistent.
  let total = first.value.length;
  while (true) {
    const next = await reader.read();
    if (next.done) break;
    total += next.value.length;
  }
  assert.ok(total > 0);
  assert.equal(rowsProduced, ROW_COUNT);
});
