const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../index");

const MAX_ROWS = 1048576;
const MAX_COLUMNS = 16384;

test("a row at the column limit is accepted", { timeout: 15000 }, async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([new Array(MAX_COLUMNS).fill(1)]);
  const buf = await xlsx.getFile();
  assert.ok(buf.length > 0);
});

test("a row past the column limit is rejected by column count", { timeout: 15000 }, async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([new Array(MAX_COLUMNS + 1).fill(1)]);
  await assert.rejects(
    xlsx.getFile(),
    /Row 1 has 16385 cells, over the Excel worksheet limit of 16384 columns/,
  );
});

// Streaming a million rows to prove the guard would cost minutes; the guard
// itself is what needs testing, so drive it directly.
test("a row past the row limit is rejected by row number", () => {
  const xlsx = new XlsxStreamWriter();
  assert.doesNotThrow(() => xlsx._getRowXml(["a"], MAX_ROWS - 1));
  assert.throws(
    () => xlsx._getRowXml(["a"], MAX_ROWS),
    /Row 1048577 exceeds the Excel worksheet limit of 1048576 rows/,
  );
});

test("the last legal cell address is the one Excel expects", () => {
  const { getCellAddress } = require("../../src/helpers");
  assert.equal(getCellAddress(1, 1), "A1");
  assert.equal(getCellAddress(1, 26), "Z1");
  assert.equal(getCellAddress(1, 27), "AA1");
  assert.equal(getCellAddress(1, 702), "ZZ1");
  assert.equal(getCellAddress(1, 703), "AAA1");
  assert.equal(getCellAddress(MAX_ROWS, MAX_COLUMNS), `XFD${MAX_ROWS}`);
});
