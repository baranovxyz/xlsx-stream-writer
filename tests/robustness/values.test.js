const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../dist");
const { getXmlFromXmlStream } = require("../../dist/helpers");
const { assertWellFormedXml } = require("../support/xml");

async function cellsFor(row, options = {}) {
  const xlsx = new XlsxStreamWriter(options);
  xlsx.addRows([row]);
  const xml = await getXmlFromXmlStream(xlsx.sheetXmlStream);
  assertWellFormedXml(assert, xml, "sheet1.xml");
  return xml.slice(xml.indexOf("<row "), xml.indexOf("</row>"));
}

test("blank-ish values become genuinely blank cells", async () => {
  const cells = await cellsFor([null, undefined, NaN]);
  // Previously these were written as t="s" with an empty <v>, a shared-string
  // reference to nothing, which Excel reports as corrupt content.
  assert.equal(cells, '<row r="1"><c r="A1"/><c r="B1"/><c r="C1"/>');
});

test("non-finite numbers become blank rather than literal Infinity", async () => {
  const cells = await cellsFor([Infinity, -Infinity, 0, -0]);
  assert.doesNotMatch(cells, /Infinity/);
  assert.match(cells, /<c r="A1"\/><c r="B1"\/>/);
  assert.match(cells, /<c r="C1" t="n"><v>0<\/v><\/c>/);
});

test("finite numbers are written as numbers", async () => {
  const cells = await cellsFor([1, -2.5, 1e21, Number.MAX_SAFE_INTEGER]);
  assert.match(cells, /<c r="A1" t="n"><v>1<\/v><\/c>/);
  assert.match(cells, /<c r="B1" t="n"><v>-2.5<\/v><\/c>/);
  assert.match(cells, /<c r="D1" t="n"><v>9007199254740991<\/v><\/c>/);
});

test("bigints are written as numbers, not as strings", async () => {
  const cells = await cellsFor([123456789012345678901234567890n]);
  assert.match(cells, /<c r="A1" t="n"><v>123456789012345678901234567890<\/v><\/c>/);
});

test("booleans become boolean cells", async () => {
  const cells = await cellsFor([true, false]);
  assert.equal(cells, '<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c>');
});

test("dates become Excel serial numbers", async () => {
  const cells = await cellsFor([new Date(Date.UTC(1970, 0, 1)), new Date(Date.UTC(2026, 6, 31))]);
  // 1970-01-01 is day 25569 of Excel's epoch, which starts at 1899-12-30.
  assert.match(cells, /<c r="A1" t="n"><v>25569<\/v><\/c>/);
  assert.match(cells, /<c r="B1" t="n"><v>46234<\/v><\/c>/);
});

test("an invalid date is blank rather than NaN", async () => {
  const cells = await cellsFor([new Date("nonsense")]);
  assert.equal(cells, '<row r="1"><c r="A1"/>');
});

test("objects with a real toString are used; opaque ones are rejected", async () => {
  const decimalLike = { toString: () => "12.340" };
  assert.match(await cellsFor([decimalLike]), /<v>0<\/v>/);

  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([[{ a: 1 }]]);
  await assert.rejects(xlsx.getFile(), /Cell A1 received a Object with no meaningful string form/);
});
