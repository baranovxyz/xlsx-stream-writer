const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../index");
const { getXmlFromXmlStream } = require("../../src/helpers");
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

test("booleans, dates and bigints stay strings on the 0.2 line", async () => {
  // 1.x writes these as typed cells, which is more correct but changes the
  // output of input that already worked. A patch release must not do that, so
  // they keep going through String() here.
  const cells = await cellsFor([true, false, new Date(Date.UTC(1970, 0, 1)), 10n]);
  assert.match(cells, /<c r="A1" t="s"><v>0<\/v><\/c>/);
  assert.doesNotMatch(cells, /t="b"/);
});

test("objects keep their toString, however unhelpful", async () => {
  const decimalLike = { toString: () => "12.340" };
  assert.match(await cellsFor([decimalLike]), /t="s"/);
  // 1.x rejects an object with no meaningful string form; 0.2 still writes it.
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([[{ a: 1 }]]);
  assert.ok((await xlsx.getFile()).length > 0);
});
