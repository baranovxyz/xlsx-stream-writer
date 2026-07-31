const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../index");
const { rows } = require("../helpers");
const { getXmlFromXmlStream } = require("../../src/helpers");
const { PARTS } = require("../support/golden");

test("correctly generates basic excel sheet xml", async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(rows);
  assert.equal(await getXmlFromXmlStream(xlsx.sheetXmlStream), PARTS["xl/worksheets/sheet1.xml"]);
});
