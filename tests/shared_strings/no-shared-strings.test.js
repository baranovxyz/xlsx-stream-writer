const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../index");
const { rows } = require("../helpers");
const { getXmlFromXmlStream } = require("../../src/helpers");

test("shared strings array is empty if inlineStrings: true option is set", async () => {
  const xlsx = new XlsxStreamWriter({ inlineStrings: true });
  xlsx.addRows(rows);
  await getXmlFromXmlStream(xlsx.sheetXmlStream);
  assert.equal(xlsx.sharedStringsArr.length, 0);
});

test("inline strings are written into the sheet instead of referenced", async () => {
  const xlsx = new XlsxStreamWriter({ inlineStrings: true });
  xlsx.addRows(rows);
  const sheetXml = await getXmlFromXmlStream(xlsx.sheetXmlStream);
  assert.match(sheetXml, /<c r="A1" t="inlineStr"><is><t>Name<\/t><\/is><\/c>/);
  assert.doesNotMatch(sheetXml, /t="s"/);
});
