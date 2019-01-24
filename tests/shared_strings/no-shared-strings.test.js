const XlsxStreamWriter = require("../../index");
const { rows } = require("../helpers");
const { getXmlFromXmlStream } = require("../../src/helpers");

test("shared strings array is empty if inlineStrings: true option is set", async () => {
  const xlsx = new XlsxStreamWriter({inlineStrings: true});
  xlsx.addRows(rows);
  await getXmlFromXmlStream(xlsx.sheetXmlStream);
  return expect(xlsx.sharedStringsArr.length).toEqual(0);
});
