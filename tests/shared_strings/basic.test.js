const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../dist");
const { rows } = require("../helpers");
const { getXmlFromXmlStream } = require("../../dist/helpers");
const { PARTS } = require("../support/golden");

test("correctly generates shared strings xml for basic excel sheet", async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(rows);
  // The shared-strings table is only populated as the sheet stream is consumed,
  // so the sheet has to be drained first.
  await getXmlFromXmlStream(xlsx.sheetXmlStream);
  assert.equal(
    await getXmlFromXmlStream(xlsx.sharedStringsXmlStream),
    PARTS["xl/sharedStrings.xml"],
  );
});
