const XlsxStreamWriter = require("../../index");
const { rows } = require("../helpers");
const { getXmlFromXmlStream } = require("../../src/helpers");

const sharedStringsXmlExpected = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="8"><si><t>Name</t></si><si><t>Location</t></si><si><t>Alpha</t></si><si><t>Adams</t></si><si><t>Bravo</t></si><si><t>Boston</t></si><si><t>Charlie</t></si><si><t>Chicago</t></si></sst>`;

test("correctly generates shared strings xml for basic excel sheet", async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(rows);
  await getXmlFromXmlStream(xlsx.sheetXmlStream);
  return expect(getXmlFromXmlStream(xlsx.sharedStringsXmlStream)).resolves.toBe(
    sharedStringsXmlExpected,
  );
});
