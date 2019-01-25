const XlsxStreamWriter = require("../src/xlsx-stream-writer");
const { wrapRowsInStream, getXmlFromXmlStream } = require("../src/helpers");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Иван", "Москва"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const streamOfRows = wrapRowsInStream(rows);

const options = { inlineStrings: true };
const xlsx = new XlsxStreamWriter(options);
xlsx.addRows(streamOfRows);
// getXmlFromXmlStream(xlsx.sheetXmlStream).then(t => console.log(t) && fs.writeFileSync('t.txt', t));
xlsx.getFile().then(buffer => {
  fs.writeFileSync("result-inline-strings.xlsx", buffer);
});
