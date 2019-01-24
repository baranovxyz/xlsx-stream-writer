const XlsxStreamWriter = require("../src/xlsx-stream-writer");
const { wrapRowsInStream } = require("../src/helpers");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const streamOfRows = wrapRowsInStream(rows);

const options = { inlineStrings: true };
const xlsx = new XlsxStreamWriter(options);
xlsx.addRows(streamOfRows);
xlsx.getFile().then(buffer => {
  fs.writeFileSync("result-inline-strings.xlsx", buffer);
});
