// Inline strings skip the shared-string table: a larger file, but nothing to
// accumulate in memory while the sheet is written.
const XlsxStreamWriter = require("../dist");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Иван", "Москва"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const xlsx = new XlsxStreamWriter({ inlineStrings: true });
xlsx.addRows(rows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result-inline-strings.xlsx", buffer);
});
