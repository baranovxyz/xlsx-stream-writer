const XlsxWriter = require("../src/xlsx-writer-browser");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const xlsx = new XlsxWriter();
xlsx.addRows(rows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
