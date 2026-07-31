const XlsxStreamWriter = require("../dist");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Иван", "Москва"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const xlsx = new XlsxStreamWriter();
xlsx.addRows(rows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
