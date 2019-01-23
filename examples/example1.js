const XlsxWriter = require("../src/xlsx-writer-browser");
const fs = require("fs");

const rows = [["Name", "Location"], ["Bob", "Sweden"], ["Alice", "France"]];

const write = async rows => {
  const xlsx = new XlsxWriter(2, 3);
  rows.map(row => xlsx.addRow(row));
  return xlsx.getFile();
};

write(rows).then(blob => {
  console.log("done!");
  fs.writeFileSync("result.xlsx", blob);
});
