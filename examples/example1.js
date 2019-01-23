const XlsxWriter = require("../src/xlsx-writer-browser");
const fs = require("fs");

const rows = [["Name", "Location"], ["Bob", "Sweden"], ["Alice", "France"]];

const write = async rows => {
  const xlsx = new XlsxWriter();
  rows.map(row => xlsx.addRow(row));
  xlsx.end();
  return xlsx.getFile();
};

write(rows).then(buffer => {
  console.log("done!");
  fs.writeFileSync("result.xlsx", buffer);
});
