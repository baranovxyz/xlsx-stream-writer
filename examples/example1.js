const XlsxWriter = require("../src/xlsx-writer-browser");
const fs = require("fs");

const name = "BobbyBobbyBobbyBobbyBobbyBobby";
const location = "RussiaRussiaRussiaRussiaRussia";

const rows = Array.from({ length: 1000 }, (_, i) => [
  name.slice(0, 5 + i % 25).padStart(100, "0"),
  location.slice(0, 5 + i % 25).padStart(100, "0"),
]);

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
