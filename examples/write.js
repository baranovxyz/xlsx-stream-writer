const XlsxWriter = require("../src/xlsx-writer-browser").XlsxWriter;
const getRowXml = require("../src/xlsx-writer-browser").getRowXml;
const fs = require("fs");

const xmlParts = require("../src/xml-parts");

// const rows = [["Name", "Location"]];

const name = "BobbyBobbyBobbyBobbyBobbyBobby";
const location = "RussiaRussiaRussiaRussiaRussia";

const rows = Array.from({ length: 10 }, (_, i) => [
  name.slice(0, 5 + (i % 25)).padStart(100, "0"),
  location.slice(0, 5 + (i % 25)).padStart(100, "0"),
]);

var Readable = require("stream-browserify").Readable;
var rs = Readable({ objectMode: true });

let c = 0;
rs._read = function() {
  if (c === rows.length) rs.push(null);
  else rs.push(rows[c]);
  c++;
};

const xlsx = new XlsxWriter({ decodeStrings: true });
xlsx.addRowsStream(rs);
// rs.pipe(xlsx);

xlsx.getFile().then(buffer => {
  console.log(buffer);
  fs.writeFileSync("result.xlsx", buffer);
});
// xlsx.end();
//
// const write = async rows => {
//   ;
//   rows.map(row => xlsx.addRow(row));
//   xlsx.end();
//   return xlsx.getFile();
// };
//
// write(rows).then(buffer => {
//   console.log("done!");
//   fs.writeFileSync("result.xlsx", buffer);
// });
