const XlsxWriter = require("../src/xlsx-writer-browser").XlsxWriter;
const getRowXml = require("../src/xlsx-writer-browser").getRowXml;
const fs = require("fs");

// const rows = [["Name", "Location"]];

const name = "BobbyBobbyBobbyBobbyBobbyBobby";
const location = "RussiaRussiaRussiaRussiaRussia";

const rows = Array.from({ length: 10 }, (_, i) => [
  name.slice(0, 5 + (i % 25)).padStart(100, "0"),
  location.slice(0, 5 + (i % 25)).padStart(100, "0"),
]);

var Readable = require("stream-browserify").Readable;
var rs = Readable();

let c = 0;
rs._read = function() {
  if (c === rows.length) return rs.push(null);
  rs.push(getRowXml.bind(xlsx)(rows[c], c));
  c++;
};

const xlsx = new XlsxWriter({decodeStrings: true});

rs.pipe(xlsx);
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
