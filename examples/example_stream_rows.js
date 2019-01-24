const XlsxWriter = require("../src/xlsx-writer-browser");
const Readable = require("stream-browserify").Readable;
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

function wrapRowsInStream(rows) {
  const rs = Readable({ objectMode: true });
  let c = 0;
  rs._read = function() {
    if (c === rows.length) rs.push(null);
    else rs.push(rows[c]);
    c++;
  };
  return rs;
}
const streamOfRows = wrapRowsInStream(rows);

const xlsx = new XlsxWriter();
xlsx.addRows(streamOfRows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
