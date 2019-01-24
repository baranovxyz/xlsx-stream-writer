const XlsxStreamWriter = require("../src/xlsx-stream-writer");
const Readable = require("stream-browserify").Readable;
const fs = require("fs");

const { crc32 } = require("crc");
const NUM_TRIES = 100000;
console.time("crc");
const objCRC1 = {};
Array.from({ length: NUM_TRIES }, (_, i) => ({
  length: String(i % 135).length,
  hash: crc32(String(i % 135)),
})).map(v => {
  if (typeof objCRC1[v.hash] === "undefined")
    objCRC1[v.hash] = { count: 1, length: v.length };
  else objCRC1[v.hash].count++;
});
console.timeEnd("crc");
console.log(objCRC1.slice(0,100));

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

const xlsx = new XlsxStreamWriter();
xlsx.addRows(streamOfRows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
