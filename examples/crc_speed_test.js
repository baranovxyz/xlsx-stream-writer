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
console.log(objCRC1.slice(0, 100));
