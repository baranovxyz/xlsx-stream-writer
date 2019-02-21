const Readable = require("stream-browserify").Readable;
const Writable = require("stream-browserify").Writable;

function getCellAddress(rowIndex, colIndex) {
  let colAddress = "";
  let input = (colIndex - 1).toString(26);
  while (input.length) {
    const a = input.charCodeAt(input.length - 1);
    colAddress =
      String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colAddress;
    input =
      input.length > 1
        ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26)
        : "";
  }
  return colAddress + rowIndex;
}

function getXmlFromXmlStream(xmlStream) {
  return new Promise((resolve, reject) => {
    const ws = Writable();
    let xml = "";
    ws._write = function(chunk, enc, next) {
      xml += chunk.toString();
      next();
    };
    xmlStream.pipe(ws);
    ws.on("finish", () => resolve(xml));
    ws.on("error", reject);
  });
}

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

function escapeXml(str = "") {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeXmlExtended(str = "") {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

module.exports = {
  getCellAddress,
  wrapRowsInStream,
  getXmlFromXmlStream,
  escapeXml,
  escapeXmlExtended,
};
