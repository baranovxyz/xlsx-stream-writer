const Writable = require("stream-browserify").Writable;

function getXml(xmlStream) {
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

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

module.exports = { getXml, rows };
