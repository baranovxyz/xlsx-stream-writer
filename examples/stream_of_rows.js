// Rows from a generator, written with getFile().
const XlsxStreamWriter = require("../src/xlsx-stream-writer");
const fs = require("fs");

function* generateRows() {
  yield ["Name", "Location"];
  yield ["Alpha", "Adams"];
  yield ["Bravo", "Boston"];
  yield ["Charlie", "Chicago"];
}

const xlsx = new XlsxStreamWriter();
xlsx.addRows(generateRows());

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
