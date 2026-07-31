// Stream a large workbook straight to disk, without ever holding it in memory.
const XlsxStreamWriter = require("../dist");
const { Readable } = require("node:stream");
const { pipeline } = require("node:stream/promises");
const fs = require("node:fs");

const ROW_COUNT = 500000;

function* generateRows() {
  yield ["Id", "Name", "Amount", "Active"];
  for (let i = 1; i < ROW_COUNT; i++) {
    yield [i, `Customer ${i}`, i * 1.5, i % 2 === 0];
  }
}

const xlsx = new XlsxStreamWriter();
xlsx.addRows(generateRows());

pipeline(Readable.fromWeb(xlsx.getStream()), fs.createWriteStream("large.xlsx"))
  .then(() => console.log(`wrote large.xlsx with ${ROW_COUNT} rows`))
  .catch(error => {
    console.error(error);
    process.exit(1);
  });
