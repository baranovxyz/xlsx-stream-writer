const test = require("node:test");
const assert = require("node:assert/strict");
const { Readable } = require("node:stream");

const XlsxStreamWriter = require("../../index");
const { rows } = require("../helpers");

function failingStream(message) {
  return new Readable({
    objectMode: true,
    read() {
      this.destroy(new Error(message));
    },
  });
}

// Each of these used to leave getFile() pending forever, because .pipe() does
// not forward errors: the destination stayed open and JSZip waited on a stream
// that would never end. The timeout is the assertion — a regression hangs.
test("a failing source stream rejects getFile() instead of hanging", { timeout: 5000 }, async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(failingStream("source blew up"));
  await assert.rejects(xlsx.getFile(), /source blew up/);
});

test("a failing source rejects even in inlineStrings mode", { timeout: 5000 }, async () => {
  const xlsx = new XlsxStreamWriter({ inlineStrings: true });
  xlsx.addRows(failingStream("inline source blew up"));
  await assert.rejects(xlsx.getFile(), /inline source blew up/);
});

test("an error raised before getFile() is still reported", { timeout: 5000 }, async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(failingStream("early failure"));
  // Give the stream time to fail while nothing is listening for the promise.
  await new Promise(resolve => setTimeout(resolve, 50));
  await assert.rejects(xlsx.getFile(), /early failure/);
});

test("a throwing styleIdFunc rejects getFile()", { timeout: 5000 }, async () => {
  const xlsx = new XlsxStreamWriter({
    styleIdFunc: () => {
      throw new Error("style lookup failed");
    },
  });
  xlsx.addRows(rows);
  await assert.rejects(xlsx.getFile(), /style lookup failed/);
});

test("a row that is not an array rejects with the row number", { timeout: 5000 }, async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([["ok"], "not a row"]);
  await assert.rejects(xlsx.getFile(), /Row 2 is not an array of cell values/);
});

test("an async source that rejects rejects getFile()", { timeout: 5000 }, async () => {
  async function* generate() {
    yield ["ok"];
    throw new Error("generator blew up");
  }
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(generate());
  await assert.rejects(xlsx.getFile(), /generator blew up/);
});

// The once-only lifecycle guards are a 1.x change: they turn misuse into an
// error rather than fixing a corrupt file, so they are not backported here.
