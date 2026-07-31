const test = require("node:test");
const assert = require("node:assert/strict");
const { Readable } = require("node:stream");

const XlsxStreamWriter = require("../../index");
const { rows } = require("../helpers");
const { getXmlFromXmlStream } = require("../../src/helpers");
const { PARTS } = require("../support/golden");

const expected = PARTS["xl/worksheets/sheet1.xml"];

async function sheetXmlFrom(input) {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(input);
  return getXmlFromXmlStream(xlsx.sheetXmlStream);
}

test("accepts a native node:stream Readable", async () => {
  // This used to throw: the instanceof check tested stream-browserify's class,
  // so the stream type every Node caller actually has was rejected.
  assert.equal(await sheetXmlFrom(Readable.from(rows, { objectMode: true })), expected);
});

test("accepts an old-style Node readable with no async iterator", async () => {
  // Streams from older libraries predate Symbol.asyncIterator; they still have
  // to work, so they are bridged through their data/end/error events.
  const stream = new Readable({ objectMode: true, read() {} });
  delete stream[Symbol.asyncIterator];
  queueMicrotask(() => {
    for (const row of rows) stream.push(row);
    stream.push(null);
  });
  assert.equal(await sheetXmlFrom(stream), expected);
});

test("accepts a web ReadableStream", async () => {
  const stream = new ReadableStream({
    start(controller) {
      for (const row of rows) controller.enqueue(row);
      controller.close();
    },
  });
  assert.equal(await sheetXmlFrom(stream), expected);
});

test("accepts an async generator", async () => {
  async function* generate() {
    for (const row of rows) yield row;
  }
  assert.equal(await sheetXmlFrom(generate()), expected);
});

test("accepts a sync iterable", async () => {
  assert.equal(await sheetXmlFrom(new Set(rows)), expected);
});

test("rejects values that are not rows", () => {
  const xlsx = new XlsxStreamWriter();
  assert.throws(() => xlsx.addRows("Name,Location"), { name: "TypeError" });
  assert.throws(() => new XlsxStreamWriter().addRows(42), { name: "TypeError" });
  assert.throws(() => new XlsxStreamWriter().addRows(null), { name: "TypeError" });
});

test("an empty workbook is still a valid sheet", async () => {
  const xml = await sheetXmlFrom([]);
  assert.match(xml, /^<\?xml /);
  assert.match(xml, /<sheetData><\/sheetData><\/worksheet>$/);
});
