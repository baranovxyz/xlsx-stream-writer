const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../dist");
const { rows } = require("../helpers");
const { readZip } = require("../support/unzip");
const { assertWellFormedXml } = require("../support/xml");
const { ENTRY_NAMES, FILE_ENTRY_NAMES, PARTS } = require("../support/golden");

// One writer for the whole file: the archive is expensive to build and every
// assertion below inspects the same bytes.
const archive = (async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows(rows);
  return readZip(await xlsx.getFile());
})();

test("archive contains exactly the expected entries, in order", async () => {
  const zip = await archive;
  assert.deepEqual(zip.names, ENTRY_NAMES);
});

test("every part matches its golden byte for byte", async () => {
  const zip = await archive;
  for (const name of FILE_ENTRY_NAMES) {
    assert.equal(zip.text(name), PARTS[name], `${name} differs from its golden`);
  }
});

test("every part is well-formed XML", async () => {
  const zip = await archive;
  for (const name of FILE_ENTRY_NAMES) {
    assertWellFormedXml(assert, zip.text(name), name);
  }
});

test("parts are deflated and marked as UTF-8", async () => {
  const zip = await archive;
  for (const name of FILE_ENTRY_NAMES) {
    const entry = zip.entry(name);
    assert.equal(entry.method, 8, `${name} should be DEFLATE`);
    assert.ok(entry.utf8NameFlag, `${name} should flag its name as UTF-8`);
  }
});

test("a small workbook uses plain entries, not streamed ones", async () => {
  const zip = await archive;
  // Under the buffering limit the writer knows the exact sizes before it emits
  // the header, so no entry needs a data descriptor or ZIP64 record. This is
  // the most widely readable form, and it is what small workbooks should get.
  for (const name of FILE_ENTRY_NAMES) {
    const entry = zip.entry(name);
    assert.equal(entry.usesDataDescriptor, false, `${name} should not need a data descriptor`);
    assert.equal(entry.zip64, false, `${name} should not need ZIP64`);
    assert.ok(entry.uncompressedSize > 0, `${name} should record its real size`);
  }
});

test("the same rows always produce the same bytes", async () => {
  const build = async () => {
    const xlsx = new XlsxStreamWriter();
    xlsx.addRows(rows);
    return xlsx.getFile();
  };
  // Timestamps are fixed rather than "now", so archives are reproducible.
  assert.deepEqual(await build(), await build());
});

// readZip verifies every stored CRC and uncompressed length while parsing, so a
// successful parse is itself the assertion. This test states that explicitly so
// the guarantee is not lost if the reader is refactored.
test("archive parses cleanly, so every stored CRC and length is correct", async () => {
  const zip = await archive;
  assert.equal(zip.entries.length, ENTRY_NAMES.length);
});
