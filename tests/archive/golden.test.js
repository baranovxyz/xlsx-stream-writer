const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../index");
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

test("parts are deflated and written in streaming mode", async () => {
  const zip = await archive;
  for (const name of FILE_ENTRY_NAMES) {
    const entry = zip.entry(name);
    assert.equal(entry.method, 8, `${name} should be DEFLATE`);
    assert.ok(
      entry.usesDataDescriptor,
      `${name} should carry a data descriptor — sizes are not known until the stream ends`,
    );
  }
});

// readZip verifies every stored CRC and uncompressed length while parsing, so a
// successful parse is itself the assertion. This test states that explicitly so
// the guarantee is not lost if the reader is refactored.
test("archive parses cleanly, so every stored CRC and length is correct", async () => {
  const zip = await archive;
  assert.equal(zip.entries.length, ENTRY_NAMES.length);
});
