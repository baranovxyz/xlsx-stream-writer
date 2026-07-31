const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../index");
const { getStyles } = require("../../src/styles");
const { readZip } = require("../support/unzip");
const { assertWellFormedXml } = require("../support/xml");

// Keys that exist on Object.prototype. Looking one up in a plain object finds an
// inherited member rather than missing, and the member then gets written into
// the document where an index belonged. This corrupted every workbook that
// contained one of these words in a cell, in every release up to 0.2.6.
const PROTOTYPE_KEYS = ["constructor", "toString", "valueOf", "hasOwnProperty", "__proto__"];

test("cell values that collide with Object.prototype are ordinary strings", async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([PROTOTYPE_KEYS, ["constructor"]]);
  const zip = readZip(await xlsx.getFile());

  const sheet = zip.text("xl/worksheets/sheet1.xml");
  const sharedStrings = zip.text("xl/sharedStrings.xml");
  assertWellFormedXml(assert, sheet, "sheet1.xml");
  assertWellFormedXml(assert, sharedStrings, "sharedStrings.xml");

  for (const [, index] of sheet.matchAll(/<v>([^<]*)<\/v>/g)) {
    assert.match(index, /^\d+$/, `shared-string reference ${JSON.stringify(index)} is not an index`);
  }
  for (const key of PROTOTYPE_KEYS) {
    assert.ok(sharedStrings.includes(`<t>${key}</t>`), `${key} should be in the string table`);
  }
  assert.match(sharedStrings, /count="6" uniqueCount="5"/);
});

test("style formats and fills that collide with Object.prototype are literal", () => {
  const stylesXml = getStyles([
    { format: "constructor" },
    { fill: "toString" },
    { format: "__proto__" },
  ]);
  assertWellFormedXml(assert, stylesXml, "styles.xml");

  const cellXfs = stylesXml.match(/<cellXfs[^]*?<\/cellXfs>/)[0];
  for (const [, value] of cellXfs.matchAll(/numFmtId="([^"]*)"/g)) {
    assert.match(value, /^\d+$/, `numFmtId ${JSON.stringify(value)} is not numeric`);
  }
  for (const [, value] of cellXfs.matchAll(/fillId="([^"]*)"/g)) {
    assert.match(value, /^\d+$/, `fillId ${JSON.stringify(value)} is not numeric`);
  }
});
