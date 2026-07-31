const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../dist");
const { rows } = require("../helpers");
const { readZip } = require("../support/unzip");

async function stylesXmlFor(options) {
  const xlsx = new XlsxStreamWriter(options);
  xlsx.addRows(rows);
  return readZip(await xlsx.getFile()).text("xl/styles.xml");
}

test("two writers with the same styles produce identical styles.xml", async () => {
  const styles = [{ fill: "FFFF0000" }];
  const first = await stylesXmlFor({ styles });
  const second = await stylesXmlFor({ styles });

  // getStyles used to append into module-level arrays, so the second workbook
  // inherited the first one's fills and every style id shifted.
  assert.equal(second, first);
  assert.equal((second.match(/<patternFill patternType="solid">/g) || []).length, 1);
});

test("a styled writer does not leak its options into the next writer", () => {
  const first = new XlsxStreamWriter({ inlineStrings: true, styles: [{ fill: "FFFF0000" }] });
  const second = new XlsxStreamWriter();

  assert.equal(first.options.inlineStrings, true);
  assert.equal(second.options.inlineStrings, false);
  assert.deepEqual(second.options.styles, []);
});

test("writers built from different style sets stay independent", async () => {
  const red = await stylesXmlFor({ styles: [{ fill: "FFFF0000" }] });
  const green = await stylesXmlFor({ styles: [{ fill: "FF00FF00" }] });

  assert.match(red, /<fgColor rgb="FFFF0000"\/>/);
  assert.doesNotMatch(red, /FF00FF00/);
  assert.match(green, /<fgColor rgb="FF00FF00"\/>/);
  assert.doesNotMatch(green, /FFFF0000/);
  assert.match(green, /<fills count="3">/);
});

test("rejects options that would fail later, at construction time", () => {
  assert.throws(() => new XlsxStreamWriter({ styles: "red" }), /options\.styles must be an array/);
  assert.throws(() => new XlsxStreamWriter({ styleIdFunc: 1 }), /options\.styleIdFunc must be a function/);
});
