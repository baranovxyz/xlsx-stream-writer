const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../dist");
const { getStyles } = require("../../dist/styles");
const { getXmlFromXmlStream } = require("../../dist/helpers");
const { readZip } = require("../support/unzip");
const { assertWellFormedXml } = require("../support/xml");

// Keys that exist on Object.prototype. Looking one up in a plain object finds
// an inherited member rather than missing, and the member then gets written
// into the document where an index belonged.
const PROTOTYPE_KEYS = ["constructor", "toString", "valueOf", "hasOwnProperty", "__proto__"];

test("cell values that collide with Object.prototype are ordinary strings", async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([PROTOTYPE_KEYS, ["constructor"]]);
  const zip = readZip(await xlsx.getFile());

  const sheet = zip.text("xl/worksheets/sheet1.xml");
  const sharedStrings = zip.text("xl/sharedStrings.xml");
  assertWellFormedXml(assert, sheet, "sheet1.xml");
  assertWellFormedXml(assert, sharedStrings, "sharedStrings.xml");

  // Every reference must be a plain integer, never a stringified function.
  for (const [, index] of sheet.matchAll(/<v>([^<]*)<\/v>/g)) {
    assert.match(index, /^\d+$/, `shared-string reference ${JSON.stringify(index)} is not an index`);
  }
  for (const key of PROTOTYPE_KEYS) {
    assert.ok(sharedStrings.includes(`<t>${key}</t>`), `${key} should be in the string table`);
  }
  // The repeated "constructor" must dedupe to the entry already there.
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

test("reading the sheet stream then building a workbook is refused, not silently empty", async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([["Name"], ["Alpha"]]);
  const sheet = await getXmlFromXmlStream(xlsx.sheetXmlStream);
  assert.equal((sheet.match(/<row /g) || []).length, 2);

  // The rows are gone; producing a workbook now would lose every one of them.
  await assert.rejects(xlsx.getFile(), /sheetXmlStream has already been handed out/);
  assert.throws(() => xlsx.getStream(), /sheetXmlStream has already been handed out/);
});

test("a style id that is not a usable index is rejected at the cell", async () => {
  const build = async styleIdFunc => {
    const xlsx = new XlsxStreamWriter({ styles: [{ fill: "FFFF0000" }], styleIdFunc });
    xlsx.addRows([["a"]]);
    return xlsx.getFile();
  };

  // Interpolated into an attribute, this would have closed it and added another.
  await assert.rejects(build(() => '0" foo="bar'), /must return a non-negative integer/);
  await assert.rejects(build(() => 1.5), /must return a non-negative integer/);
  await assert.rejects(build(() => -1), /must return a non-negative integer/);
  await assert.rejects(build(() => NaN), /must return a non-negative integer/);
  // Style 2 does not exist: styles has one entry, so 0 and 1 are the valid ids.
  await assert.rejects(build(() => 2), /but only 2 styles are defined/);
});

test("writers running concurrently do not see each other's state", async () => {
  const build = (fill, value) => {
    const xlsx = new XlsxStreamWriter({ styles: [{ fill }], styleIdFunc: () => 1 });
    xlsx.addRows([[value]]);
    return xlsx.getFile();
  };

  // Interleaved rather than sequential: the state leaks this package used to
  // have would show up here even if the sequential tests passed.
  const cases = [
    ["FFFF0000", "red"],
    ["FF00FF00", "green"],
    ["FF0000FF", "blue"],
  ];
  const built = await Promise.all(cases.map(([fill, value]) => build(fill, value)));

  built.forEach((buffer, i) => {
    const [fill, value] = cases[i];
    const zip = readZip(buffer);
    const stylesXml = zip.text("xl/styles.xml");
    assert.match(stylesXml, new RegExp(`<fgColor rgb="${fill}"/>`));
    assert.match(stylesXml, /<fills count="3">/, "each writer should declare exactly one custom fill");
    assert.match(zip.text("xl/sharedStrings.xml"), new RegExp(`<t>${value}</t>`));
    for (const [, other] of cases.filter(c => c[0] !== fill)) {
      assert.doesNotMatch(zip.text("xl/sharedStrings.xml"), new RegExp(other));
    }
  });
});

test("style ids inside the declared range still work", async () => {
  const xlsx = new XlsxStreamWriter({
    styles: [{ fill: "FFFF0000" }, { format: "0.00" }],
    styleIdFunc: (value, columnId) => columnId,
  });
  xlsx.addRows([["a", "b", "c"]]);
  const sheet = readZip(await xlsx.getFile()).text("xl/worksheets/sheet1.xml");

  assert.match(sheet, /<c r="A1" t="s"><v>0<\/v><\/c>/); // style 0 is implicit
  assert.match(sheet, /<c r="B1" t="s" s="1">/);
  assert.match(sheet, /<c r="C1" t="s" s="2">/);
});
