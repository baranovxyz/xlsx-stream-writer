const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../dist");
const { getXmlFromXmlStream, escapeXml, escapeXmlExtended } = require("../../dist/helpers");
const { assertWellFormedXml, findIllegalChar } = require("../support/xml");
const { readZip } = require("../support/unzip");

async function sheetXmlFor(row, options) {
  const xlsx = new XlsxStreamWriter(options);
  xlsx.addRows([row]);
  return getXmlFromXmlStream(xlsx.sheetXmlStream);
}

test("XML metacharacters are escaped", () => {
  assert.equal(escapeXml("a & b < c > d"), "a &amp; b &lt; c &gt; d");
  assert.equal(escapeXmlExtended(`"quoted" 'single' & <tag>`), "&quot;quoted&quot; &apos;single&apos; &amp; &lt;tag&gt;");
});

test("control characters are removed, legal whitespace is kept", () => {
  assert.equal(escapeXml("a\u0000b\u001Fc"), "abc");
  assert.equal(escapeXml("a\tb\nc\rd"), "a\tb\nc\rd");
  assert.equal(escapeXml("a\uFFFEb\uFFFFc"), "abc");
});

test("lone surrogates are replaced, valid pairs survive", () => {
  assert.equal(escapeXml("a\uD800b"), "a\uFFFDb");
  assert.equal(escapeXml("a\uDC00b"), "a\uFFFDb");
  assert.equal(escapeXml("emoji \u{1F600} ok"), "emoji \u{1F600} ok");
});

test("plain text is returned unchanged", () => {
  assert.equal(escapeXml("Location"), "Location");
  assert.equal(escapeXml(""), "");
  assert.equal(escapeXml(), "");
});

test("a control character in a cell does not corrupt the inline-string sheet", async () => {
  const xml = await sheetXmlFor(["a\u0000b", "c\u001Fd", "e\uD800f"], { inlineStrings: true });
  assert.equal(findIllegalChar(xml), null);
  assertWellFormedXml(assert, xml, "sheet1.xml");
  assert.match(xml, /<t>ab<\/t>/);
  assert.match(xml, /<t>cd<\/t>/);
});

test("a control character in a cell does not corrupt the shared strings part", async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([["a\u0000b", "<script>", "R&D"]]);
  const zip = readZip(await xlsx.getFile());

  const sharedStrings = zip.text("xl/sharedStrings.xml");
  assertWellFormedXml(assert, sharedStrings, "xl/sharedStrings.xml");
  assert.match(sharedStrings, /<si><t>ab<\/t><\/si>/);
  assert.match(sharedStrings, /<si><t>&lt;script&gt;<\/t><\/si>/);
  assert.match(sharedStrings, /<si><t>R&amp;D<\/t><\/si>/);
});

test("style fills and formats cannot break out of their attributes", async () => {
  const xlsx = new XlsxStreamWriter({
    styles: [{ fill: 'FF0000"/><evil x="', format: '0.00"/><evil x="' }],
  });
  xlsx.addRows([["a"]]);
  const stylesXml = readZip(await xlsx.getFile()).text("xl/styles.xml");

  assertWellFormedXml(assert, stylesXml, "xl/styles.xml");
  assert.doesNotMatch(stylesXml, /<evil/);
});

test("shared strings count references, uniqueCount distinct values", async () => {
  const xlsx = new XlsxStreamWriter();
  xlsx.addRows([["a", "a", "b"], ["a", "c", 1]]);
  const sharedStrings = readZip(await xlsx.getFile()).text("xl/sharedStrings.xml");

  // Five string cells across the two rows; three distinct values.
  assert.match(sharedStrings, /count="5" uniqueCount="3"/);
});
