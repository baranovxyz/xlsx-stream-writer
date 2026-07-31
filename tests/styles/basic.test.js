const test = require("node:test");
const assert = require("node:assert/strict");

const XlsxStreamWriter = require("../../dist");
const { rows } = require("../helpers");
const { getXmlFromXmlStream } = require("../../dist/helpers");
const { readZip } = require("../support/unzip");
const { assertWellFormedXml } = require("../support/xml");

const styles = [{ fill: "FFFF0000" }, { format: "0.00" }, { format: "dd.mm.yy" }];

// Custom number formats start at 166; ids below that are reserved for the
// built-in formats Excel already knows.
const expectedStylesXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><numFmts count="2"><numFmt numFmtId="166" formatCode="0.00"/><numFmt numFmtId="167" formatCode="dd.mm.yy"/></numFmts><fonts count="1" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>';

test("declared styles become numFmts, fills and cellXfs in styles.xml", async () => {
  const xlsx = new XlsxStreamWriter({ styles });
  xlsx.addRows(rows);
  const zip = readZip(await xlsx.getFile());
  const stylesXml = zip.text("xl/styles.xml");

  assertWellFormedXml(assert, stylesXml, "xl/styles.xml");
  assert.equal(stylesXml, expectedStylesXml);
});

test("styleIdFunc selects the cellXfs entry each cell references", async () => {
  const xlsx = new XlsxStreamWriter({
    styles,
    styleIdFunc: (value, columnId, rowId) => (rowId === 0 ? 1 : 0),
  });
  xlsx.addRows(rows);
  const sheetXml = await getXmlFromXmlStream(xlsx.sheetXmlStream);

  // Style 0 is the implicit default and is written as an absent attribute.
  assert.match(sheetXml, /<c r="A1" t="s" s="1">/);
  assert.match(sheetXml, /<c r="A2" t="s"><v>/);
});
