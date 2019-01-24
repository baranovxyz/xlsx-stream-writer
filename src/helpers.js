const xmlParts = require("./xml-parts");

function getCellAddress(rowIndex, colIndex) {
  let colAddress = "";
  let input = (colIndex - 1).toString(26);
  while (input.length) {
    const a = input.charCodeAt(input.length - 1);
    colAddress =
      String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colAddress;
    input =
      input.length > 1
        ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26)
        : "";
  }
  return colAddress + rowIndex;
}

function getRowXml(row, rowIndex) {
  let rowBuffer = xmlParts.getRowStart(rowIndex);
  row.forEach((cellValue, colIndex) => {
    const cellAddress = getCellAddress(rowIndex + 1, colIndex + 1);
    rowBuffer += getCellXml.bind(this)(cellValue, cellAddress);
  });
  rowBuffer += xmlParts.rowEnd;
  return rowBuffer;
}

function getCellXml(value, address) {
  let cellXml;
  if (Number.isNaN(value) || value === null || typeof value === "undefined")
    cellXml = xmlParts.getStringCellXml("", address);
  else if (typeof value === "number")
    cellXml = xmlParts.getNumberCellXml(value, address);
  else
    cellXml = xmlParts.getStringCellXml(
      this._lookupString(String(value)),
      address,
    );
  return cellXml;
}

function getRange(cntRows, cntColumns) {
  return "A1:" + getCellAddress(cntRows, cntColumns);
}

module.exports = {
  getCellAddress,
  getRowXml,
  getCellXml,
};
