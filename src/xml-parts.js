const getRowStart = row => `<row r="${row + 1}">`;
const rowEnd = "</row>";
const getStringCellXml = (index, cell) => `<c r="${cell}" t="s"><v>${index}</v></c>`;
const getNumberCellXml = (value, cell) => `<c r="${cell}" t="n"><v>${value}</v></c>`;

const getSheetHeader = dimensions =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             mc:Ignorable="x14ac xr xr2 xr3"
             xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <dimension ref="${dimensions}"/>
    <sheetViews>
        <sheetView workbookViewId="0"/>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
    <sheetData>
  `.replace(/\n\s*/g, "");

const sheetFooter = "</sheetData></worksheet>".replace(/\n\s*/g, "");

const getSharedStringsHeader = count =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="${count}"
     uniqueCount="${count}">
  `.replace(/\n\s*/g, "");

const getSharedStringXml = s => `<si><t>${s}</t></si>`.replace(/\n\s*/g, "");
const sharedStringsFooter = "</sst>".replace(/\n\s*/g, "");

module.exports = {
  getSharedStringsHeader,
  getSharedStringXml,
  sharedStringsFooter,
  getSheetHeader,
  getRowStart,
  rowEnd,
  getStringCellXml,
  getNumberCellXml,
  sheetFooter,
};
