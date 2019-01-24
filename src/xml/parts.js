const replaceRegex = /\s+/g;

const getRowStart = row => `<row r="${row + 1}">`;
const rowEnd = "</row>";
const getStringCellXml = (index, cell) => `<c r="${cell}" t="s"><v>${index}</v></c>`;
const getNumberCellXml = (value, cell) => `<c r="${cell}" t="n"><v>${value}</v></c>`;

const sheetHeader =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="x14ac"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <sheetViews>
        <sheetView workbookViewId="0"/>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
    <sheetData>
  `.replace(replaceRegex, " ");
// const getSheetHeader = dimensions =>
//   `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//   <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
//              xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
//              mc:Ignorable="x14ac xr xr2 xr3"
//              xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
//     <dimension ref="${dimensions}"/>
//     <sheetViews>
//         <sheetView workbookViewId="0"/>
//     </sheetViews>
//     <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
//     <sheetData>
//   `.replace(replaceRegex, " ");

const sheetFooter = "</sheetData></worksheet>";

const getSharedStringsHeader = count =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="${count}"
     uniqueCount="${count}">
  `.replace(replaceRegex, " ");

const getSharedStringXml = s => `<si><t>${s}</t></si>`.replace(replaceRegex, " ");
const sharedStringsFooter = "</sst>";

module.exports = {
  getSharedStringsHeader,
  getSharedStringXml,
  sharedStringsFooter,
  sheetHeader,
  getRowStart,
  rowEnd,
  getStringCellXml,
  getNumberCellXml,
  sheetFooter,
};
