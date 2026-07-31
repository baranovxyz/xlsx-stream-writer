const replaceRegex = /\s+/g;
const replaceReSec = />\s+</g;

export const getRowStart = (row: number): string => `<row r="${row + 1}">`;
export const rowEnd = "</row>";

const $s = (styleId: number): string => (styleId === 0 ? "" : ' s="' + styleId + '"');

export const getStringCellXml = (index: number, cell: string, styleId = 0): string =>
  `<c r="${cell}" t="s"${$s(styleId)}><v>${index}</v></c>`;

export const getInlineStringCellXml = (s: string, cell: string, styleId = 0): string =>
  `<c r="${cell}" t="inlineStr"${$s(styleId)}><is><t>${s}</t></is></c>`;

export const getNumberCellXml = (
  value: number | string,
  cell: string,
  styleId = 0,
): string => `<c r="${cell}" t="n"${$s(styleId)}><v>${value}</v></c>`;

export const getBooleanCellXml = (value: boolean, cell: string, styleId = 0): string =>
  `<c r="${cell}" t="b"${$s(styleId)}><v>${value ? 1 : 0}</v></c>`;

// A blank cell carries no value element at all. Writing an empty one instead —
// `t="s"` with `<v></v>` — is a shared-string reference to nothing, which Excel
// treats as corrupt content.
export const getBlankCellXml = (cell: string, styleId = 0): string =>
  `<c r="${cell}"${$s(styleId)}/>`;

export const sheetHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <sheetData>`
  .replace(replaceRegex, " ")
  .replace(replaceReSec, "><")
  .trim();

export const sheetFooter = "</sheetData></worksheet>";

// `count` is how many cells reference the table; `uniqueCount` is how many
// distinct strings it holds.
export const getSharedStringsHeader = (
  uniqueCount: number,
  count: number = uniqueCount,
): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="${count}"
     uniqueCount="${uniqueCount}">`
    .replace(replaceRegex, " ")
    .replace(replaceReSec, "><")
    .trim();

export const getSharedStringXml = (s: string): string => `<si><t>${s}</t></si>`;
export const sharedStringsFooter = "</sst>";
