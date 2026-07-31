import { escapeXmlExtended } from "./helpers";

const replaceRegex = /\s+/g;
const replaceReSec = />\s+</g;

/** A cell format, referenced by index from `styleIdFunc`. */
export interface CellStyle {
  /** ARGB fill colour, for example `FFFF0000`. */
  fill?: string;
  /** Excel number format string, for example `0.00` or `dd.mm.yyyy`. */
  format?: string;
}

const header = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x14ac"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">`;

const bottom = "</styleSheet>";

const getFillXmlHeader = (numFills: number) => `<fills count="${numFills}">`;
const fillXmlDefault = [
  `<fill>
  <patternFill patternType="none"/>
  </fill>`,
  `<fill>
  <patternFill patternType="gray125"/>
  </fill>`,
];

const getFillXml = (fillColor: string) =>
  `<fill><patternFill patternType="solid"><fgColor rgb="${fillColor}"/><bgColor indexed="64"/></patternFill></fill>`;

const fillXmlBottom = "</fills>";

const fontsXml = `<fonts count="1" x14ac:knownFonts="1">
        <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>`;

const bordersXml = `
    <borders count="1">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
    </borders>`;

const cellStyleXfs = `<cellStyleXfs count="1">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>`;

const getCellXfXml = ({ numFmtId, fillId }: { numFmtId?: number; fillId?: number }) =>
  `<xf numFmtId="${numFmtId === undefined ? 0 : numFmtId}" fontId="0" fillId="${
    fillId === undefined ? 0 : fillId
  }" borderId="0" xfId="0"/>`;

const cellXfXmlDefault = [`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`];

function getCellXfsBlock(cellXfs: string[]) {
  return `<cellXfs count="${cellXfs.length}">${cellXfs.join("")}</cellXfs>`;
}

const restXml = `<cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2"
                 defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
             xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
        </ext>
    </extLst>`;

const compact = (xml: string) =>
  xml.replace(replaceRegex, " ").replace(replaceReSec, "><").trim();

/**
 * Build `xl/styles.xml`.
 *
 * The parts appear in the order the schema fixes: number formats, fonts, fills,
 * borders, cell style formats, then cell formats — which is what a cell's style
 * index actually points into.
 */
export function getStyles(styles?: readonly CellStyle[] | null): string {
  const NUM_FORMATS_START = 166;
  const numFormatsXml: string[] = [];
  const numFormatsIndex: Record<string, number> = {};
  // Copy the defaults. Aliasing them would append every writer's fills and
  // cell formats to the module-level arrays, so the second workbook built in a
  // process would inherit the first one's styles and its style ids would point
  // at the wrong entries.
  const fillsXml = [...fillXmlDefault];
  const fillsIndex: Record<string, number> = {};
  const cellXfsXml = [...cellXfXmlDefault];

  for (const style of styles ?? []) {
    const { fill, format } = style;
    if (format !== undefined && numFormatsIndex[format] === undefined) {
      const formatIndex = numFormatsXml.length + NUM_FORMATS_START;
      numFormatsIndex[format] = formatIndex;
      numFormatsXml.push(getFormatXml(escapeXmlExtended(format), formatIndex));
    }
    if (fill !== undefined && fillsIndex[fill] === undefined) {
      fillsIndex[fill] = fillsXml.length;
      fillsXml.push(getFillXml(escapeXmlExtended(fill)));
    }
    cellXfsXml.push(
      getCellXfXml({
        numFmtId: format === undefined ? undefined : numFormatsIndex[format],
        fillId: fill === undefined ? undefined : fillsIndex[fill],
      }),
    );
  }

  let xml = "";
  xml += header;
  xml += getNumFormatsXmlBlock(numFormatsXml);
  xml += fontsXml;
  xml += getFillXmlBlock(fillsXml);
  xml += bordersXml;
  xml += cellStyleXfs;
  xml += getCellXfsBlock(cellXfsXml);
  xml += restXml;
  xml += bottom;
  return compact(xml);
}

const getFormatXml = (format: string, id: number) =>
  `<numFmt numFmtId="${id}" formatCode="${format}"/>`;

function getNumFormatsXmlBlock(formats: string[]) {
  if (!formats.length) return "";
  return `<numFmts count="${formats.length}">${formats.join("")}</numFmts>`;
}

function getFillXmlBlock(fillsXml: string[]) {
  return getFillXmlHeader(fillsXml.length) + fillsXml.join("") + fillXmlBottom;
}
