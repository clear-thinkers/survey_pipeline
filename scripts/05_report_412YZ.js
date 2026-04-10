"use strict";
/**
 * 05_report_412YZ.js
 * Generate report/412YZ/report_412YZ_v2.docx from output/412YZ/analysis_412YZ.xlsx.
 *
 * Usage:
 *   node scripts/05_report_412YZ.js
 */

const path = require("path");
const fs   = require("fs");
// ---------------------------------------------------------------------------
// NOTE: This file is a complete replacement — all sections are implemented.
// ---------------------------------------------------------------------------

// Global node_modules fallback for when script is run without a local package.json
const GLOBAL_NM = "C:\\Users\\alexi\\AppData\\Roaming\\npm\\node_modules";
function requireGlobal(name) {
  try { return require(name); } catch (_) {}
  return require(path.join(GLOBAL_NM, name));
}

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  ImageRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  VerticalAlign,
  LevelFormat,
  PageNumber,
  PageBreak,
  Header,
  Footer,
  convertInchesToTwip,
  UnderlineType,
  TableLayoutType,
} = requireGlobal("docx");

const XLSX = requireGlobal("xlsx");
const { parse: parseCsv } = require(path.join(GLOBAL_NM, "csv-parse", "dist", "cjs", "sync.cjs"));

// sharp — local install preferred, global fallback
let sharp;
try { sharp = require(path.join(__dirname, "..", "node_modules", "sharp")); }
catch (_) { try { sharp = requireGlobal("sharp"); } catch (__) { sharp = null; } }

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const BASE_DIR      = path.join(__dirname, "..");
const ANALYSIS_PATH = path.join(BASE_DIR, "output", "412YZ", "analysis_412YZ.xlsx");
const CSV_PATH      = path.join(BASE_DIR, "output", "412YZ", "survey_data_412YZ.csv");
const OUT_PATH      = path.join(BASE_DIR, "report", "412YZ", "report_412YZ_v2.docx");

const SURVEY_MONTH  = "March 2026";
let   N_RESPONDENTS = 0; // set dynamically from CSV row count in main()
const HDR_FILL      = "DCE6F1";

// Prior-year Q1 coach satisfaction benchmarks (% top-2 box, from example report)
const Q1_BENCHMARKS = {
  col:  ["Sep-19", "Mar-22", "Feb-23", "Feb-24", "Feb-25"],
  n:    ["n=154",  "n=103",  "n=128",  "n=142",  "n=167"],
  "Is trustworthy":                      ["94%", "91%", "95%", "95%", "90%"],
  "Is reliable":                         ["92%", "90%", "88%", "93%", "89%"],
  "Values my opinions about my life":    ["94%", "91%", "91%", "93%", "91%"],
  "Is available to me when I need them": ["85%", "88%", "88%", "91%", "88%"],
  "Makes me feel heard and understood":  ["92%", "90%", "92%", "91%", "89%"],
};

// Loaded in main() from table_widths_412YZ.json
let TABLE_WIDTHS = {};

// ---------------------------------------------------------------------------
// Chart helper
// ---------------------------------------------------------------------------

const CHARTS_DIR = path.join(BASE_DIR, "output", "412YZ", "charts");
const BODY_PARAGRAPH_SPACING = { before: 240, after: 240, line: 240 };
const SPACER_PARAGRAPHS = new WeakSet();

/**
 * embedChart(chartFilename, widthInches = 6)
 * Reads the PNG from output/412YZ/charts/, returns a centered Paragraph
 * containing an ImageRun. docx v9 transformation.width/height are in screen
 * pixels (96 DPI). Height is computed proportionally using sharp.
 */
async function embedChart(chartFilename, widthInches = 6) {
  const chartPath = path.join(CHARTS_DIR, chartFilename);
  if (!fs.existsSync(chartPath)) {
    console.warn(`  [warn] Chart not found: ${chartPath}`);
    return new Paragraph({ text: "" });
  }
  const data = fs.readFileSync(chartPath);

  // docx v9 multiplies by 9525 (EMU/px at 96 DPI) internally, so pass pixels
  let widthPx  = Math.round(widthInches * 96);
  let heightPx = Math.round(widthPx * 0.5625); // fallback 16:9

  if (sharp) {
    try {
      const meta  = await sharp(chartPath).metadata();
      const ratio = meta.height / meta.width;
      heightPx    = Math.round(widthPx * ratio);
    } catch (e) {
      console.warn(`  [warn] sharp metadata failed for ${chartFilename}: ${e.message}`);
    }
  }

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [
      new ImageRun({
        data,
        type: "png",
        transformation: { width: widthPx, height: heightPx },
      }),
    ],
  });
}


// ---------------------------------------------------------------------------
// Numbering config for bullet lists
// ---------------------------------------------------------------------------

const BULLET_NUMBERING = {
  config: [
    {
      reference: "bullets",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: 720, hanging: 360 },
            },
          },
        },
      ],
    },
  ],
};

// ---------------------------------------------------------------------------
// Document styles
// ---------------------------------------------------------------------------

const DOC_STYLES = {
  default: {
    document: {
      run: {
        font: "Calibri",
        size: 22, // half-points → 11pt
      },
    },
  },
  paragraphStyles: [
    {
      id: "Heading1",
      name: "Heading 1",
      basedOn: "Normal",
      next: "Normal",
      quickFormat: true,
      run: {
        bold: true,
        size: 26, // 13pt
        color: "000000",
        font: "Calibri",
      },
      paragraph: {
        spacing: { before: 240, after: 240 },
        outlineLevel: 0,
      },
    },
    {
      id: "Heading2",
      name: "Heading 2",
      basedOn: "Normal",
      next: "Normal",
      quickFormat: true,
      run: {
        bold: true,
        size: 22, // 11pt
        color: "000000",
        font: "Calibri",
      },
      paragraph: {
        spacing: { before: 180, after: 240 },
        outlineLevel: 1,
      },
    },
    {
      id: "Caption",
      name: "Caption",
      basedOn: "Normal",
      next: "Normal",
      run: {
        bold: true,
        size: 22, // 11pt
        font: "Calibri",
      },
      paragraph: {
        spacing: { before: 120, after: 240 },
      },
    },
  ],
};

// ---------------------------------------------------------------------------
// Helpers — document elements
// ---------------------------------------------------------------------------

const NO_BORDER = {
  top:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function makeStyledTableCell(text, {
  width,
  bold = false,
  italic = false,
  shading,
  columnSpan,
  align,
  indentLeft = 0,
  size = 22,
  borders,
} = {}) {
  return new TableCell({
    width,
    columnSpan,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    borders: borders || NO_BORDER,
    shading: shading
      ? { fill: shading, type: ShadingType.CLEAR, color: "auto" }
      : undefined,
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        alignment: align,
        indent: indentLeft ? { left: indentLeft } : undefined,
        children: [
          new TextRun({
            text: String(text ?? ""),
            bold,
            italic,
            font: "Calibri",
            size,
            underline: { type: UnderlineType.NONE },
          }),
        ],
      }),
    ],
  });
}

function makeCoachSatisfactionTable(dataRows, currN, captionText = "") {
  const fixedWidths = captionText && TABLE_WIDTHS[captionText]
    ? TABLE_WIDTHS[captionText]
    : [3900, 910, 910, 910, 910, 910, 910];
  const yearCols = [...Q1_BENCHMARKS.col, "Mar-26"];
  const sampleSizes = [...Q1_BENCHMARKS.n, currN];
  const headerFill = HDR_FILL;
  const subrowFill = "EEF3F8";

  const headerTop = new TableRow({
    children: [
      makeStyledTableCell("My Youth Coach...", {
        width: { size: fixedWidths[0], type: WidthType.DXA },
        bold: true,
        shading: headerFill,
      }),
      makeStyledTableCell("% Often or All the Time", {
        width: { size: fixedWidths.slice(1).reduce((sum, val) => sum + val, 0), type: WidthType.DXA },
        bold: true,
        shading: headerFill,
        columnSpan: yearCols.length,
        align: AlignmentType.CENTER,
      }),
    ],
  });

  const headerYears = new TableRow({
    children: [
      makeStyledTableCell("", {
        width: { size: fixedWidths[0], type: WidthType.DXA },
        shading: headerFill,
      }),
      ...yearCols.map((label, idx) => makeStyledTableCell(label, {
        width: { size: fixedWidths[idx + 1], type: WidthType.DXA },
        bold: true,
        shading: headerFill,
        align: AlignmentType.CENTER,
      })),
    ],
  });

  const sampleRow = new TableRow({
    children: [
      makeStyledTableCell("n", {
        width: { size: fixedWidths[0], type: WidthType.DXA },
        italic: true,
        shading: subrowFill,
      }),
      ...sampleSizes.map((label, idx) => makeStyledTableCell(label, {
        width: { size: fixedWidths[idx + 1], type: WidthType.DXA },
        italic: true,
        shading: subrowFill,
        align: AlignmentType.CENTER,
      })),
    ],
  });

  const itemRows = dataRows.map((row) => {
    const label = row.label;
    const values = [...Q1_BENCHMARKS[label], row.current];
    return new TableRow({
      children: [
        makeStyledTableCell(label, {
          width: { size: fixedWidths[0], type: WidthType.DXA },
        }),
        ...values.map((value, idx) => makeStyledTableCell(value, {
          width: { size: fixedWidths[idx + 1], type: WidthType.DXA },
          align: AlignmentType.CENTER,
        })),
      ],
    });
  });

  return new Table({
    width: { size: fixedWidths.reduce((sum, val) => sum + val, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: {
      top:     { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideH: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideV: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: [headerTop, headerYears, sampleRow, ...itemRows],
  });
}

const LIGHT_TABLE_BORDER = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const BANKING_HIGHLIGHT_FILL = "FBE4D5";

function makeCellBorders({ top = false, bottom = false, left = false, right = false } = {}) {
  return {
    top: top ? LIGHT_TABLE_BORDER : NO_BORDER.top,
    bottom: bottom ? LIGHT_TABLE_BORDER : NO_BORDER.bottom,
    left: left ? LIGHT_TABLE_BORDER : NO_BORDER.left,
    right: right ? LIGHT_TABLE_BORDER : NO_BORDER.right,
  };
}

function getBankingTableWidths(colCount) {
  if (colCount === 6) return [3510, 1170, 1260, 1170, 920, 920];
  if (colCount === 7) return [3150, 1050, 1050, 1050, 1050, 800, 800];
  const firstColWidth = 3200;
  const remainder = 8950 - firstColWidth;
  const otherWidth = Math.floor(remainder / Math.max(colCount - 1, 1));
  const widths = [firstColWidth, ...Array(Math.max(colCount - 1, 0)).fill(otherWidth)];
  const used = widths.reduce((sum, value) => sum + value, 0);
  widths[widths.length - 1] += 8950 - used;
  return widths;
}

function makeBankingTable(rows, captionText = "", {
  highlightLabel,
  italicLabels = new Set(),
} = {}) {
  if (!rows || rows.length === 0) return new Paragraph({ text: "" });

  const cols = Object.keys(rows[0]).filter((k) => k !== "_header");
  const fixedWidths = captionText && TABLE_WIDTHS[captionText]
    ? TABLE_WIDTHS[captionText]
    : getBankingTableWidths(cols.length);
  const totalIdx = cols.findIndex((col) => String(col).trim().toLowerCase() === "total");
  const percentIdx = cols.findIndex((col) => /^percent/i.test(String(col).trim()));
  const ageColsEnd = totalIdx === -1 ? cols.length : totalIdx;
  const ageCols = cols.slice(1, ageColsEnd);
  const trailingCols = cols.slice(ageColsEnd);
  const normalize = (value) => String(value ?? "").trim().toLowerCase();
  const firstTrailingIdx = totalIdx === -1 ? cols.length : totalIdx;
  const labelSeparatorIdx = ageColsEnd - 1;

  const makeRow = (cellConfigs) => new TableRow({
    children: cellConfigs.map((cell) => makeStyledTableCell(cell.text, cell.options)),
  });

  const topHeaderCells = [
    {
      text: "",
      options: {
        width: { size: fixedWidths[0], type: WidthType.DXA },
        borders: makeCellBorders({ bottom: true }),
      },
    },
    {
      text: "Age",
      options: {
        width: { size: fixedWidths.slice(1, ageColsEnd).reduce((sum, value) => sum + value, 0), type: WidthType.DXA },
        columnSpan: ageCols.length,
        italic: true,
        align: AlignmentType.LEFT,
        borders: makeCellBorders({ bottom: true }),
      },
    },
  ];

  trailingCols.forEach((_, trailingIdx) => {
    const colIdx = firstTrailingIdx + trailingIdx;
    topHeaderCells.push({
      text: "",
      options: {
        width: { size: fixedWidths[colIdx], type: WidthType.DXA },
        borders: makeCellBorders({ bottom: true, right: colIdx < cols.length - 1 }),
      },
    });
  });

  const headerCells = cols.map((col, colIdx) => ({
    text: colIdx === 0 ? "" : col,
    options: {
      width: { size: fixedWidths[colIdx], type: WidthType.DXA },
      shading: HDR_FILL,
      bold: colIdx > 0,
      align: colIdx === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
      borders: makeCellBorders({
        bottom: true,
        right: colIdx === labelSeparatorIdx || colIdx === totalIdx,
      }),
    },
  }));

  const dataRows = rows.filter((row) => !row._header).map((rowObj) => {
    const label = String(rowObj[cols[0]] ?? "").trim();
    const isYouthCount = label === "Number of Youth";
    const isHighlight = highlightLabel && normalize(label) === normalize(highlightLabel);

    return new TableRow({
      children: cols.map((col, colIdx) => {
        let text = String(rowObj[col] ?? "");
        if (colIdx === 0 && isHighlight && /^currently have a bank account$/i.test(text)) {
          text = `${text}:`;
        }

        const isLabelCell = colIdx === 0;
        const fill = isYouthCount ? HDR_FILL : (isHighlight ? BANKING_HIGHLIGHT_FILL : undefined);
        const align = isLabelCell ? AlignmentType.LEFT : AlignmentType.CENTER;
        const italic = isLabelCell && (isYouthCount || italicLabels.has(label));
        const bold = isYouthCount || isHighlight;
        const indentLeft = isLabelCell && italicLabels.has(label) ? 240 : 0;

        return makeStyledTableCell(text, {
          width: { size: fixedWidths[colIdx], type: WidthType.DXA },
          shading: fill,
          bold,
          italic,
          indentLeft,
          align,
          borders: makeCellBorders({
            bottom: true,
            right: colIdx === labelSeparatorIdx || colIdx === totalIdx,
          }),
        });
      }),
    });
  });

  return new Table({
    width: { size: fixedWidths.reduce((sum, value) => sum + value, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: {
      top: NO_BORDER.top,
      bottom: NO_BORDER.bottom,
      left: NO_BORDER.left,
      right: NO_BORDER.right,
      insideH: NO_BORDER.top,
      insideV: NO_BORDER.left,
    },
    rows: [makeRow(topHeaderCells), makeRow(headerCells), ...dataRows],
  });
}

/**
 * makeTable(rows, totalLabel)
 * rows: array of plain objects. Columns are derived from Object.keys(rows[0]).
 * Special key "_header" on the first row triggers header styling.
 * Total rows: first cell value === totalLabel → header fill + bold.
 * Width: 9360 DXA total, columns evenly distributed.
 */
function makeTable(rows, totalLabel = "Total", captionText = "") {
  if (!rows || rows.length === 0) return new Paragraph({ text: "" });

  const cols        = Object.keys(rows[0]).filter((k) => k !== "_header");
  const fixedWidths = captionText ? TABLE_WIDTHS[captionText] : null;

  const tableRows = rows.map((rowObj) => {
    const isHeader = rowObj._header === true;
    const firstVal = String(rowObj[cols[0]] ?? "").trim();
    const isTotal  = !isHeader && firstVal === totalLabel;
    const shaded   = isHeader || isTotal;

    const cells = cols.map((col, colIdx) => {
      const cellText  = String(rowObj[col] ?? "");
      const leadingWhitespace = colIdx === 0 ? (cellText.match(/^(\s+)/)?.[1].length ?? 0) : 0;
      const displayText = colIdx === 0 ? cellText.replace(/^\s+/, "") : cellText;
      const cellWidth = fixedWidths
        ? { size: fixedWidths[colIdx] ?? 0, type: WidthType.DXA }
        : { size: 0, type: WidthType.AUTO };
      return makeStyledTableCell(displayText, {
        width: cellWidth,
        shading: shaded ? HDR_FILL : undefined,
        bold: shaded,
        indentLeft: leadingWhitespace * 120,
      });
    });

    return new TableRow({ children: cells });
  });

  const tblWidth = fixedWidths
    ? { size: fixedWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA }
    : { size: 0, type: WidthType.AUTO };

  return new Table({
    width: tblWidth,
    layout: fixedWidths ? TableLayoutType.FIXED : TableLayoutType.AUTOFIT,
    borders: {
      top:     { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideH: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideV: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: tableRows,
  });
}

function makeGroupedEducationTable(rows, captionText = "") {
  if (!rows || rows.length === 0) return new Paragraph({ text: "" });

  const cols = Object.keys(rows[0]).filter((k) => k !== "_header");
  const fixedWidths = captionText && TABLE_WIDTHS[captionText]
    ? TABLE_WIDTHS[captionText]
    : [1080, 5280, 1200, 1800];
  const groupFill = "EEF3F8";

  const renderRow = (rowObj, { shading, bold = false, indentLevel = false } = {}) => new TableRow({
    children: cols.map((col, colIdx) => makeStyledTableCell(rowObj[col] ?? "", {
      width: { size: fixedWidths[colIdx] ?? 0, type: WidthType.DXA },
      shading,
      bold,
      align: colIdx >= 2 ? AlignmentType.CENTER : undefined,
      indentLeft: indentLevel && col === "Level" ? 240 : 0,
    })),
  });

  const headerRow = new TableRow({
    children: cols.map((col, colIdx) => makeStyledTableCell(col, {
      width: { size: fixedWidths[colIdx] ?? 0, type: WidthType.DXA },
      bold: true,
      shading: HDR_FILL,
      align: colIdx >= 2 ? AlignmentType.CENTER : undefined,
    })),
  });

  const tableRows = [headerRow];
  let currentGroup = null;

  for (const rowObj of rows.filter((row) => !row._header)) {
    const groupLabel = String(rowObj.Group ?? "").trim();
    const levelLabel = String(rowObj.Level ?? "").trim();
    const isGrandTotal = groupLabel === "Total";
    const isSubtotal = !isGrandTotal && /^Total\s+/i.test(levelLabel);

    if (!isGrandTotal && groupLabel && groupLabel !== currentGroup) {
      tableRows.push(new TableRow({
        children: [makeStyledTableCell(groupLabel, {
          width: { size: fixedWidths.reduce((sum, value) => sum + value, 0), type: WidthType.DXA },
          bold: true,
          shading: groupFill,
          columnSpan: cols.length,
        })],
      }));
      currentGroup = groupLabel;
    }

    if (isGrandTotal) {
      tableRows.push(renderRow(rowObj, { shading: HDR_FILL, bold: true }));
      continue;
    }

    tableRows.push(renderRow({
      ...rowObj,
      Group: "",
    }, {
      shading: isSubtotal ? groupFill : undefined,
      bold: isSubtotal,
      indentLevel: !isSubtotal,
    }));
  }

  return new Table({
    width: { size: fixedWidths.reduce((sum, value) => sum + value, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: {
      top:     { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideH: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideV: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: tableRows,
  });
}

/**
 * makeHeading(text, level)
 * level 1 → HeadingLevel.HEADING_1, level 2 → HeadingLevel.HEADING_2
 */
function makeHeading(text, level = 1) {
  return new Paragraph({
    text,
    heading: level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
  });
}

/**
 * makeCaption(text) — bold paragraph using Caption style
 */
function makeCaption(text) {
  return new Paragraph({
    style: "Caption",
    children: [new TextRun({ text, bold: true, font: "Calibri", size: 22 })],
  });
}

/**
 * makePara(text, options)
 * options: { bold, italic, indent }
 */
function makePara(text, options = {}) {
  const { bold = false, italic = false, indent = false } = options;
  const para = new Paragraph({
    style: indent ? "ListParagraph" : "Normal",
    spacing: BODY_PARAGRAPH_SPACING,
    children: [new TextRun({ text, bold, italic, font: "Calibri", size: 22 })],
  });
  if (!String(text).trim()) SPACER_PARAGRAPHS.add(para);
  return para;
}

/**
 * makeBullet(text) — bullet list paragraph
 */
function makeBullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: BODY_PARAGRAPH_SPACING,
    children: [new TextRun({ text, font: "Calibri", size: 22 })],
  });
}

/**
 * makePlaceholder(text) — yellow-highlighted paragraph
 */
function makePlaceholder(text) {
  return new Paragraph({
    spacing: BODY_PARAGRAPH_SPACING,
    children: [
      new TextRun({
        text: `[${text}]`,
        highlight: "yellow",
        bold: true,
        font: "Calibri",
        size: 22,
      }),
    ],
  });
}

// ---------------------------------------------------------------------------
// Helpers — data loading
// ---------------------------------------------------------------------------

/**
 * loadSheet(sheetName)
 * Returns rows as array of objects (first data row = header).
 * Row 0 of the xlsx sheet is the section title; row 1 is the column header.
 * The first returned object has _header: true.
 */
function loadSheet(sheetName) {
  const wb  = XLSX.readFile(ANALYSIS_PATH);
  const ws  = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);

  // Get all rows as raw arrays
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  // raw[0] = section title, raw[1] = column headers, raw[2+] = data
  if (raw.length < 2) return [];

  const headers = raw[1].map((h) => String(h ?? ""));
  const dataRows = raw.slice(2).filter((r) => r.some((v) => String(v).trim() !== ""));

  // Build header row object
  const headerObj = { _header: true };
  headers.forEach((h, i) => { headerObj[h || `col${i}`] = h; });

  const result = [headerObj];
  for (const rawRow of dataRows) {
    const obj = {};
    headers.forEach((h, i) => {
      const key = h || `col${i}`;
      const val = rawRow[i];
      obj[key] = val === null || val === undefined ? "" : String(val);
    });
    result.push(obj);
  }
  return result;
}

// ---------------------------------------------------------------------------
// Section: sec_age
// ---------------------------------------------------------------------------

function sec_age() {
  const rows = loadSheet("01_age");
  const renamed = rows.map((r) => {
    if (r._header) return { _header: true, Age: "Age", Count: "Count" };
    const entries = Object.entries(r);
    return { Age: entries[0]?.[1] ?? "", Count: entries[1]?.[1] ?? "" };
  });
  const cap = "Survey Respondents by Age";
  return [
    makeCaption(cap),
    makeTable(renamed, "Total", cap),
    makePara(""),
  ];
}

// ---------------------------------------------------------------------------
// Additional data helpers
// ---------------------------------------------------------------------------

/** pct(n, d, decimals) — returns "XX%" string, like Python's pct() */
function pct(n, d, decimals = 0) {
  if (!d) return "";
  const val = (100 * n) / d;
  return decimals === 0 ? `${Math.round(val)}%` : `${val.toFixed(decimals)}%`;
}

/** firstCol(row) — value of the first non-_header key in a row object */
function firstCol(row) {
  return String(Object.entries(row).filter(([k]) => k !== "_header")[0]?.[1] ?? "");
}

/** getCol(row, n) — value of the nth column (0-based, ignoring _header) */
function getCol(row, n) {
  const keys = Object.keys(row).filter((k) => k !== "_header");
  return String(row[keys[n]] ?? "");
}

/**
 * splitSheet(sheetName)
 * Mirrors Python _split_sheet(): splits a sheet into named sub-tables.
 * Section-title rows = first cell non-blank, all other cells blank.
 * Returns { sectionTitle: rowObjects[] } where rowObjects[0] has _header:true.
 */
function splitSheet(sheetName) {
  const wb = XLSX.readFile(ANALYSIS_PATH);
  const ws = wb.Sheets[sheetName];
  if (!ws) {
    console.warn(`Warning: Sheet not found: ${sheetName}`);
    return {};
  }
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  const sections = {};
  let currentTitle = null;
  let currentRows  = [];

  for (const rawRow of raw) {
    const first   = String(rawRow[0] ?? "").trim();
    const isTitle = first !== "" && rawRow.slice(1).every((v) => String(v ?? "").trim() === "");
    if (isTitle) {
      if (currentTitle !== null && currentRows.length > 0) {
        sections[currentTitle] = currentRows;
      }
      currentTitle = first;
      currentRows  = [];
    } else if (currentTitle !== null) {
      currentRows.push(rawRow);
    }
  }
  if (currentTitle !== null && currentRows.length > 0) {
    sections[currentTitle] = currentRows;
  }

  const result = {};
  for (const [title, rows] of Object.entries(sections)) {
    const nonBlank = rows.filter((r) => r.some((v) => String(v ?? "").trim() !== ""));
    if (nonBlank.length === 0) { result[title] = []; continue; }

    let maxCol = 0;
    for (const r of nonBlank) {
      for (let i = r.length - 1; i >= 0; i--) {
        if (String(r[i] ?? "").trim() !== "") { maxCol = Math.max(maxCol, i + 1); break; }
      }
    }
    const headers = nonBlank[0].slice(0, maxCol).map((h) => String(h ?? ""));
    const headerObj = { _header: true };
    headers.forEach((h, i) => { headerObj[h || `col${i}`] = h; });

    const dataRows = nonBlank.slice(1).map((r) => {
      const obj = {};
      headers.forEach((h, i) => { obj[h || `col${i}`] = String(r[i] ?? ""); });
      return obj;
    });
    result[title] = [headerObj, ...dataRows];
  }
  return result;
}

/** loadCsv() — returns array of row objects from survey_data_412YZ.csv */
function loadCsv() {
  let content = fs.readFileSync(CSV_PATH, "utf8");
  if (content.charCodeAt(0) === 0xFEFF) content = content.slice(1); // strip BOM
  return parseCsv(content, { columns: true, skip_empty_lines: true });
}

// ---------------------------------------------------------------------------
// Section functions (in report order)
// ---------------------------------------------------------------------------

function sec_title() {
  return [
    makeHeading("Youth Zone Survey Results", 1),
    makePara(SURVEY_MONTH),
    makePara(""),
    makePara(
      "All individuals active with the 412 Youth Zone had the opportunity to " +
      "participate in a survey in early March 2026. Surveys were administered on paper " +
      "and online."
    ),
    new Paragraph({
      spacing: BODY_PARAGRAPH_SPACING,
      children: [
        new TextRun({ text: `${N_RESPONDENTS} unique youth (of `, font: "Calibri", size: 22 }),
        new TextRun({ text: "840", font: "Calibri", size: 22 }),
        new TextRun({ text: " total active) responded to the survey, for a response rate of ", font: "Calibri", size: 22 }),
        new TextRun({ text: "24%", font: "Calibri", size: 22 }),
        new TextRun({ text: ". Most respondents were age 18 or older. About half (49%) of the respondents who reported their age were between 18 and 20 years old, and 40% were 21 to 23 years old. Only 10% were 16 or 17 years old, down from 16% in 2025.", font: "Calibri", size: 22 }),
      ],
    }),
    makePara(""),
  ];
}

function sec_gender_orient() {
  const rows = loadSheet("02_gender_orient");
  const data = rows.filter((r) => !r._header);
  const count = (row, key) => parseInt(row?.[key] || "0") || 0;
  const numRow = data.find((r) => firstCol(r) === "Number of Youth");
  const lgbtqLabels = new Set([
    "Asexual",
    "Bisexual",
    "Demisexual",
    "Gay, Lesbian, or Same Gender Loving",
    "Mostly heterosexual",
    "Pansexual",
    "Queer",
  ]);
  let nF = 0, nM = 0, nNB = 0;
  if (numRow) {
    nF  = parseInt(numRow["Female"]           || "0") || 0;
    nM  = parseInt(numRow["Male"]             || "0") || 0;
    nNB = parseInt(numRow["Trans, Non-binary"]|| "0") || 0;
  }
  const nKnown = nF + nM + nNB;
  const lgbtqRows = data.filter((r) => lgbtqLabels.has(firstCol(r)));
  const nLgbtqF = lgbtqRows.reduce((sum, row) => sum + count(row, "Female"), 0);
  const nLgbtqM = lgbtqRows.reduce((sum, row) => sum + count(row, "Male"), 0);
  const nLgbtq = lgbtqRows.reduce((sum, row) => sum + count(row, "Total"), 0);
  const pctF = pct(nF, nKnown);
  const pctM = pct(nM, nKnown);
  const pctNB = pct(nNB, nKnown);
  const pctLgbtqF = pct(nLgbtqF, nF);
  const pctLgbtqM = pct(nLgbtqM, nM);
  const pctLgbtq = pct(nLgbtq, nKnown);
  const cap = "Survey Respondents by Gender and Sexual Orientation";
  return [
    makePara(
      `More females (${pctF}) responded to the survey than males (${pctM}). ` +
      `Twenty-one transgender and non-binary youth also responded to the survey (${pctNB}), up from 12 youth (7%) in 2025. ` +
      "As represented in the table below, youth of all genders selected a variety of terms to describe their sexual orientation. " +
      `Young women were more likely than their male peers to identify as LGBTQ (${pctLgbtqF} v. ${pctLgbtqM}), and ${pctLgbtq} of respondents identified as LGBTQ in some way.`
    ),
    makeCaption(cap),
    makeTable(rows, "Total", cap),
    makePara(""),
  ];
}

async function sec_race() {
  const rowsOnce = loadSheet("03_race_once");
  const rowsMulti = loadSheet("04_race_multi");
  const cap1 = "Youth by Race and Gender (all Youth are Counted Once)";
  const cap2 = "Youth with Full or Partial Racial Identities (Some Youth Are Counted Multiple Times)";
  return [
    makePara(
      "The tables below display respondents\u2019 racial identities and genders. " +
      "Please note: Youth Zone participants\u2019 racial identities are self-reported " +
      "and reflect the full range of how young people describe themselves."
    ),
    makePara(
      "Just over half of survey respondents identified as Black when each youth was counted once, " +
      "while 28% identified as White, 17% as Multiracial, and 3% reported another single racial identity. " +
      "When full or partial identities are counted using the unique respondent total, 62% of youth identified as Black, " +
      "31% as White, 11% as Multiracial, and 6% as Hispanic or Latinx. " +
      "This marks a change from the prior year, when 66% of respondents identified as Black either fully or partially. " +
      "White identification increased from 27% to 31%, while Multiracial identification remained steady at 11% " +
      "and Hispanic or Latinx identification rose from 4% to 6%."
    ),
    makeCaption(cap1),
    makeTable(rowsOnce, "Total", cap1),
    makePara("Note: If a youth selected 2 or more race tokens, they are attributed to \"Multiracial\" only.", { italic: true }),
    makeCaption(cap2),
    makeTable(rowsMulti, "Total", cap2),
    makePara(""),
  ];
}

function sec_coach_satisfaction() {
  const rows05 = loadSheet("05_q1");
  const data05 = rows05.filter((r) => !r._header);
  const cols05 = Object.keys(rows05[0] || {}).filter((k) => k !== "_header");
  const hasCols = cols05.length > 2;

  const trustRow = data05.find((r) => String(r[cols05[0]] ?? "").toLowerCase().includes("trustworthy"));
  const pctTrust = (trustRow && hasCols) ? getCol(trustRow, 2) : "[PLACEHOLDER]";
  const valsRow  = data05.find((r) => String(r[cols05[0]] ?? "").toLowerCase().includes("values"));
  const pctVals  = (valsRow  && hasCols) ? getCol(valsRow,  2) : "[PLACEHOLDER]";

  const currPct = {};
  for (const row of data05) {
    const label  = String(row[cols05[0]] ?? "").trim();
    const pctVal = hasCols ? getCol(row, 2) : "";
    currPct[label] = pctVal;
  }

  const priorCols = Q1_BENCHMARKS.col;
  const currN     = `n=${N_RESPONDENTS}`;

  const FIELD_LABELS = [
    "Is trustworthy",
    "Is reliable",
    "Values my opinions about my life",
    "Is available to me when I need them",
    "Makes me feel heard and understood",
  ];
  const dataRowsQ1 = FIELD_LABELS.map((label) => {
    return {
      label,
      current: currPct[label] || "",
    };
  });

  const cap = "Satisfaction Ratings for Youth Coaches Over Time";
  return { pctTrust, pctVals, currN, dataRowsQ1, cap };
}

async function sec_coach_satisfaction_async() {
  const { pctTrust, pctVals, currN, dataRowsQ1, cap } = sec_coach_satisfaction();
  return [
    makeHeading("FINDINGS", 1),
    makePara(""),
    makeHeading("Relationships with Coach", 2),
    makePara(
      `${pctTrust} of youth reported their coaches were trustworthy, and ` +
      `${pctVals} indicated their coach values their opinions about their life \u2014 both consistent with February 2025. ` +
      "However, ratings for coach availability and reliability declined more notably this year: " +
      "\u201CIs available to me when I need them\u201D fell 8 percentage points to 80%, and " +
      "\u201CIs reliable\u201D fell 5 percentage points to 84%, the lowest ratings recorded for those items " +
      "across all survey cycles shown in the table below."
    ),
    makePara(
      "Among the responses where youth did not rate their coaches in the top two categories, " +
      "most scores indicated their coach Sometimes exhibits the listed characteristic; " +
      "the availability and reliability items drew the most non-top-2 responses, " +
      "consistent with the declines noted above. " +
      "A small number of youth said their coach Rarely or Never exhibits one or more of these qualities."
    ),
    makeCaption(cap),
    makeCoachSatisfactionTable(dataRowsQ1, currN, cap),
    await embedChart("chart_01_coach_satisfaction.png", 5.5),
    makePara(""),
  ];
}

async function sec_communication() {
  // Q3 communication level counts from analysis sheet
  const rows = loadSheet("06_communication");
  const data = rows.filter((r) => !r._header);
  const cols = Object.keys(rows[0] || {}).filter((k) => k !== "_header");
  let goodPct = "", notEnoughPct = "", tooMuchPct = "";
  for (const row of data) {
    const first = String(row[cols[0]] ?? "").trim().toLowerCase();
    if (first === "good amount") {
      goodPct = cols.length > 2 ? String(row[cols[2]] ?? "") : "";
    }
    if (first === "not enough") {
      notEnoughPct   = cols.length > 2 ? String(row[cols[2]] ?? "") : "";
    }
    if (first === "too much") {
      tooMuchPct = cols.length > 2 ? String(row[cols[2]] ?? "") : "";
    }
  }

  // Q2 frequency breakdown from raw CSV
  const dfCsv = loadCsv();
  const goodRows     = dfCsv.filter((r) => r.q3_communication_level === "good_amount");
  const neRows       = dfCsv.filter((r) => r.q3_communication_level === "not_enough");
  const goodQ2Total  = goodRows.filter((r) => (r.q2_communication_frequency || "").trim()).length;
  const nGoodMonthly = goodRows.filter((r) => r.q2_communication_frequency === "1_2_times_per_month").length;
  const nGoodWeekly  = goodRows.filter((r) => r.q2_communication_frequency === "about_once_a_week").length;
  const nGoodDaily   = goodRows.filter((r) => r.q2_communication_frequency === "almost_every_day").length;
  const pctGoodMonthly = goodQ2Total ? pct(nGoodMonthly, goodQ2Total) : "";
  const pctGoodWeekly  = goodQ2Total ? pct(nGoodWeekly,  goodQ2Total) : "";
  const pctGoodDaily   = goodQ2Total ? pct(nGoodDaily,   goodQ2Total) : "";
  const notEnoughRespondentCount = neRows.length;
  const nNeMonthly   = neRows.filter((r) => r.q2_communication_frequency === "1_2_times_per_month").length;
  const nNeLess      = neRows.filter((r) => r.q2_communication_frequency === "less_than_once_a_month").length;

  const para1 = (goodPct && notEnoughPct && tooMuchPct && notEnoughRespondentCount)
    ? `Most respondents rated the amount of communication with their coach positively. ` +
      `${goodPct} reported the amount was a Good amount, while ${notEnoughPct} (${notEnoughRespondentCount} youth) reported it was Not enough and ${tooMuchPct} reported Too much. ` +
      `This is a slight increase from February 2025, when 11% (19 youth) reported their communication was Not enough.`
    : "";

  const para2 = goodQ2Total
    ? `Among youth who rated communication a Good amount, nearly half communicated ` +
      `1\u20132 times per month (${pctGoodMonthly}), about a third connected with their coach ` +
      `about once a week (${pctGoodWeekly}), and ${pctGoodDaily} communicated almost every day.`
    : "";

  const para3 = (notEnoughPct && notEnoughRespondentCount)
    ? `${notEnoughPct} of respondents (n=${notEnoughRespondentCount}) reported their communication with their coach was Not enough. ` +
      `Most of these youth communicated 1\u20132 times per month (${nNeMonthly} youth) or less than once a month (${nNeLess} youth).`
    : "";

  const items = [
    makePara(para1),
    await embedChart("chart_07_communication_satisfaction.png", 4.5),
    makePara(para2),
    makePara(para3),
    await embedChart("chart_08_communication_freq_not_enough.png", 4.5),
    makePara(""),
  ];
  return items;
}

async function sec_housing() {
  const rows07 = loadSheet("07_housing");
  const data07 = rows07.filter((r) => !r._header);
  const stableRow = data07.find((r) => {
    const v = firstCol(r).toLowerCase();
    return v.includes("stable") && !v.includes("un");
  });
  const pctStable = (stableRow && stableRow["Percent"]) ? stableRow["Percent"] : "[PLACEHOLDER]";
  const rows08 = loadSheet("08_housing_reasons");
  const cap1 = "Housing Status and Current Sleeping Arrangements";
  const cap2 = "Reasons for Unstable Housing in the Past 6 Months, by Age\u00b9";
  return [
    makeHeading("Stable Housing", 1),
    makePara(
      "Survey respondents were asked to describe whether their current housing is " +
      "safe and stable, meaning they can stay there for at least the next 90 days. " +
      "Youth with unstable housing were then asked where they are currently sleeping."
    ),
    makePara(
      `About ${pctStable} of respondents reported safe and stable housing, a decrease of approximately ` +
      "10 percentage points from 2025. Among unstably housed youth, staying with family or friends was " +
      "the most commonly reported current sleeping arrangement, followed by couch surfing."
    ),
    makeCaption(cap1),
    makeTable(rows07, "Total", cap1),
    makePara("Current housing status varied by age group, as shown in the chart below."),
    await embedChart("chart_02_housing_stability.png", 5.5),
    makePara(
      "Youth ages 16\u201317 reported stable housing at a notably higher rate (89%) than " +
      "respondents ages 18\u201320 and 21\u201323 (70% and 67%, respectively). " +
      "Among 18\u201320 year olds, \u201Cno place to stay\u201D was the most prevalent unstable category (17%). " +
      "Youth ages 21\u201323 were more likely than other groups to report housing they could stay in but considered unsafe (7%)."
    ),
    makePara(
      "Regardless of their current living situation, if youth experienced unstable " +
      "housing in the prior six months, they were asked to share why. " +
      "The table below displays the answers by age. " +
      "The most common reasons youth experienced housing instability were that family or friends " +
      "could no longer let them stay (49%) and feeling unsafe at home (39%), consistent with the " +
      "pattern seen in the prior year. Among youth ages 21\u201323, feeling unsafe was the leading " +
      "reason cited, while among 18\u201320 year olds, family or friends being unavailable was the " +
      "more predominant reason."
    ),
    makeCaption(cap2),
    makeTable(rows08, "Total youth reporting unstable housing", cap2),
    makePara("\u00b9 Youth could report more than one reason for experiencing unstable housing.", { italic: true }),
    makePara("Note: \u2018Other\u2019 responses (n=15) included lease expiration or rent increases, unsafe or inadequate housing conditions (e.g., landlord issues, maintenance failures, pest infestation), and domestic or family safety concerns.", { italic: true }),
    makePara(""),
  ];
}

async function sec_education_employment(dfCsv) {
  // Read employed count and Q8 total directly from the 10_employment sheet
  // so narrative percentages match the table.
  const rows10raw = loadSheet("10_employment");
  const data10    = rows10raw.filter((r) => !r._header);
  const totalRow10 = data10.find((r) => firstCol(r) === "Total");
  const nQ8Total  = totalRow10 ? (parseInt(totalRow10["Total"]) || dfCsv.length) : dfCsv.length;
  const ftRow  = data10.find((r) => firstCol(r) === "Full time");
  const ptRow  = data10.find((r) => firstCol(r) === "Part time");
  const nFT    = ftRow ? (parseInt(ftRow["Total"]) || 0) : 0;
  const nPT    = ptRow ? (parseInt(ptRow["Total"]) || 0) : 0;
  const employed = nFT + nPT;

  const total      = dfCsv.length;
  const inSchool   = dfCsv.filter((r) => ["high_school","college_career","ged","graduate"].includes(r.q5_school_status)).length;
  const inSRows    = dfCsv.filter((r) => ["high_school","college_career","ged","graduate"].includes(r.q5_school_status));
  const inSUnemp   = inSRows.filter((r) => r.q8_employment_status === "no");
  const inSUnempSk = inSUnemp.filter((r) => r.q8b_job_seeking === "yes").length;
  const notSUnemp  = dfCsv.filter((r) => r.q5_school_status === "not_in_school" && r.q8_employment_status === "no");
  const notSUnempSk= notSUnemp.filter((r) => r.q8b_job_seeking === "yes").length;
  const nisuRows   = dfCsv.filter((r) => r.q5_school_status === "not_in_school" && r.q8_employment_status === "no");
  const noDiploma  = nisuRows.filter((r) => r.q5a_highest_education === "some_hs").length;

  const rows09 = loadSheet("09_education");
  const data09 = rows09.filter((r) => !r._header);
  const notInSchoolTotalRow = data09.find((r) => firstCol(r) === "Not in School" && String(r.Level ?? "").trim() === "Total Not in School");
  const hsDiplomaRow = data09.find((r) => String(r.Level ?? "").trim() === "HS Diploma or GED");
  const someCollegeRow = data09.find((r) => String(r.Level ?? "").trim() === "Some College");
  const degreeRow = data09.find((r) => String(r.Level ?? "").trim() === "College Degree or Certificate");
  const notInSchoolTotal = notInSchoolTotalRow ? (parseInt(notInSchoolTotalRow.Count) || 0) : 0;
  const graduateRow = data09.find((r) => String(r.Level ?? "").trim() === "Graduate School");
  const collegeVocationalRow = data09.find((r) => String(r.Level ?? "").trim() === "College/Vocational");
  const gedProgramRow = data09.find((r) => String(r.Level ?? "").trim() === "GED Program");
  const highSchoolRow = data09.find((r) => String(r.Level ?? "").trim() === "High School");
  const totalEducationRow = data09.find((r) => firstCol(r) === "Total");
  const highSchoolCount = highSchoolRow ? (parseInt(highSchoolRow.Count) || 0) : 0;
  const educationTotal = totalEducationRow ? (parseInt(totalEducationRow.Count) || 0) : total;
  const hsDiplomaOrGed = (graduateRow ? (parseInt(graduateRow.Count) || 0) : 0) +
    (collegeVocationalRow ? (parseInt(collegeVocationalRow.Count) || 0) : 0) +
    (hsDiplomaRow ? (parseInt(hsDiplomaRow.Count) || 0) : 0) +
    (someCollegeRow ? (parseInt(someCollegeRow.Count) || 0) : 0) +
    (degreeRow ? (parseInt(degreeRow.Count) || 0) : 0);
  const higherEdOrDegree = (graduateRow ? (parseInt(graduateRow.Count) || 0) : 0) +
    (collegeVocationalRow ? (parseInt(collegeVocationalRow.Count) || 0) : 0) +
    (degreeRow ? (parseInt(degreeRow.Count) || 0) : 0);
  const cap = "Educational Enrollment and Attainment";
  const items = [
    makeHeading("Employment and Education", 1),
    makePara(
      "Youth were asked whether they are attending school, working, or looking for work."
    ),
    makeBullet(
      `This year, enrollment and employment rates were nearly equal — ${pct(inSchool, total)} enrolled in school, ` +
      `${pct(employed, nQ8Total)} employed — compared to 53% enrolled and 44% employed in March 2025.`
    ),
    makeBullet(
      `Among youth who are neither in school nor employed (${pct(nisuRows.length, total)} of all respondents, down from 26% last year), ` +
      `job-seeking rates were high: ${pct(notSUnempSk, notSUnemp.length)} reported actively looking for work, up significantly from 68% last year. ` +
      `${noDiploma} of these ${nisuRows.length} youth had not completed high school or a GED.`
    ),
  ];
  if (inSUnemp.length) items.push(makeBullet(
    `Job-seeking was also common among youth in school but not working: ${pct(inSUnempSk, inSUnemp.length)} reported looking for a job, ` +
    `down slightly from 78% last year.`
  ));
  items.push(makePara(
    "Among youth who reported being enrolled in school, respondents were split almost evenly between high school students (45%) and youth in college, vocational, or graduate programs (51% combined), while 4% reported being enrolled in a GED program. " +
    "This is a notable change from March 2025, when about two-thirds of enrolled youth were in high school, suggesting this year’s in-school respondents were more evenly distributed across secondary and postsecondary education."
  ));
  if (notInSchoolTotal) items.push(makePara(
    `When excluding the ${highSchoolCount} youth enrolled in high school, ` +
    `${pct(hsDiplomaOrGed, educationTotal - highSchoolCount)} of respondents have their high school diploma or GED, up from 85% in March 2025. ` +
    `${pct(higherEdOrDegree, educationTotal)} of respondents are currently enrolled in higher education or have already received a college degree or certificate, up notably from 18% last year — driven largely by a higher share of enrolled respondents being in college or vocational programs this year compared to last.`
  ));
  items.push(makeCaption(cap));
  items.push(makeGroupedEducationTable(rows09, cap));
  const educationChartPath = path.join(CHARTS_DIR, "chart_09_employment_by_school.png");
  if (fs.existsSync(educationChartPath)) {
    items.push(await embedChart("chart_09_employment_by_school.png", 6.6));
  }
  return items;
}

function sec_job_tenure(dfCsv) {
  const rows11 = loadSheet("11_job_tenure");
  const data11 = rows11.filter((r) => !r._header);
  const totRow  = data11.find((r) => firstCol(r) === "Total");
  const shortRow = data11.find((r) => firstCol(r).toLowerCase().includes("less than 3"));
  const midRow = data11.find((r) => firstCol(r).toLowerCase().includes("3 to 6"));
  const nEmp    = totRow ? (parseInt(totRow["Total"]) || 0) : 0;
  const shortTenure = shortRow ? (parseInt(shortRow["Total"]) || 0) : 0;
  const midTenure = midRow ? (parseInt(midRow["Total"]) || 0) : 0;
  const longRow = data11.find((r) => firstCol(r).toLowerCase().includes("more than 6"));
  const longTenure = longRow ? (parseInt(longRow["Total"]) || 0) : 0;
  // nQ8Total: read from 10_employment sheet to get consistent denominator
  const rows10 = loadSheet("10_employment");
  const totRow10 = rows10.filter((r) => !r._header).find((r) => firstCol(r) === "Total");
  const nQ8Total = totRow10 ? (parseInt(totRow10["Total"]) || dfCsv.length) : dfCsv.length;
  const cap = "Length of Employment for Youth Currently Employed";
  return [
    makePara(
      `Of the ${pct(nEmp, nQ8Total)} of survey respondents who reported being employed, ` +
      `${pct(longTenure, nEmp)} have been at their job for more than six months, ` +
      `${pct(midTenure, nEmp)} have been at their current job for three to six months, and ` +
      `${pct(shortTenure, nEmp)} have been working there for less than three months.`
    ),
    makePara(
      "Longer-term employment was more common among youth working full time than part time, and among both 18 to 20 year olds and 21 to 23 year olds the largest share of employed youth had already been in their jobs for more than six months. " +
      "This suggests that once youth are connected to work, many are maintaining employment for at least part of the year."
    ),
    makeCaption(cap),
    makeTable(rows11, "Total", cap),
    makePara(""),
  ];
}

async function sec_employment_by_age() {
  const rows10 = loadSheet("10_employment");
  const cap = "Employment Status by Age";
  return [
    makePara(
      "Employment status varied notably by age. Youth ages 16 to 17 were the least likely to be working, with most reporting they were not employed, while youth ages 21 to 23 were the most likely to report full-time work. " +
      "Respondents ages 18 to 20 fell between those two groups and were more likely to report part-time than full-time employment."
    ),
    makeCaption(cap),
    makeTable(rows10, "Total", cap),
    await embedChart("chart_03_employment_by_age.png", 5.5),
    makePara(""),
  ];
}

function sec_job_barriers(dfCsv) {
  const rows12  = loadSheet("12_job_barriers");
  const data12  = rows12.filter((r) => !r._header);
  const cols12  = Object.keys(rows12[0] || {}).filter((k) => k !== "_header");
  const topRow  = data12[0];
  const secondRow = data12[1];
  const thirdRow = data12[2];
  const topLabel = topRow ? firstCol(topRow).toLowerCase() : "";
  const topPct   = (topRow && cols12.length > 2) ? getCol(topRow, 2) : "";
  const secondLabel = secondRow ? firstCol(secondRow).toLowerCase() : "";
  const secondPct = (secondRow && cols12.length > 2) ? getCol(secondRow, 2) : "";
  const thirdLabel = thirdRow ? firstCol(thirdRow).toLowerCase() : "";
  const thirdPct = (thirdRow && cols12.length > 2) ? getCol(thirdRow, 2) : "";
  const mentalRow = data12.find((row) => firstCol(row).trim() === "Mental or physical health");
  const mentalPct = (mentalRow && cols12.length > 2) ? getCol(mentalRow, 2) : "";
  const mentalHealthRows = dfCsv.filter((row) =>
    String(row.q10_job_barriers || "").split("|").map((part) => part.trim()).includes("mental_physical_health")
  );
  const mentalHealth18to23 = mentalHealthRows.filter((row) => ["18_20", "21_23"].includes(String(row.age_range || "").trim())).length;
  const mentalHealthSentence = !mentalPct || mentalHealthRows.length === 0
    ? ""
    : mentalHealth18to23 === mentalHealthRows.length
      ? ` Mental or physical health was cited by ${mentalPct}, and all of those responses came from youth ages 18 to 23.`
      : ` Mental or physical health was cited by ${mentalPct}; ${mentalHealth18to23} of the ${mentalHealthRows.length} youth reporting that barrier were ages 18 to 23.`;
  const cap = "Reasons Youth Have Trouble Finding Jobs (Reasons Given by 2 or More People)";
  return [
    makePara(
      "If survey respondents had trouble finding a job in the prior twelve months, " +
      "they were asked to share some of the reasons they believed were contributing to those challenges. " +
      "Youth were offered a multiple-choice list of reasons and could also add their own responses."
    ),
    makePara(
      `The most common challenge identified this year, reported by ${topPct} of youth who described at least one job barrier, was ${topLabel}, followed closely by ${secondLabel} (${secondPct}) and ${thirdLabel} (${thirdPct}). ` +
      `Both transportation issues and applying and not getting called were reported more often than in March 2025, when those barriers stood at 32% and 28%, respectively.` +
      mentalHealthSentence
    ),
    makeCaption(cap),
    makeTable(rows12, "__none__", cap),
    makePara("Note: Youth could select more than one option", { italic: true }),
    makePara(""),
  ];
}

function sec_left_job() {
  const rows13 = loadSheet("13_left_job");
  const data13 = rows13.filter((r) => !r._header);
  const cols13 = Object.keys(rows13[0] || {}).filter((k) => k !== "_header");
  const findRow = (label) => data13.find((r) => firstCol(r).trim() === label);
  const foundBetterPct = getCol(findRow("Found a better job") || {}, 2) || "";
  const quitPct = getCol(findRow("Quit") || {}, 2) || "";
  const seasonalPct = getCol(findRow("Seasonal/temporary") || {}, 2) || "";
  const lowPayPct = getCol(findRow("Low pay or not enough hours") || {}, 2) || getCol(findRow("    Low pay or not enough hours") || {}, 2) || "";
  const mentalHealthPct = getCol(findRow("Mental/emotional health") || {}, 2) || getCol(findRow("    Mental/emotional health") || {}, 2) || "";
  const personalFamilyPct = getCol(findRow("Personal or family reasons") || {}, 2) || getCol(findRow("    Personal or family reasons") || {}, 2) || "";
  const cap = "Reasons Youth Lost or Quit a Job in the Past Year (Reasons Given by 2 or More People)";
  return [
    makePara(
      "If the survey respondent lost or left a job in the past year, they were asked " +
      "to share the reason(s)."
    ),
    makePara(
      `This year, finding a better job and quitting were tied as the most common reasons youth reported leaving a job, each cited by ${foundBetterPct} of youth who reported at least one reason for leaving. ${seasonalPct} said the job was seasonal or temporary. Among the youth who quit, the most common specific reasons were low pay or not enough hours (${lowPayPct}) and mental or emotional health (${mentalHealthPct}), followed by personal or family reasons (${personalFamilyPct}).`
    ),
    makeCaption(cap),
    makeTable(rows13, "__none__", cap),
    makePara(""),
  ];
}

function sec_employment_equity() {
  const dfCsv = loadCsv();
  const splitTokens = (value) => String(value || "").split("|").map((part) => part.trim()).filter(Boolean);
  const transTokens = new Set(["Non-binary", "Gender Nonconforming", "Transgender Male", "Transgender Female", "Genderqueer", "Two-Spirit"]);
  const lgbtqTokens = new Set(["Asexual", "Bisexual", "Demisexual", "Gay or Lesbian", "Same Gender Loving", "Mostly heterosexual", "Pansexual", "Queer"]);
  const tokenToRaceGroup = (token) => {
    const t = String(token || "").trim();
    if (t.includes("Black")) return "Black";
    if (t.includes("White")) return "White";
    if (t.includes("Multi")) return "Multiracial";
    if (t.includes("Hispanic") || t.includes("Latinx")) return "Hispanic or Latinx";
    if (t.includes("Asian")) return "Asian";
    if (t.includes("Native")) return "Native American or Native Hawaiian";
    if (t.includes("Prefer not")) return "Prefer not to answer";
    return "Other";
  };
  const genderBucket = (raw) => {
    const tokens = splitTokens(raw);
    if (tokens.some((token) => transTokens.has(token))) return "Trans/Non-binary/GNC";
    if (tokens.length === 1 && tokens[0] === "Female") return "Female";
    if (tokens.length === 1 && tokens[0] === "Male") return "Male";
    return "";
  };
  const raceBucket = (raw) => {
    const tokens = splitTokens(raw);
    if (!tokens.length) return "";
    if (tokens.length >= 2) return "Multiracial";
    return tokenToRaceGroup(tokens[0]);
  };
  const orientBucket = (raw) => {
    const tokens = splitTokens(raw);
    if (!tokens.length) return "";
    if (tokens.some((token) => lgbtqTokens.has(token))) return "LGBTQ+";
    if (tokens.length === 1 && (tokens[0] === "Heterosexual/Straight" || tokens[0] === "Heterosexual")) return "Heterosexual/Straight";
    return "";
  };
  const hasBarrier = (row) => String(row.q10_job_barriers || "").trim() !== "";
  const leftJob = (row) => String(row.q11_left_job_reasons || "").trim() !== "";
  const hasMentalBarrier = (row) => splitTokens(row.q10_job_barriers).includes("mental_physical_health");
  const buildSummary = (rows) => ({
    n: rows.length,
    barrierPct: pct(rows.filter(hasBarrier).length, rows.length),
    leftPct: pct(rows.filter(leftJob).length, rows.length),
    mentalPct: pct(rows.filter(hasMentalBarrier).length, rows.length),
  });
  const overall = buildSummary(dfCsv);
  const groups = {
    gender: [
      ["Female", dfCsv.filter((row) => genderBucket(row.gender) === "Female")],
      ["Male", dfCsv.filter((row) => genderBucket(row.gender) === "Male")],
      ["Trans/Non-binary/GNC", dfCsv.filter((row) => genderBucket(row.gender) === "Trans/Non-binary/GNC")],
    ],
    race: [
      ["Black", dfCsv.filter((row) => raceBucket(row.race_ethnicity) === "Black")],
      ["White", dfCsv.filter((row) => raceBucket(row.race_ethnicity) === "White")],
      ["Multiracial", dfCsv.filter((row) => raceBucket(row.race_ethnicity) === "Multiracial")],
    ],
    orientation: [
      ["Heterosexual/Straight", dfCsv.filter((row) => orientBucket(row.sexual_orientation) === "Heterosexual/Straight")],
      ["LGBTQ+", dfCsv.filter((row) => orientBucket(row.sexual_orientation) === "LGBTQ+")],
    ],
  };
  const genderSummaries = Object.fromEntries(groups.gender.map(([label, rows]) => [label, buildSummary(rows)]));
  const raceSummaries = Object.fromEntries(groups.race.map(([label, rows]) => [label, buildSummary(rows)]));
  const orientSummaries = Object.fromEntries(groups.orientation.map(([label, rows]) => [label, buildSummary(rows)]));
  const cap = "Job Barriers and Job Loss by Demographic Group";
  const equityRows = [
    { _header: true,  "Demographic Group": "Demographic Group",    "N": "N",   "Q10 Had Barrier (%)": "Q10 Had Barrier (%)", "Q11 Left Job (%)": "Q11 Left Job (%)" },
    {                 "Demographic Group": "Overall",              "N": String(overall.n), "Q10 Had Barrier (%)": overall.barrierPct, "Q11 Left Job (%)": overall.leftPct },
    { _header: true,  "Demographic Group": "Gender",               "N": "",    "Q10 Had Barrier (%)": "",                    "Q11 Left Job (%)": "" },
    {                 "Demographic Group": "Female",               "N": String(genderSummaries["Female"].n),  "Q10 Had Barrier (%)": genderSummaries["Female"].barrierPct,                 "Q11 Left Job (%)": genderSummaries["Female"].leftPct },
    {                 "Demographic Group": "Male",                 "N": String(genderSummaries["Male"].n),  "Q10 Had Barrier (%)": genderSummaries["Male"].barrierPct,                 "Q11 Left Job (%)": genderSummaries["Male"].leftPct },
    {                 "Demographic Group": "Trans/Non-binary/GNC", "N": String(genderSummaries["Trans/Non-binary/GNC"].n),  "Q10 Had Barrier (%)": genderSummaries["Trans/Non-binary/GNC"].barrierPct,                 "Q11 Left Job (%)": genderSummaries["Trans/Non-binary/GNC"].leftPct },
    { _header: true,  "Demographic Group": "Race/Ethnicity",       "N": "",    "Q10 Had Barrier (%)": "",                    "Q11 Left Job (%)": "" },
    {                 "Demographic Group": "Black",                "N": String(raceSummaries["Black"].n),  "Q10 Had Barrier (%)": raceSummaries["Black"].barrierPct,                 "Q11 Left Job (%)": raceSummaries["Black"].leftPct },
    {                 "Demographic Group": "White",                "N": String(raceSummaries["White"].n),  "Q10 Had Barrier (%)": raceSummaries["White"].barrierPct,                 "Q11 Left Job (%)": raceSummaries["White"].leftPct },
    {                 "Demographic Group": "Multiracial",          "N": String(raceSummaries["Multiracial"].n),  "Q10 Had Barrier (%)": raceSummaries["Multiracial"].barrierPct,                 "Q11 Left Job (%)": raceSummaries["Multiracial"].leftPct },
    { _header: true,  "Demographic Group": "Sexual Orientation",   "N": "",    "Q10 Had Barrier (%)": "",                    "Q11 Left Job (%)": "" },
    {                 "Demographic Group": "Heterosexual/Straight", "N": String(orientSummaries["Heterosexual/Straight"].n), "Q10 Had Barrier (%)": orientSummaries["Heterosexual/Straight"].barrierPct,                 "Q11 Left Job (%)": orientSummaries["Heterosexual/Straight"].leftPct },
    {                 "Demographic Group": "LGBTQ+",               "N": String(orientSummaries["LGBTQ+"].n),  "Q10 Had Barrier (%)": orientSummaries["LGBTQ+"].barrierPct,                 "Q11 Left Job (%)": orientSummaries["LGBTQ+"].leftPct },
  ];
  return [
    makeHeading("Employment Barriers and Job Loss by Demographic Group", 2),
    makePara(
      "Examining job barriers and job loss by gender, race/ethnicity, and sexual orientation reveals that overall rates are broadly consistent across groups, though one specific barrier shows wider variation. " +
      `Because group sizes vary considerably, rates for smaller groups \u2014 particularly Trans/Non-binary/GNC respondents (n=${genderSummaries["Trans/Non-binary/GNC"].n}) \u2014 should be interpreted with caution.`
    ),
    makePara(
      `The share of youth reporting any job barrier ranged from ${genderSummaries["Male"].barrierPct} to ${genderSummaries["Trans/Non-binary/GNC"].barrierPct} across the demographic groups shown, with no group standing out for overall barrier prevalence. ` +
      `Rates of leaving a job were somewhat elevated for Black respondents (${raceSummaries["Black"].leftPct}) and LGBTQ+ respondents (${orientSummaries["LGBTQ+"].leftPct}) compared to the overall rate of ${overall.leftPct}. ` +
      `The clearest disparity was in mental or physical health cited as a barrier: reported by ${genderSummaries["Trans/Non-binary/GNC"].mentalPct} of Trans/Non-binary/GNC youth and ${raceSummaries["White"].mentalPct} of White youth, ` +
      `compared to ${raceSummaries["Black"].mentalPct} of Black youth and ${genderSummaries["Male"].mentalPct} of male youth. Full per-barrier and per-reason breakouts by demographic group are available in the analysis workbook (sheets q10_barriers_equity and q11_reasons_equity).`
    ),
    makeCaption(cap),
    makeTable(equityRows, "Overall", cap),
    makePara(
      "Note: Percentages are shares of each demographic group. Race uses a counted-once approach (respondents identifying with two or more groups are counted as Multiracial). " +
      `Respondents who did not report gender, race, or orientation are excluded from their respective group comparisons. Trans/Non-binary/GNC n=${genderSummaries["Trans/Non-binary/GNC"].n}; interpret rates for this group with caution.`,
      { italic: true }
    ),
    makePara(""),
  ];
}

function sec_transportation() {
  const subs = splitSheet("14_transport");
  const cap1 = "Driver\u2019s License Status by Age";
  const cap2 = "Drivers\u2019 Access to a Reliable Vehicle by Age";
  const cap3 = "Primary Way Youth Get to Work";
  return [
    makeHeading("TRANSPORTATION", 1),
    makePara(
      "Youth were also asked about whether they have a driver\u2019s license and the type " +
      "of transportation they rely on for work. Overall, 31% of respondents indicated they " +
      "have a driver\u2019s license, and an additional 9% have their learner\u2019s permit. " +
      "The likelihood a young person has their license increases with age: 17% of youth ages " +
      "16\u201317 have a license compared to 23% of those ages 18\u201320 and 44% of those ages " +
      "21\u201323. Compared to March 2025, license-holding increased slightly from 30% to 31%, " +
      "while the learner\u2019s-permit share remained at 9%."
    ),
    makeCaption(cap1),
    makeTable(subs["Driver's License by Age"] || [], "Total", cap1),
    makePara(
      "Of those with a driver\u2019s license, 43% regularly have access to a reliable vehicle. " +
      "About one-third have their own reliable vehicle, and another 10% share a reliable vehicle " +
      "or can borrow one when needed. By comparison, 44% report they do not usually have access " +
      "to a reliable vehicle, a slight worsening from March 2025 when 49% of licensed youth " +
      "reported regular reliable vehicle access."
    ),
    makeCaption(cap2),
    makeTable(subs["Vehicle Access (licensed)"] || [], "Total", cap2),
    makePara(
      "All youth were asked about the primary way they get to work when they are employed. " +
      "The majority rely on bus or other public transportation, reported by 68% of employed youth. " +
      "Another 11% indicated they use a combination of non-bus, non-driving methods, while 15% " +
      "report driving themselves to work. Reliance on public transportation increased from 60% " +
      "in March 2025, and the share driving themselves rose from 9%."
    ),
    makeCaption(cap3),
    makeTable(subs["Primary Transport"] || [], "Total", cap3),
    makePara(""),
  ];
}

function sec_voter_reg(dfCsv) {
  const subs      = splitSheet("15_voter_reg");
  const dfReg     = subs["Voter Registration by Age"]       || [];
  const dfReasons = subs["Not Registered Reasons by Age"]   || [];
  const dfGender  = subs["Registration by Gender (18-23)"]  || [];
  const dfOrient  = subs["Registration by Sexual Orientation (18-23)"] || [];
  let totalPctReg = "[PLACEHOLDER]";
  let pct18to20   = "[PLACEHOLDER]";
  let pct21to23   = "[PLACEHOLDER]";
  const regData = dfReg.filter((r) => !r._header);
  const regRow  = regData.find((r) => firstCol(r) === "Registered to Vote");
  if (regRow && regRow["Total"]) totalPctReg = regRow["Total"];
  if (regRow && regRow["18-20 years old"]) pct18to20 = regRow["18-20 years old"];
  if (regRow && regRow["21-23 years old"]) pct21to23 = regRow["21-23 years old"];

  const genderData = dfGender.filter((r) => !r._header);
  const femaleRow = genderData.find((r) => firstCol(r) === "Female");
  const maleRow = genderData.find((r) => firstCol(r) === "Male");
  const transRow = genderData.find((r) => firstCol(r) === "Trans, Non-binary");
  const pctFemaleReg = femaleRow ? (femaleRow["Percent Registered"] || "[PLACEHOLDER]") : "[PLACEHOLDER]";
  const pctMaleReg = maleRow ? (maleRow["Percent Registered"] || "[PLACEHOLDER]") : "[PLACEHOLDER]";
  const pctTransReg = transRow ? (transRow["Percent Registered"] || "[PLACEHOLDER]") : "[PLACEHOLDER]";

  const orientData = dfOrient.filter((r) => !r._header);
  const heteroRow = orientData.find((r) => firstCol(r) === "Heterosexual/Straight");
  const lgbtqRow = orientData.find((r) => firstCol(r) === "LGBTQ+");
  const pctHeteroReg = heteroRow ? (heteroRow["Percent Registered"] || "[PLACEHOLDER]") : "[PLACEHOLDER]";
  const pctLgbtqReg = lgbtqRow ? (lgbtqRow["Percent Registered"] || "[PLACEHOLDER]") : "[PLACEHOLDER]";

  const eligibleRows = dfCsv.filter((row) => {
    const age = String(row.age_range || "").trim();
    const reg = String(row.q7_registered_to_vote || "").trim();
    return ["18_20", "21_23"].includes(age) && ["yes", "no"].includes(reg);
  });
  const reasonRows = eligibleRows.filter((row) => String(row.q7a_not_registered_reasons || "").trim() !== "");
  const dontKnowHowCount = reasonRows.filter((row) =>
    String(row.q7a_not_registered_reasons || "").split("|").map((part) => part.trim()).includes("dont_know_how")
  ).length;
  const pctDontKnowHow = pct(dontKnowHowCount, reasonRows.length) || "[PLACEHOLDER]";
  const inconsistentRegistered = eligibleRows.filter((row) =>
    row.q7_registered_to_vote === "yes" && String(row.q7a_not_registered_reasons || "").trim() !== ""
  ).length;

  const cap1 = "Self-Reported Voter Registration by Age";
  const cap2 = "Reasons Youth Report Not Registering to Vote";
  return [
    makeHeading("Voting", 1),
    makePara(
      "Youth ages 18 and older were asked whether they are registered to vote; overall, " +
      `${totalPctReg} of eligible respondents reported being registered, down from 67% in March 2025.`
    ),
    makePara(
      `Older youth were more likely to be registered than youth ages 18 to 20 (${pct21to23} of 21- to 23-year-olds, compared with ${pct18to20} of 18- to 20-year-olds). Registration rates also varied by gender, with ${pctFemaleReg} of female youth registered compared with ${pctMaleReg} of male youth and ${pctTransReg} of transgender and non-binary youth. By sexual orientation, registration rates were similar at ${pctLgbtqReg} for LGBTQ+ youth and ${pctHeteroReg} for heterosexual youth.`
    ),
    makeCaption(cap1),
    makeTable(dfReg, "__none__", cap1),
    makePara(
      "As in March 2025, the most common reason youth gave for not registering to vote was believing their vote would not make a difference, followed by not understanding politics."
    ),
    makePara(
      `${dontKnowHowCount} of the ${reasonRows.length} youth who provided a reason (${pctDontKnowHow}) said they did not know how to register, and ${inconsistentRegistered === 0 ? "unlike March 2025, no respondents in the current data reported being registered while also selecting a reason for not being registered" : `${inconsistentRegistered} respondents reported being registered while also selecting a reason for not being registered`}.`
    ),
    makeCaption(cap2),
    makeTable(dfReasons, "__none__", cap2),
    makePara(""),
  ];
}

async function sec_zone_visit() {
  const subs      = splitSheet("16_visit");
  const dfFreq    = subs["Visit Frequency by Age"]       || [];
  const dfReasons = subs["Visit Reasons (frequent)"]     || [];
  const dfBarriers= subs["Visit Barriers (infrequent)"]  || [];
  const dfFocusSupport = subs["Stay Focused Narrative Support"] || [];
  const freqData = dfFreq.filter((r) => !r._header);
  const reasonData = dfReasons.filter((r) => !r._header);
  const barrierData = dfBarriers.filter((r) => !r._header);
  const focusSupportData = dfFocusSupport.filter((r) => !r._header);
  const num = (row, col) => Number(String(row?.[col] ?? "").replace(/[^0-9.-]/g, "")) || 0;

  const totalVisitRow = freqData.find((r) => firstCol(r) === "Total");
  const weeklyRow = freqData.find((r) => firstCol(r) === "Every week");
  const monthlyRow = freqData.find((r) => firstCol(r) === "1-3 times per month");
  const lessMonthlyRow = freqData.find((r) => firstCol(r) === "Less than once per month");
  const neverRow = freqData.find((r) => firstCol(r) === "Never");

  const totalVisits = num(totalVisitRow, "Total");
  const weeklyTotal = num(weeklyRow, "Total");
  const monthlyTotal = num(monthlyRow, "Total");
  const lessMonthlyTotal = num(lessMonthlyRow, "Total");
  const neverTotal = num(neverRow, "Total");
  const atLeastMonthlyTotal = weeklyTotal + monthlyTotal;

  const pctWeekly = pct(weeklyTotal, totalVisits) || "[PLACEHOLDER]";
  const pctMonthly = pct(monthlyTotal, totalVisits) || "[PLACEHOLDER]";
  const pctLessMonthly = pct(lessMonthlyTotal, totalVisits) || "[PLACEHOLDER]";
  const pctNever = pct(neverTotal, totalVisits) || "[PLACEHOLDER]";
  const pctWeekly16to17 = pct(num(weeklyRow, "16-17 years old"), num(totalVisitRow, "16-17 years old")) || "[PLACEHOLDER]";
  const pctWeekly18to20 = pct(num(weeklyRow, "18-20 years old"), num(totalVisitRow, "18-20 years old")) || "[PLACEHOLDER]";
  const pctWeekly21to23 = pct(num(weeklyRow, "21-23 years old"), num(totalVisitRow, "21-23 years old")) || "[PLACEHOLDER]";

  const seeCoachRow = reasonData.find((r) => firstCol(r) === "See my Youth Coach or other Zone Staff");
  const foodRow = reasonData.find((r) => firstCol(r) === "Eat food");
  const scheduledRow = reasonData.find((r) => firstCol(r) === "Participate in a scheduled activity");
  const goalsRow = reasonData.find((r) => firstCol(r) === "Work toward my goals");
  const safePlaceRow = reasonData.find((r) => firstCol(r) === "Be in a safe place");
  const escapeRow = reasonData.find((r) => firstCol(r) === "Escape problems/issues");
  const frequent18to20 = num(weeklyRow, "18-20 years old") + num(monthlyRow, "18-20 years old");
  const frequent21to23 = num(weeklyRow, "21-23 years old") + num(monthlyRow, "21-23 years old");

  const pctSeeCoach = pct(num(seeCoachRow, "Total"), atLeastMonthlyTotal) || "[PLACEHOLDER]";
  const pctFood = pct(num(foodRow, "Total"), atLeastMonthlyTotal) || "[PLACEHOLDER]";
  const pctScheduled = pct(num(scheduledRow, "Total"), atLeastMonthlyTotal) || "[PLACEHOLDER]";
  const pctGoals = pct(num(goalsRow, "Total"), atLeastMonthlyTotal) || "[PLACEHOLDER]";
  const pctSafe21to23 = pct(num(safePlaceRow, "21-23 years old"), frequent21to23) || "[PLACEHOLDER]";
  const pctEscape21to23 = pct(num(escapeRow, "21-23 years old"), frequent21to23) || "[PLACEHOLDER]";
  const pctSafe18to20 = pct(num(safePlaceRow, "18-20 years old"), frequent18to20) || "[PLACEHOLDER]";
  const pctEscape18to20 = pct(num(escapeRow, "18-20 years old"), frequent18to20) || "[PLACEHOLDER]";

  const barrierRespondentsRow = barrierData.find((r) => firstCol(r) === "Total respondents");
  const activitiesRow = barrierData.find((r) => firstCol(r) === "More activities that interest me");
  const inviteRow = barrierData.find((r) => firstCol(r) === "Invitation from my Youth Coach");
  const infoRow = barrierData.find((r) => firstCol(r) === "Knowing more about activities");
  const barrierTotal = num(barrierRespondentsRow, "Total");
  const pctActivities = pct(num(activitiesRow, "Total"), barrierTotal) || "[PLACEHOLDER]";
  const pctInvite = pct(num(inviteRow, "Total"), barrierTotal) || "[PLACEHOLDER]";
  const pctInfo = pct(num(infoRow, "Total"), barrierTotal) || "[PLACEHOLDER]";
  const invite18to20 = num(inviteRow, "18-20 years old");
  const infrequent18to20 = num(barrierRespondentsRow, "18-20 years old");

  const overallAgreeRow = focusSupportData.find((r) => firstCol(r) === "Overall fully agree Youth Zone helps stay focused");
  const monthlyAgreeRow = focusSupportData.find((r) => firstCol(r) === "At least monthly visitors agree or somewhat agree Youth Zone helps stay focused");
  const infrequentAgreeRow = focusSupportData.find((r) => firstCol(r) === "Less than monthly or never visitors agree or somewhat agree Youth Zone helps stay focused");
  const monthlyUnsureRow = focusSupportData.find((r) => firstCol(r) === "At least monthly visitors unsure about goals");
  const infrequentUnsureRow = focusSupportData.find((r) => firstCol(r) === "Less than monthly or never visitors unsure about goals");
  const pctOverallAgree = String(overallAgreeRow?.Percent || "[PLACEHOLDER]");
  const pctMonthlyAgree = String(monthlyAgreeRow?.Percent || "[PLACEHOLDER]");
  const pctInfrequentAgree = String(infrequentAgreeRow?.Percent || "[PLACEHOLDER]");
  const pctMonthlyUnsure = String(monthlyUnsureRow?.Percent || "[PLACEHOLDER]");
  const pctInfrequentUnsure = String(infrequentUnsureRow?.Percent || "[PLACEHOLDER]");

  const cap1 = "Visit Frequency by Age";
  const cap2 = "What Are the Main Reasons Youth Come to the Youth Zone?";
  const cap3 = "What Would Make Someone Who Rarely Visits the Zone Want to Come, by Age";
  return [
    makeHeading("Zone Experience", 1),
    makePara(
      "Attendance patterns at the Zone vary by age, with older youth coming more frequently than youth ages 16 to 17. " +
      `Overall, ${pctWeekly} reported attending every week and ${pctMonthly} one to three times per month, while ${pctLessMonthly} reported visiting less than once per month and ${pctNever} indicated they never visit the downtown Zone. ` +
      `This represents a modest decrease in regular attendance from March 2025, when 52% of respondents reported coming at least monthly; no 16- to 17-year-olds reported weekly attendance, compared with ${pctWeekly18to20} of youth ages 18 to 20 and ${pctWeekly21to23} of youth ages 21 to 23.`
    ),
    makeCaption(cap1),
    makeTable(dfFreq, "Total", cap1),
    await embedChart("chart_04_visit_frequency.png", 5.5),
    makePara(
      `${pctOverallAgree} of respondents fully agreed that support from the Youth Zone helps them stay focused on their goals, up from 71% in February 2025. ` +
      `As in the prior report, there appears to be a mild correlation between attendance at the Zone and youth reporting that the Youth Zone helps them stay focused on their goals: ${pctMonthlyAgree} of youth who report attending the Zone at least monthly agree or somewhat agree, compared to ${pctInfrequentAgree} of youth who report attending less than monthly or never. ` +
      `Youth who attend less frequently are also more likely to report not having clear goals right now (${pctInfrequentUnsure} vs. ${pctMonthlyUnsure}).`
    ),
    await embedChart("chart_10_stay_focused_visit_frequency.png", 6.6),
    makePara(
      `As in prior years, most youth who come at least monthly report coming to the 412 Youth Zone downtown to see their Youth Coach or other Zone staff (${pctSeeCoach}). ` +
      `Food (${pctFood}) and scheduled activities (${pctScheduled}) were also among the most common reasons for coming this year, ahead of working toward goals (${pctGoals}). ` +
      `Among frequent visitors ages 21 to 23, ${pctSafe21to23} reported coming to be in a safe place and the same share reported coming to escape problems or issues, compared with ${pctSafe18to20} and ${pctEscape18to20}, respectively, of frequent visitors ages 18 to 20.`
    ),
    await embedChart("chart_11_visit_reasons_combo.png", 7.1),
    makePara(
      `For youth who never visit the Zone, or visit less than monthly, the most common response for what would make them want to come more frequently was more activities that interest them (${pctActivities}), followed closely by an invitation from their Youth Coach (${pctInvite}) and knowing more about activities (${pctInfo}). ` +
      `An invitation from a coach was especially salient for 18- to 20-year-olds, selected by ${invite18to20} of ${infrequent18to20} infrequent visitors in that age group. Open-text \"other\" responses most often referenced work or school schedules, transportation, and distance from the Zone.`
    ),
    makeCaption(cap3),
    makeTable(dfBarriers, "__none__", cap3),
    makePara(""),
  ];
}

async function sec_program_impact(dfCsv) {
  const helpedAny = dfCsv.filter((r) => (r.q17_program_helped || "").trim() !== "").length;
  const pctHelped = pct(helpedAny, dfCsv.length);
  const helpedTwoPlus = dfCsv.filter((r) => {
    const parts = String(r.q17_program_helped || "").split("|").map((t) => t.trim()).filter(Boolean);
    return parts.length >= 2;
  }).length;
  const pctHelpedTwoPlus = pct(helpedTwoPlus, dfCsv.length);
  const noQ17 = dfCsv.filter((r) => (r.q17_program_helped || "").trim() === "");
  const limitedContactCount = noQ17.filter((r) => {
    const combined = `${r.q17_none_explain_text || ""} ${r.q23_other_comments || ""}`.toLowerCase();
    return [
      "just joined",
      "first time",
      "not much opportunity",
      "don't know my youth coach",
      "dont know my youth coach",
    ].some((term) => combined.includes(term));
  }).length;

  const subs = splitSheet("17_impact");
  let dfQ17  = subs["Program Helped With (Q17) by A"] || [];
  const chartRows = (subs["Program Helped With (Q17) Chart Reference"] || []).filter((r) => !r._header);
  if (!dfQ17.length) {
    const key = Object.keys(subs).find((k) => k.startsWith("Program Helped"));
    if (key) dfQ17 = subs[key];
  }
  const chartPct = (label, col = "Total") => {
    const row = chartRows.find((r) => firstCol(r) === label);
    return row?.[col] || "[PLACEHOLDER]";
  };
  const pctFutureTotal = chartPct("Think about my future");
  const pctProblemsTotal = chartPct("Figure out how to handle problems");
  const pctVitalDocsTotal = chartPct("Obtain vital documents");
  const pctDecisionTotal = chartPct("Make good decisions");
  const pctHousingTotal = chartPct("Find or maintain housing");
  const pctRelationshipsTotal = chartPct("Establish positive relationships");
  const pctHousing16to17 = chartPct("Find or maintain housing", "16-17 years old");
  const pctHousing18to20 = chartPct("Find or maintain housing", "18-20 years old");
  const pctHousing21to23 = chartPct("Find or maintain housing", "21-23 years old");
  const pctVitalDocs16to17 = chartPct("Obtain vital documents", "16-17 years old");
  const pctLicense16to17 = chartPct("Get my driver's license", "16-17 years old");
  const items = [
    makeHeading("Impact of Assistance", 1),
    makePara(
      "Across the core outcome areas in which Youth Zone staff are helping young " +
      "people make progress, youth most often reported that their coaches and the Zone helped them think about their future. " +
      `Among respondents with known ages who answered this question, ${pctFutureTotal} selected that item, followed by figuring out how to handle problems (${pctProblemsTotal}). About four in ten also reported help obtaining vital documents (${pctVitalDocsTotal}), making good decisions (${pctDecisionTotal}), finding or maintaining housing (${pctHousingTotal}), and establishing positive relationships (${pctRelationshipsTotal}).`
    ),
    makePara(
      "Coaches also frequently provided concrete assistance by helping youth obtain vital documents, secure driver’s licenses, access health care or counseling, and find or maintain housing. These supports were generally reported more often by older youth, particularly help with housing (" +
      `${pctHousing21to23} of respondents ages 21 to 23, compared with ${pctHousing18to20} of ages 18 to 20 and ${pctHousing16to17} of ages 16 to 17), while younger youth were more likely to report help obtaining vital documents or getting a driver’s license (${pctVitalDocs16to17} and ${pctLicense16to17}, respectively, among ages 16 to 17).`
    ),
    makePara(
      `${pctHelped} of respondents indicated progress supported by the Zone in at least one area, and ${pctHelpedTwoPlus} reported support in two or more areas. Of the ${noQ17.length} youth who did not select an area in Q17, ${limitedContactCount} wrote comments suggesting they had just joined the program or had limited opportunity to connect with their coach.`
    ),
  ];
  if (dfQ17.length) {
    items.push(await embedChart("chart_12_program_helped_combo.png", 7.4));
  }
  items.push(makePara(""));
  return items;
}

async function sec_respect_environment() {
  const subs18 = splitSheet("18_respect");
  const rows18 = (subs18["Respect Summary"] || loadSheet("18_respect")).filter((r) => !r._header);
  const support18 = (subs18["Respect Narrative Support"] || []).filter((r) => !r._header);
  let pctStaff = "[PLACEHOLDER]", pctPeer = "[PLACEHOLDER]";
  let staffN = "[PLACEHOLDER]", peerN = "[PLACEHOLDER]";
  if (rows18.length && rows18[0]["% Often or All the Time"] !== undefined) {
    const staffRow = rows18.find((r) => firstCol(r).includes("Staff"));
    const peerRow  = rows18.find((r) => firstCol(r).includes("Peer"));
    if (staffRow) {
      pctStaff = staffRow["% Often or All the Time"] || "[PLACEHOLDER]";
      staffN = String(staffRow.n || "[PLACEHOLDER]");
    }
    if (peerRow) {
      pctPeer  = peerRow["% Often or All the Time"]  || "[PLACEHOLDER]";
      peerN = String(peerRow.n || "[PLACEHOLDER]");
    }
  }
  const majorityRow = support18.find((r) => firstCol(r) === "Known-age rarely/never group ages 21-23");
  const under18Row = support18.find((r) => firstCol(r) === "Known-age rarely/never group younger than 18");
  const pctMajority21to23 = String(majorityRow?.Percent || "[PLACEHOLDER]");
  const pctUnder18RareNever = String(under18Row?.Percent || "[PLACEHOLDER]");
  const pct = (n, d) => {
    if (!d) return "";
    return `${Math.round((100 * n) / d)}%`;
  };
  const countNum = (row, key) => Number(String(row?.[key] ?? "").replace(/[^0-9.-]/g, "")) || 0;
  const staffRow = rows18.find((r) => firstCol(r).includes("Staff"));
  const peerRow  = rows18.find((r) => firstCol(r).includes("Peer"));
  const tinyNotes = [];
  const maybeAddTiny = (label, row, nVal) => {
    const denom = countNum({ n: nVal }, "n");
    const rarelyPct = pct(countNum(row, "Rarely"), denom);
    const neverPct = pct(countNum(row, "Never"), denom);
    if (countNum(row, "Rarely") > 0 && (100 * countNum(row, "Rarely") / denom) < 3) tinyNotes.push(`${label} rarely ${rarelyPct}`);
    if (countNum(row, "Never") > 0 && (100 * countNum(row, "Never") / denom) < 3) tinyNotes.push(`${label} never ${neverPct}`);
  };
  if (staffRow) maybeAddTiny("Staff", staffRow, staffN);
  if (peerRow) maybeAddTiny("Peers", peerRow, peerN);
  const respectCaption = `Responses: n=${staffN} for staff and n=${peerN} for peers.${tinyNotes.length ? ` Labels under 3% omitted: ${tinyNotes.join("; ")}.` : ""}`;
  const rows19 = loadSheet("19_environment");
  const has19  = rows19.filter((r) => !r._header).length > 0;
  const items = [
    makePara(
      "Respondents were asked to rate how often they felt respected and accepted " +
      `for who they are at the Youth Zone. ${pctStaff} of youth reported staff ` +
      `treat them with respect often or all the time; ${pctPeer} said the same ` +
      "about their peers at the Zone."
    ),
    makePara(
      `Although the number of youth reporting that they rarely or never felt respected and accepted was small, some disparities were apparent beyond age. Among youth who responded to the respect questions, 20% of Trans/Non-binary/GNC youth reported rarely or never feeling respected and accepted by staff or peers, compared with 9% of female youth and 3% of male youth. By sexual orientation, 11% of LGBTQ+ youth reported rarely or never feeling respected and accepted, compared with 7% of heterosexual youth. These patterns are consistent with the disparities noted in the Employment Barriers and Job Loss by Demographic Group section, where Trans/Non-binary/GNC youth also stood out on selected indicators, though all subgroup findings should be interpreted with caution given the small number of respondents.`
    ),
    await embedChart("chart_13_respect_acceptance.png", 6.6),
    makeCaption(respectCaption),
  ];
  if (has19) {
    items.push(makePara(
      "Youth also rated five statements about the program environment on a " +
      "1\u20135 scale. Results are shown in terms of the percentage selecting 4 or 5 (top-2 box)."
    ));
    items.push(makePara(
      "Responses to these items were broadly positive. The strongest ratings were for being treated fairly (85%) and for diversity of backgrounds being valued (84%), while feeling that people around them care about their success was the lowest-rated item, though it was still selected by nearly three-quarters of respondents (74%). Overall, each of the five environment measures received top-2 ratings from at least 74% of youth who answered that item."
    ));
    items.push(await embedChart("chart_14_environment_ratings.png", 7.1));
  }
  items.push(makePara(""));
  return items;
}

function sec_banking(dfCsv) {
  const subs = splitSheet("20_banking");
  const bankRows = (subs["Bank Account Status by Age"] || []).filter((r) => !r._header);
  const methodRows = (subs["Money Methods by Age (Q24)"] || []).filter((r) => !r._header);
  const usageRows = (subs["Account Usage by Age (Q26b)"] || []).filter((r) => !r._header);
  const hasAcctRow = bankRows.find((r) => firstCol(r).toLowerCase().includes("currently have"));
  const checkingRow = bankRows.find((r) => firstCol(r).toLowerCase().includes("checking account"));
  const savingsRow = bankRows.find((r) => firstCol(r).toLowerCase().includes("savings account"));
  const digitalAppsRow = methodRows.find((r) => firstCol(r) === "Venmo, Zelle, CashApp, etc.");
  const bankMethodRow = methodRows.find((r) => firstCol(r) === "Bank account");
  const cashHomeRow = methodRows.find((r) => firstCol(r) === "Saving/storing cash at home");
  const savingMoneyRow = usageRows.find((r) => firstCol(r) === "Saving money");
  const directDepositRow = usageRows.find((r) => firstCol(r) === "Direct deposit");
  const billsRow = usageRows.find((r) => firstCol(r) === "Paying household bills");
  const pctHas = (hasAcctRow && hasAcctRow["Percent of Total"])
    ? hasAcctRow["Percent of Total"]
    : pct(dfCsv.filter((r) => {
        const parts = (r.q25_bank_account || "").split("|");
        return parts.some((t) => ["checking", "savings"].includes(t.trim()));
      }).length, dfCsv.length);
  const pctCheckingAll = checkingRow?.["Percent of Total"] || "[PLACEHOLDER]";
  const pctSavingsAll = savingsRow?.["Percent of Total"] || "[PLACEHOLDER]";
  const pctApps = digitalAppsRow?.["Percent of All"] || "[PLACEHOLDER]";
  const pctBankMethod = bankMethodRow?.["Percent of All"] || "[PLACEHOLDER]";
  const pctCashHome = cashHomeRow?.["Percent of All"] || "[PLACEHOLDER]";
  const pctSavingMoney = savingMoneyRow?.["% of Account Holders"] || "[PLACEHOLDER]";
  const pctDirectDeposit = directDepositRow?.["% of Account Holders"] || "[PLACEHOLDER]";
  const pctBillsOlder = (() => {
    const num = Number(String(billsRow?.["21-23 years old"] ?? "").replace(/[^0-9.-]/g, "")) || 0;
    const den = Number(String(hasAcctRow?.["21-23 years old"] ?? "").replace(/[^0-9.-]/g, "")) || 0;
    return pct(num, den) || "[PLACEHOLDER]";
  })();
  const acctHolders = dfCsv.filter((r) => {
    const parts = String(r.q25_bank_account || "").split("|").map((t) => t.trim());
    return parts.some((t) => ["checking", "savings"].includes(t));
  });
  const acctNoBankMethod = acctHolders.filter((r) => {
    const parts = String(r.q24_money_methods || "").split("|").map((t) => t.trim());
    return !parts.includes("bank_account");
  });
  const acctNoBankMethodDigitalApps = acctNoBankMethod.filter((r) => {
    const parts = String(r.q24_money_methods || "").split("|").map((t) => t.trim());
    return parts.includes("digital_apps");
  });
  const pctAppsAmongNonBankUsers = pct(acctNoBankMethodDigitalApps.length, acctNoBankMethod.length) || "[PLACEHOLDER]";
  const cap1 = "Banking Status by Age";
  const cap2 = "Methods Youth Use to Store, Receive, and Transfer Money, by Age";
  const cap3 = "Ways Respondents Use Their Bank Account(s), by Age";
  return [
    makeHeading("Banking", 1),
    makePara(
      "Participants were asked questions about their use of banks and other ways " +
      "that they store, receive, and transfer money. Banking practices vary by age, with older youth more likely than 16- to 17-year-olds to report having an account. " +
      `Overall, ${pctHas} of respondents reported currently having a bank account, down slightly from 69% in March 2025.`
    ),
    makePara(
      `Nearly all youth who reported having an account also indicated they have a checking account (${pctCheckingAll} of all respondents), and ${pctSavingsAll} reported having a savings account. Digital payment apps such as Venmo, Zelle, and CashApp (${pctApps}) were used slightly more often than bank accounts themselves (${pctBankMethod}) to store, receive, and transfer money, while ${pctCashHome} reported keeping cash at home.`
    ),
    makePara(
      `Among youth with accounts, the most common uses were saving money (${pctSavingMoney}) and direct deposit (${pctDirectDeposit}). Older account holders were also more likely to report using their accounts to pay household bills (${pctBillsOlder} of account holders ages 21 to 23), and ${pctAppsAmongNonBankUsers} of youth who had an account but did not report using it to manage money said they rely on digital payment apps instead.`
    ),
    makeCaption(cap1),
    makeBankingTable(subs["Bank Account Status by Age"] || [], cap1, {
      highlightLabel: "Currently have a bank account",
      italicLabels: new Set(["Checking account", "Savings account"]),
    }),
    makeCaption(cap2),
    makeBankingTable(subs["Money Methods by Age (Q24)"] || [], cap2, {
      highlightLabel: "Bank account",
    }),
    makeCaption(cap3),
    makeTable(subs["Account Usage by Age (Q26b)"]    || [], "__none__", cap3),
    makePara(""),
  ];
}

async function sec_nps() {
  const rows21   = loadSheet("21_nps");
  const data21   = rows21.filter((r) => !r._header);
  const npsRow   = data21.find((r) => firstCol(r) === "NPS Score");
  const npsScore = (npsRow && npsRow["Count"])   ? npsRow["Count"]   : "[PLACEHOLDER]";
  const promRow  = data21.find((r) => firstCol(r).includes("Promoter"));
  const pctProm  = (promRow && promRow["Percent"]) ? promRow["Percent"] : "[PLACEHOLDER]";
  return [
    makePara(
      `Youth were asked to rate on a scale of 0\u201310 how likely they would be to ` +
      "recommend the Youth Zone to a friend or family member. " +
      `The Net Promoter Score (NPS) is ${npsScore}. ` +
      `${pctProm} of respondents were Promoters (9\u201310).`
    ),
    await embedChart("chart_05_nps.png", 5),
    makePara(""),
  ];
}

function sec_comments() {
  const subs = splitSheet("22_comments");
  const summaryRows = (subs["Comment Summary"] || []).filter((r) => !r._header);
  const commentRows = (subs["Comment Listing"] || []).filter((r) => !r._header);
  const summaryCount = (label) => {
    const row = summaryRows.find((r) => firstCol(r) === label);
    return row ? String(row["Count"] ?? "") : "";
  };
  const nCom = summaryCount("Total non-blank comments") || String(commentRows.length);
  const substantiveCount = summaryCount("Substantive comments used for narrative review") || "[PLACEHOLDER]";
  const commentKey = (() => {
    const cols = Object.keys(commentRows[0] || {}).filter((k) => k !== "_header");
    return cols.includes("Comment") ? "Comment" : cols[cols.length - 1];
  })();
  const substantive = commentRows
    .map((row) => String(row[commentKey] ?? "").trim())
    .filter((text) => text && text !== "nan")
    .filter((text) => ![
      "no",
      "nope",
      "none.",
      "none",
      "n/a",
      "na",
      "not at the moment",
      "not at this time.",
      "no comment all good",
      "nah i'm good thxs",
      "a/k",
      "no.",
      "not at this time",
      "not at the moment.",
    ].includes(text.toLowerCase()));
  const items = [
    makeHeading("Additional Comments", 1),
    makePara(
      `Finally, youth had the option to share any other comments or feedback ` +
      `they had about the Zone. ${nCom} youth provided additional comments.`
    ),
    makePara(
      `As in March 2025, most substantive comments were positive and focused on appreciation for Youth Zone staff, coaches, and the support youth receive through the program. Of the ${substantiveCount} comments that provided more than a brief no/none response, many described staff as welcoming, helpful, and attentive, and several youth specifically said the Zone helped them feel safe, supported, or make progress toward their goals.`
    ),
    makePara(
      "A smaller number of comments identified areas for improvement, most often around communication, access to resources such as driver's licensing help, or making activities easier to learn about in advance. A few quotes are included below as examples of the responses received."
    ),
  ];
  if (commentRows.length) {
    items.push(makeHeading("Comments", 1));
    for (const row of commentRows) {
      const text = String(row[commentKey] ?? "").trim();
      if (text && text !== "nan") items.push(makeBullet(text));
    }
  }
  return items;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  if (!fs.existsSync(ANALYSIS_PATH)) {
    console.error(`Analysis file not found: ${ANALYSIS_PATH}`);
    console.error("Run 04_analyze_412YZ.py first.");
    process.exit(1);
  }

  // Load table width config from JSON
  const twPath = path.join(__dirname, "table_widths_412YZ.json");
  if (fs.existsSync(twPath)) {
    const raw = JSON.parse(fs.readFileSync(twPath, "utf8"));
    for (const [k, v] of Object.entries(raw)) {
      if (k !== "_instructions") TABLE_WIDTHS[k] = v;
    }
    console.log(`Loaded ${Object.keys(TABLE_WIDTHS).length} fixed column-width entries.`);
  }

  console.log("Loading CSV data...");
  const dfCsv = loadCsv();
  N_RESPONDENTS = dfCsv.length;
  console.log(`N_RESPONDENTS set to ${N_RESPONDENTS} (from CSV row count)`);

  console.log("Building sections...");
  const children = [
    ...sec_title(),
    ...sec_age(),
    ...sec_gender_orient(),
    ...(await sec_race()),
    ...(await sec_coach_satisfaction_async()),
    ...(await sec_communication()),
    ...(await sec_housing()),
    ...(await sec_education_employment(dfCsv)),
    ...sec_job_tenure(dfCsv),
    ...(await sec_employment_by_age()),
    ...sec_job_barriers(dfCsv),
    ...sec_left_job(),
    ...sec_employment_equity(),
    ...sec_transportation(),
    ...sec_voter_reg(dfCsv),
    ...(await sec_zone_visit()),
    ...(await sec_program_impact(dfCsv)),
    ...(await sec_respect_environment()),
    ...sec_banking(dfCsv),
    ...(await sec_nps()),
    ...sec_comments(),
  ].filter((node) => !SPACER_PARAGRAPHS.has(node));

  console.log("Assembling document...");
  const doc = new Document({
    numbering: BULLET_NUMBERING,
    styles: DOC_STYLES,
    sections: [
      {
        properties: {
          page: {
            size:   { width: 12240, height: 15840 },
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
          },
        },
        children,
      },
    ],
  });

  const outDir = path.dirname(OUT_PATH);
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  console.log("Writing file...");
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUT_PATH, buffer);

  const stats = fs.statSync(OUT_PATH);
  console.log(`\nSaved: ${OUT_PATH}`);
  console.log(`File size: ${(stats.size / 1024).toFixed(1)} KB`);
}

main().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
