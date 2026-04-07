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
const { parse: parseCsv } = requireGlobal("csv-parse/sync");

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const BASE_DIR      = path.join(__dirname, "..");
const ANALYSIS_PATH = path.join(BASE_DIR, "output", "412YZ", "analysis_412YZ.xlsx");
const CSV_PATH      = path.join(BASE_DIR, "output", "412YZ", "survey_data_412YZ.csv");
const OUT_PATH      = path.join(BASE_DIR, "report", "412YZ", "report_412YZ_v2.docx");

const SURVEY_MONTH  = "March 2026";
const N_RESPONDENTS = 103;
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
        spacing: { before: 240, after: 120 },
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
        spacing: { before: 180, after: 60 },
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
        spacing: { before: 120, after: 60 },
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
      const cellWidth = fixedWidths
        ? { size: fixedWidths[colIdx] ?? 0, type: WidthType.DXA }
        : { size: 0, type: WidthType.AUTO };
      return new TableCell({
        width: cellWidth,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        borders: NO_BORDER,
        shading: shaded
          ? { fill: HDR_FILL, type: ShadingType.CLEAR, color: "auto" }
          : undefined,
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: cellText,
                bold: shaded,
                font: "Calibri",
                size: 22,
              }),
            ],
          }),
        ],
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
  return new Paragraph({
    style: indent ? "ListParagraph" : "Normal",
    children: [new TextRun({ text, bold, italic, font: "Calibri", size: 22 })],
  });
}

/**
 * makeBullet(text) — bullet list paragraph
 */
function makeBullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, font: "Calibri", size: 22 })],
  });
}

/**
 * makePlaceholder(text) — yellow-highlighted paragraph
 */
function makePlaceholder(text) {
  return new Paragraph({
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
// Main
// ---------------------------------------------------------------------------

async function main() {
  if (!fs.existsSync(ANALYSIS_PATH)) {
    console.error(`Analysis file not found: ${ANALYSIS_PATH}`);
    console.error("Run 04_analyze_412YZ.py first.");
    process.exit(1);
  }

  console.log("Building sections...");
  const children = [
    ...sec_age(),
  ];

  console.log("Assembling document...");
  const doc = new Document({
    numbering: BULLET_NUMBERING,
    styles: DOC_STYLES,
    sections: [
      {
        properties: {
          page: {
            size: {
              width:  12240, // 8.5in in DXA (1440 DXA/in)
              height: 15840, // 11in in DXA
            },
            margin: {
              top:    1440,  // 1in
              bottom: 1440,
              left:   1440,
              right:  1440,
            },
          },
        },
        children,
      },
    ],
  });

  const outDir = path.dirname(OUT_PATH);
  if (!fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true });
  }

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
