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
// Chart helper
// ---------------------------------------------------------------------------

const CHARTS_DIR = path.join(BASE_DIR, "output", "412YZ", "charts");

/**
 * embedChart(chartFilename, widthInches = 6)
 * Reads the PNG from output/412YZ/charts/, returns a centered Paragraph
 * containing an ImageRun. Width in EMUs = widthInches * 914400; height is
 * computed proportionally using sharp.
 */
async function embedChart(chartFilename, widthInches = 6) {
  const chartPath = path.join(CHARTS_DIR, chartFilename);
  if (!fs.existsSync(chartPath)) {
    console.warn(`  [warn] Chart not found: ${chartPath}`);
    return new Paragraph({ text: "" });
  }
  const data = fs.readFileSync(chartPath);

  let widthEmu  = Math.round(widthInches * 914400);
  let heightEmu = Math.round(widthEmu * 0.5625); // fallback 16:9

  if (sharp) {
    try {
      const meta  = await sharp(chartPath).metadata();
      const ratio = meta.height / meta.width;
      heightEmu   = Math.round(widthEmu * ratio);
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
        transformation: { width: widthEmu, height: heightEmu },
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
      "participate in a survey in early 2026. Surveys were administered on paper " +
      "and results were digitized for this report."
    ),
    new Paragraph({
      children: [
        new TextRun({ text: `${N_RESPONDENTS} unique youth (of `, font: "Calibri", size: 22 }),
        new TextRun({ text: "[TOTAL ACTIVE \u2014 fill in denominator]", highlight: "yellow", bold: true, font: "Calibri", size: 22 }),
        new TextRun({ text: " total active) responded to the survey, for a response rate of ", font: "Calibri", size: 22 }),
        new TextRun({ text: "[RESPONSE RATE %]", highlight: "yellow", bold: true, font: "Calibri", size: 22 }),
        new TextRun({ text: ". Most respondents were age 18\u201320 years old.", font: "Calibri", size: 22 }),
      ],
    }),
    makePara(""),
  ];
}

function sec_gender_orient() {
  const rows = loadSheet("02_gender_orient");
  const data = rows.filter((r) => !r._header);
  const numRow = data.find((r) => firstCol(r) === "Number of Youth");
  let nF = 0, nM = 0, nNB = 0;
  if (numRow) {
    nF  = parseInt(numRow["Female"]           || "0") || 0;
    nM  = parseInt(numRow["Male"]             || "0") || 0;
    nNB = parseInt(numRow["Trans, Non-binary"]|| "0") || 0;
  }
  const nKnown = nF + nM + nNB;
  const pctF = pct(nF, nKnown);
  const pctM = pct(nM, nKnown);
  const cap = "Survey Respondents by Gender and Sexual Orientation";
  return [
    makePara(
      `More females (${pctF}) responded to the survey than males (${pctM}). ` +
      "The table below shows how respondents identified by gender and sexual orientation."
    ),
    makeCaption(cap),
    makeTable(rows, "Total", cap),
    makePara(""),
  ];
}

async function sec_race() {
  const rowsOnce = loadSheet("03_race_once");
  const data = rowsOnce.filter((r) => !r._header);
  const findPct = (label) => {
    const row = data.find((r) => firstCol(r) === label);
    return row && row["Percent"] ? row["Percent"] : "[PLACEHOLDER]";
  };
  const pctBlack = findPct("Black");
  const pctWhite = findPct("White");
  const pctMulti = findPct("Multiracial");
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
      `About ${pctBlack} of survey respondents identified as Black, ` +
      `${pctWhite} as White, and ${pctMulti} as Multiracial.`
    ),
    makeCaption(cap1),
    makeTable(rowsOnce, "Total", cap1),
    await embedChart("chart_06_race_distribution.png", 5),
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
  const currCol   = "Mar-26";
  const currN     = `n=${N_RESPONDENTS}`;
  const allCols   = ["My Youth Coach\u2026", ...priorCols, currCol];

  const headerRow = { _header: true };
  allCols.forEach((c) => { headerRow[c] = c; });

  const row1 = {}; allCols.forEach((c, i) => { row1[c] = i === 0 ? "My Youth Coach\u2026" : "% Often or All the Time"; });
  const row2 = {}; allCols.forEach((c) => { row2[c] = c; });
  const row3 = {}; allCols.forEach((c, i) => {
    if (i === 0) row3[c] = "My Youth Coach\u2026";
    else if (i <= priorCols.length) row3[c] = Q1_BENCHMARKS.n[i - 1];
    else row3[c] = currN;
  });

  const FIELD_LABELS = [
    "Is trustworthy",
    "Is reliable",
    "Values my opinions about my life",
    "Is available to me when I need them",
    "Makes me feel heard and understood",
  ];
  const dataRowsQ1 = FIELD_LABELS.map((label) => {
    const obj = {};
    allCols.forEach((c, i) => {
      if (i === 0) obj[c] = label;
      else if (i <= priorCols.length) obj[c] = Q1_BENCHMARKS[label][i - 1];
      else obj[c] = currPct[label] || "";
    });
    return obj;
  });

  const cap = "Satisfaction Ratings for Youth Coaches Over Time";
  return { pctTrust, pctVals, headerRow, row1, row2, row3, dataRowsQ1, cap };
}

async function sec_coach_satisfaction_async() {
  const { pctTrust, pctVals, headerRow, row1, row2, row3, dataRowsQ1, cap } = sec_coach_satisfaction();
  return [
    makeHeading("FINDINGS", 1),
    makePara(""),
    makeHeading("Relationships with Coach", 2),
    makePara(
      `${pctTrust} of youth reported their coaches were trustworthy, and ` +
      `${pctVals} indicated their coach values their opinions about their life. ` +
      "Ratings across all five coach relationship items are shown in the table below."
    ),
    makeCaption(cap),
    makeTable([headerRow, row1, row2, row3, ...dataRowsQ1], "__none__", cap),
    await embedChart("chart_01_coach_satisfaction.png", 5.5),
    makePara(""),
  ];
}

function sec_communication() {
  const rows = loadSheet("06_communication");
  const data = rows.filter((r) => !r._header);
  const cols = Object.keys(rows[0] || {}).filter((k) => k !== "_header");
  let notEnoughCount = "", notEnoughPct = "";
  for (const row of data) {
    const first = String(row[cols[0]] ?? "").toLowerCase();
    if (first.includes("not enough")) {
      notEnoughCount = String(row[cols[1]] ?? "");
      notEnoughPct   = cols.length > 2 ? String(row[cols[2]] ?? "") : "";
    }
  }
  const text = (notEnoughCount && notEnoughPct)
    ? `The majority of youth communicate with their coaches weekly or monthly. ` +
      `${notEnoughPct} of respondents (${notEnoughCount} youth) reported ` +
      `their communication was Not Enough; most of these youth communicated ` +
      `with their coach about once a week or 1\u20132 times per month.`
    : "The majority of youth communicate with their coaches weekly or monthly. " +
      "A small number of youth reported their communication was Not Enough.";
  return [makePara(text), makePara("")];
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
      "safe and stable, meaning they can stay there for at least 90 days. Youth " +
      "with unstable housing were then asked where they are currently sleeping."
    ),
    makePara(
      `About ${pctStable} of respondents reported safe and stable housing. ` +
      "For the remainder, current sleeping arrangements are shown in the table below."
    ),
    makeCaption(cap1),
    makeTable(rows07, "Total", cap1),
    await embedChart("chart_02_housing_stability.png", 5.5),
    makePara(
      "Regardless of their current living situation, if youth experienced unstable " +
      "housing in the prior six months, they were asked to identify the reason(s)."
    ),
    makeCaption(cap2),
    makeTable(rows08, "Total youth reporting unstable housing", cap2),
    makePara("\u00b9 Youth could report more than one reason for experiencing unstable housing.", { italic: true }),
    makePara(""),
  ];
}

function sec_education_employment(dfCsv) {
  const total      = dfCsv.length;
  const inSchool   = dfCsv.filter((r) => ["high_school","college_career","ged","graduate"].includes(r.q5_school_status)).length;
  const employed   = dfCsv.filter((r) => ["yes_full_time","yes_part_time"].includes(r.q8_employment_status)).length;
  const inSRows    = dfCsv.filter((r) => ["high_school","college_career","ged","graduate"].includes(r.q5_school_status));
  const inSUnemp   = inSRows.filter((r) => r.q8_employment_status === "no");
  const inSUnempSk = inSUnemp.filter((r) => r.q8b_job_seeking === "yes").length;
  const notSUnemp  = dfCsv.filter((r) => r.q5_school_status === "not_in_school" && r.q8_employment_status === "no");
  const notSUnempSk= notSUnemp.filter((r) => r.q8b_job_seeking === "yes").length;
  const nisuRows   = dfCsv.filter((r) => r.q5_school_status === "not_in_school" && r.q8_employment_status === "no");
  const noDiploma  = nisuRows.filter((r) => r.q5a_highest_education === "some_hs").length;
  const empRows    = dfCsv.filter((r) => ["yes_full_time","yes_part_time"].includes(r.q8_employment_status));
  const empAlsoSch = empRows.filter((r) => ["high_school","college_career","ged","graduate"].includes(r.q5_school_status)).length;

  const rows09 = loadSheet("09_education");
  const cap = "Educational Enrollment and Attainment";
  const items = [
    makeHeading("Employment and Education", 1),
    makeBullet(`${pct(inSchool, total)} of all respondents reported being enrolled in school`),
    makeBullet(
      `${pct(employed, total)} of all respondents reported being employed ` +
      `(${pct(empAlsoSch, employed)} of these youth are also enrolled in school)`
    ),
  ];
  if (inSUnemp.length) items.push(makeBullet(
    `${pct(inSUnempSk, inSUnemp.length)} of respondents who are in school and unemployed are looking for a job`
  ));
  if (notSUnemp.length) items.push(makeBullet(
    `${pct(notSUnempSk, notSUnemp.length)} of respondents who are not in school and unemployed are looking for a job`
  ));
  items.push(makeBullet(
    `${pct(nisuRows.length, total)} of respondents are both not in school and unemployed; ` +
    `${noDiploma} of these ${nisuRows.length} youth report not completing high school or a GED`
  ));
  items.push(makeCaption(cap));
  items.push(makeTable(rows09, "Total", cap));
  items.push(makePara(""));
  return items;
}

function sec_job_tenure(dfCsv) {
  const empRows   = dfCsv.filter((r) => ["yes_full_time","yes_part_time"].includes(r.q8_employment_status));
  const nEmp      = empRows.length;
  const longTenure = empRows.filter((r) => r.q8a_job_tenure === "more_6mo").length;
  const rows11 = loadSheet("11_job_tenure");
  const cap = "Length of Employment for Youth Currently Employed";
  return [
    makePara(
      `Of the ${pct(nEmp, N_RESPONDENTS)} of survey respondents that reported being employed, ` +
      `${pct(longTenure, nEmp)} have been at their job for six months or longer.`
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
    makeCaption(cap),
    makeTable(rows10, "Total", cap),
    await embedChart("chart_03_employment_by_age.png", 5.5),
    makePara(""),
  ];
}

function sec_job_barriers() {
  const rows12  = loadSheet("12_job_barriers");
  const data12  = rows12.filter((r) => !r._header);
  const cols12  = Object.keys(rows12[0] || {}).filter((k) => k !== "_header");
  const topRow  = data12[0];
  const topLabel = topRow ? firstCol(topRow).toLowerCase() : "";
  const topPct   = (topRow && cols12.length > 2) ? getCol(topRow, 2) : "";
  const cap = "Reasons Youth Have Trouble Finding Jobs (Reasons Given by 2 or More People)";
  return [
    makePara(
      "If survey respondents had trouble finding a job in the prior twelve months, " +
      "they were asked to share some of the reasons why. " +
      `The most common challenge identified, reported by ${topPct} of youth, was ${topLabel}.`
    ),
    makeCaption(cap),
    makeTable(rows12, "__none__", cap),
    makePara("Note: Youth could select more than one option", { italic: true }),
    makePara(""),
  ];
}

function sec_left_job() {
  const rows13 = loadSheet("13_left_job");
  const cap = "Reasons Youth Lost or Quit a Job in the Past Year (Reasons Given by 2 or More People)";
  return [
    makePara(
      "If the survey respondent lost or left a job in the past year, they were asked " +
      "to share the reason(s). Sub-reasons for youth who quit are shown indented below the Quit row."
    ),
    makeCaption(cap),
    makeTable(rows13, "__none__", cap),
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
      "of transportation they rely on for work. The table below displays driver\u2019s " +
      "license status by age."
    ),
    makeCaption(cap1),
    makeTable(subs["Driver's License by Age"] || [], "Total", cap1),
    makePara("Of those with a driver\u2019s license, about half regularly have access to a reliable vehicle."),
    makeCaption(cap2),
    makeTable(subs["Vehicle Access (licensed)"] || [], "Total", cap2),
    makePara(
      "All youth were asked about the primary way they get to work when they are " +
      "employed. The majority rely on public transportation."
    ),
    makeCaption(cap3),
    makeTable(subs["Primary Transport"] || [], "Total", cap3),
    makePara(""),
  ];
}

function sec_voter_reg() {
  const subs      = splitSheet("15_voter_reg");
  const dfReg     = subs["Voter Registration by Age"]       || [];
  const dfReasons = subs["Not Registered Reasons by Age"]   || [];
  let totalPctReg = "[PLACEHOLDER]";
  const regData = dfReg.filter((r) => !r._header);
  const regRow  = regData.find((r) => firstCol(r) === "Registered to Vote");
  if (regRow && regRow["Total"]) totalPctReg = regRow["Total"];
  const cap1 = "Self-Reported Voter Registration by Age";
  const cap2 = "Reasons Youth Report Not Registering to Vote";
  return [
    makeHeading("Voting", 1),
    makePara(
      "Youth ages 18 and older were asked to report on whether they are registered " +
      `to vote. Overall, ${totalPctReg} of eligible respondents reported being registered to vote.`
    ),
    makeCaption(cap1),
    makeTable(dfReg, "__none__", cap1),
    makePara(
      "The most common reason youth provided for not registering to vote is that " +
      "they believe their vote won\u2019t make a difference."
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
  const cap1 = "Visit Frequency by Age";
  const cap2 = "What Are the Main Reasons Youth Come to the Youth Zone?";
  const cap3 = "What Would Make Someone Who Rarely Visits the Zone Want to Come, by Age";
  return [
    makeHeading("Zone Experience", 1),
    makePara(
      "This survey includes questions to better understand participants\u2019 " +
      "experiences at the Zone. Attendance patterns at the Zone vary by age, with " +
      "older youth coming more frequently than youth ages 16 to 17."
    ),
    makeCaption(cap1),
    makeTable(dfFreq, "Total", cap1),
    await embedChart("chart_04_visit_frequency.png", 5.5),
    makePara(
      "As in prior years, most youth report coming to the 412 Youth Zone downtown " +
      "to see their coach and to work toward their goals."
    ),
    makeCaption(cap2),
    makeTable(dfReasons, "__none__", cap2),
    makePara(
      "For youth who never visit the Zone, or visit less than monthly, they were " +
      "asked what would make them want to come more frequently."
    ),
    makeCaption(cap3),
    makeTable(dfBarriers, "__none__", cap3),
    makePara(""),
  ];
}

function sec_program_impact(dfCsv) {
  const helpedAny = dfCsv.filter((r) => (r.q17_program_helped || "").trim() !== "").length;
  const pctHelped = pct(helpedAny, N_RESPONDENTS);
  const q16Total  = dfCsv.filter((r) => (r.q16_stay_focused || "") !== "").length;
  const q16Agree  = dfCsv.filter((r) => ["agree","somewhat_agree"].includes(r.q16_stay_focused)).length;
  const pctQ16    = pct(q16Agree, q16Total);

  const subs = splitSheet("17_impact");
  let dfQ17  = subs["Program Helped With (Q17) by A"] || [];
  if (!dfQ17.length) {
    const key = Object.keys(subs).find((k) => k.startsWith("Program Helped"));
    if (key) dfQ17 = subs[key];
  }
  const items = [
    makeHeading("Impact of Assistance", 1),
    makePara(
      "Across the core outcome areas in which Youth Zone staff are helping young " +
      "people make progress, youth reported that their coaches and the Zone have " +
      `helped them in a variety of ways. ${pctHelped} of respondents indicated ` +
      "progress supported by the Zone in at least one area."
    ),
    makePara(
      `${pctQ16} of respondents agreed or somewhat agreed that their coach or ` +
      "the Zone helped them stay focused on their goals."
    ),
  ];
  if (dfQ17.length) {
    const cap = "My Coach or the Youth Zone has Helped Me To\u2026 (by Age)";
    items.push(makeCaption(cap));
    items.push(makeTable(dfQ17, "__none__", cap));
  }
  items.push(makePara(""));
  return items;
}

function sec_respect_environment() {
  const rows18 = loadSheet("18_respect");
  const data18 = rows18.filter((r) => !r._header);
  let pctStaff = "[PLACEHOLDER]", pctPeer = "[PLACEHOLDER]";
  if (rows18.length && rows18[0]["% Often or All the Time"] !== undefined) {
    const staffRow = data18.find((r) => firstCol(r).includes("Staff"));
    const peerRow  = data18.find((r) => firstCol(r).includes("Peer"));
    if (staffRow) pctStaff = staffRow["% Often or All the Time"] || "[PLACEHOLDER]";
    if (peerRow)  pctPeer  = peerRow["% Often or All the Time"]  || "[PLACEHOLDER]";
  }
  const rows19 = loadSheet("19_environment");
  const has19  = rows19.filter((r) => !r._header).length > 0;
  const items = [
    makePara(
      "Respondents were asked to rate how often they felt respected and accepted " +
      `for who they are at the Youth Zone. ${pctStaff} of youth reported staff ` +
      `treat them with respect often or all the time; ${pctPeer} said the same ` +
      "about their peers at the Zone."
    ),
  ];
  if (has19) {
    items.push(makePara(
      "Youth also rated five statements about the program environment on a " +
      "1\u20135 scale. Results are shown in terms of the percentage selecting 4 or 5 (top-2 box)."
    ));
  }
  items.push(makePara(""));
  return items;
}

function sec_banking(dfCsv) {
  const hasAccount = dfCsv.filter((r) => {
    const parts = (r.q25_bank_account || "").split("|");
    return parts.some((t) => t === " checking" || t === " savings" || t === "checking" || t === "savings");
  }).length;
  const pctHas = pct(hasAccount, N_RESPONDENTS);
  const subs = splitSheet("20_banking");
  const cap1 = "Banking Status by Age";
  const cap2 = "Methods Youth Use to Store, Receive, and Transfer Money, by Age";
  const cap3 = "Ways Respondents Use Their Bank Account(s), by Age";
  return [
    makeHeading("Banking", 1),
    makePara(
      "Participants were asked questions about their use of banks and other ways " +
      "that they store, receive, and transfer money. " +
      `Overall, ${pctHas} of respondents reported currently having a bank account.`
    ),
    makeCaption(cap1),
    makeTable(subs["Bank Account Status by Age"]     || [], "__none__", cap1),
    makeCaption(cap2),
    makeTable(subs["Money Methods by Age (Q24)"]     || [], "__none__", cap2),
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
  const rows22  = loadSheet("22_comments");
  const data22  = rows22.filter((r) => !r._header);
  const nCom    = data22.filter((r) => firstCol(r) !== "").length;
  const items = [
    makeHeading("Additional Comments", 1),
    makePara(
      `Finally, youth had the option to share any other comments or feedback ` +
      `they had about the Zone. ${nCom} youth provided additional comments.`
    ),
  ];
  if (data22.length) {
    items.push(makeHeading("Comments", 1));
    const cols = Object.keys(rows22[0] || {}).filter((k) => k !== "_header");
    const commentKey = cols.includes("Comment") ? "Comment" : cols[cols.length - 1];
    for (const row of data22) {
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

  console.log("Building sections...");
  const children = [
    ...sec_title(),
    ...sec_age(),
    ...sec_gender_orient(),
    ...(await sec_race()),
    ...(await sec_coach_satisfaction_async()),
    ...sec_communication(),
    ...(await sec_housing()),
    ...sec_education_employment(dfCsv),
    ...sec_job_tenure(dfCsv),
    ...(await sec_employment_by_age()),
    ...sec_job_barriers(),
    ...sec_left_job(),
    ...sec_transportation(),
    ...sec_voter_reg(),
    ...(await sec_zone_visit()),
    ...sec_program_impact(dfCsv),
    ...sec_respect_environment(),
    ...sec_banking(dfCsv),
    ...(await sec_nps()),
    ...sec_comments(),
  ];

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
