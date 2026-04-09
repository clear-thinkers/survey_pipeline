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
      "participate in a survey in early 2026. Surveys were administered on paper " +
      "and online."
    ),
    new Paragraph({
      spacing: BODY_PARAGRAPH_SPACING,
      children: [
        new TextRun({ text: `${N_RESPONDENTS} unique youth (of `, font: "Calibri", size: 22 }),
        new TextRun({ text: "[TOTAL ACTIVE \u2014 fill in denominator]", highlight: "yellow", bold: true, font: "Calibri", size: 22 }),
        new TextRun({ text: " total active) responded to the survey, for a response rate of ", font: "Calibri", size: 22 }),
        new TextRun({ text: "[RESPONSE RATE %]", highlight: "yellow", bold: true, font: "Calibri", size: 22 }),
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
    makeTable([headerRow, row1, row2, row3, ...dataRowsQ1], "__none__", cap),
    await embedChart("chart_01_coach_satisfaction.png", 5.5),
    makePara(""),
  ];
}

async function sec_communication() {
  // Q3 communication level counts from analysis sheet
  const rows = loadSheet("06_communication");
  const data = rows.filter((r) => !r._header);
  const cols = Object.keys(rows[0] || {}).filter((k) => k !== "_header");
  let goodPct = "", notEnoughCount = "", notEnoughPct = "", tooMuchPct = "";
  for (const row of data) {
    const first = String(row[cols[0]] ?? "").toLowerCase();
    if (first.includes("good amount")) {
      goodPct = cols.length > 2 ? String(row[cols[2]] ?? "") : "";
    }
    if (first.includes("not enough")) {
      notEnoughCount = String(row[cols[1]] ?? "");
      notEnoughPct   = cols.length > 2 ? String(row[cols[2]] ?? "") : "";
    }
    if (first.includes("too much")) {
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
  const nNeMonthly   = neRows.filter((r) => r.q2_communication_frequency === "1_2_times_per_month").length;
  const nNeLess      = neRows.filter((r) => r.q2_communication_frequency === "less_than_once_a_month").length;
  const nNeWeekly    = neRows.filter((r) => r.q2_communication_frequency === "about_once_a_week").length;

  const para1 = (goodPct && notEnoughPct && tooMuchPct)
    ? `${goodPct} of respondents rated their communication with their coach a Good amount. ` +
      `${notEnoughPct} of respondents (${notEnoughCount} youth) reported their communication was Not Enough, ` +
      `a slight increase from 11% (19 youth) in February 2025; ${tooMuchPct} reported Too much.`
    : "The majority of youth communicate with their coaches weekly or monthly. " +
      "A small number of youth reported their communication was Not Enough.";

  const para2 = goodQ2Total
    ? `Among youth who rated communication a Good amount, nearly half communicated ` +
      `1\u20132 times per month (${pctGoodMonthly}), about a third connected with their coach ` +
      `about once a week (${pctGoodWeekly}), and ${pctGoodDaily} communicated almost every day.`
    : "";

  const para3 = neRows.length
    ? `Most of the ${notEnoughCount} youth who reported Not Enough communicated with their coach ` +
      `1\u20132 times per month (${nNeMonthly} youth) or less than once a month (${nNeLess} youth); ` +
      `${nNeWeekly} youth reported about once a week.`
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
  const empRows    = dfCsv.filter((r) => ["yes_full_time","yes_part_time"].includes(r.q8_employment_status));
  const empAlsoSch = empRows.filter((r) => ["high_school","college_career","ged","graduate"].includes(r.q5_school_status)).length;

  const rows09 = loadSheet("09_education");
  const cap = "Educational Enrollment and Attainment";
  const items = [
    makeHeading("Employment and Education", 1),
    makePara(
      "Youth were asked to report on whether they are attending school, working, and, if not working, trying to find a job. This year:"
    ),
    makeBullet(`${pct(inSchool, total)} of all respondents reported being enrolled in school, down from 53% in March 2025`),
    makeBullet(
      `${pct(employed, nQ8Total)} of all respondents reported being employed ` +
      `(${pct(empAlsoSch, employed)} of these youth are also enrolled in school)`
    ),
  ];
  if (inSUnemp.length) items.push(makeBullet(
    `${pct(inSUnempSk, inSUnemp.length)} of respondents who are in school and unemployed are looking for a job, down from 78% in March 2025`
  ));
  if (notSUnemp.length) items.push(makeBullet(
    `${pct(notSUnempSk, notSUnemp.length)} of respondents who are not in school and unemployed are looking for a job, up from 68% in March 2025`
  ));
  items.push(makeBullet(
    `${pct(nisuRows.length, total)} of respondents are both not in school and unemployed, down from 26% last year; ` +
    `${noDiploma} of these ${nisuRows.length} youth report not completing high school or a GED. ` +
    `Most of this group (${pct(notSUnempSk, notSUnemp.length)}) reported looking for work.`
  ));
  items.push(makePara(
    "Among youth who reported being enrolled in school, respondents were split almost evenly between high school students (45%) and youth in college, vocational, or graduate programs (51% combined), while 4% reported being enrolled in a GED program. " +
    "This is a notable change from March 2025, when about two-thirds of enrolled youth were in high school, suggesting this year’s in-school respondents were more evenly distributed across secondary and postsecondary education."
  ));
  items.push(await embedChart("chart_09_employment_by_school.png", 6.6));
  items.push(makeCaption(cap));
  items.push(makeTable(rows09, "Total", cap));
  items.push(makePara(""));
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

function sec_job_barriers() {
  const rows12  = loadSheet("12_job_barriers");
  const data12  = rows12.filter((r) => !r._header);
  const cols12  = Object.keys(rows12[0] || {}).filter((k) => k !== "_header");
  const topRow  = data12[0];
  const secondRow = data12[1];
  const thirdRow = data12[2];
  const fourthRow = data12[3];
  const topLabel = topRow ? firstCol(topRow).toLowerCase() : "";
  const topPct   = (topRow && cols12.length > 2) ? getCol(topRow, 2) : "";
  const secondLabel = secondRow ? firstCol(secondRow).toLowerCase() : "";
  const secondPct = (secondRow && cols12.length > 2) ? getCol(secondRow, 2) : "";
  const thirdLabel = thirdRow ? firstCol(thirdRow).toLowerCase() : "";
  const thirdPct = (thirdRow && cols12.length > 2) ? getCol(thirdRow, 2) : "";
  const fourthPct = (fourthRow && cols12.length > 2) ? getCol(fourthRow, 2) : "";
  const cap = "Reasons Youth Have Trouble Finding Jobs (Reasons Given by 2 or More People)";
  return [
    makePara(
      "If survey respondents had trouble finding a job in the prior twelve months, " +
      "they were asked to share some of the reasons they believed were contributing to those challenges. " +
      "Youth were offered a multiple-choice list of reasons and could also add their own responses."
    ),
    makePara(
      `The most common challenge identified this year, reported by ${topPct} of youth, was ${topLabel}, followed closely by ${secondLabel} (${secondPct}) and ${thirdLabel} (${thirdPct}). ` +
      `This represents a clear shift from March 2025, when transportation issues were the leading barrier at 32% and applying without getting called back was reported by 28% of youth. Mental or physical health remained a substantial barrier for ${fourthPct} of youth, and nearly all of those reporting that challenge were age 18 or older.`
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
      "to share the reason(s)."
    ),
    makePara(
      "This year, the most common reason youth reported leaving a job was finding a better job, cited by 33% of youth who reported at least one reason for leaving. Twenty-eight percent reported quitting, and 26% said the job was seasonal or temporary. Among the youth who quit, the most common specific reasons were low pay or not enough hours and mental or emotional health, each reported by 8% of youth who left a job; personal or family reasons followed closely at 7%."
    ),
    makePara(
      "Unlike the prior year, quitting was not the dominant reason for job separation in this year’s data. Youth ages 18 to 20 and 21 to 23 both accounted for most reports of leaving a job, but no single age group clearly dominated the overall pattern of job separation. Sub-reasons for youth who quit are shown indented below the Quit row."
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
  const pctHelped = pct(helpedAny, dfCsv.length);
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
  const subs = splitSheet("20_banking");
  const bankRows = (subs["Bank Account Status by Age"] || []).filter((r) => !r._header);
  const hasAcctRow = bankRows.find((r) => firstCol(r).toLowerCase().includes("currently have"));
  const pctHas = (hasAcctRow && hasAcctRow["Percent of Total"])
    ? hasAcctRow["Percent of Total"]
    : pct(dfCsv.filter((r) => {
        const parts = (r.q25_bank_account || "").split("|");
        return parts.some((t) => ["checking", "savings"].includes(t.trim()));
      }).length, dfCsv.length);
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
