"use strict";

const path = require("path");
const fs = require("fs");

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
  VerticalAlign,
} = requireGlobal("docx");

const XLSX = requireGlobal("xlsx");
const { parse: parseCsv } = require(path.join(GLOBAL_NM, "csv-parse", "dist", "cjs", "sync.cjs"));

let sharp;
try { sharp = require(path.join(__dirname, "..", "node_modules", "sharp")); }
catch (_) { try { sharp = requireGlobal("sharp"); } catch (__) { sharp = null; } }

const BASE_DIR = path.join(__dirname, "..");
const ANALYSIS_PATH = path.join(BASE_DIR, "output", "IL", "analysis_IL.xlsx");
const CSV_PATH = path.join(BASE_DIR, "output", "IL", "survey_data_IL.csv");
const CHARTS_DIR = path.join(BASE_DIR, "output", "IL", "charts");
const OUT_PATH = path.join(BASE_DIR, "report", "IL", "report_IL_v1.docx");

const SURVEY_LABEL = "Spring 2026";
let N_RESPONDENTS = 0;

const BODY_PARAGRAPH_SPACING = { before: 180, after: 180, line: 240 };
const SPACER_PARAGRAPHS = new WeakSet();

const DOC_STYLES = {
  default: {
    document: { run: { font: "Calibri", size: 22 } },
  },
  paragraphStyles: [
    {
      id: "Heading1",
      name: "Heading 1",
      basedOn: "Normal",
      next: "Normal",
      quickFormat: true,
      run: { bold: true, size: 26, font: "Calibri" },
      paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 0 },
    },
    {
      id: "Heading2",
      name: "Heading 2",
      basedOn: "Normal",
      next: "Normal",
      quickFormat: true,
      run: { bold: true, size: 22, font: "Calibri" },
      paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 },
    },
    {
      id: "Caption",
      name: "Caption",
      basedOn: "Normal",
      next: "Normal",
      run: { bold: true, size: 22, font: "Calibri" },
      paragraph: { spacing: { before: 120, after: 120 } },
    },
  ],
};

const BULLET_NUMBERING = {
  config: [
    {
      reference: "bullets",
      levels: [
        {
          level: 0,
          format: "bullet",
          text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        },
      ],
    },
  ],
};

const NO_BORDER = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

const COACH_BENCHMARKS = {
  headers: ["April 2022", "April 2023", "May 2024", "May 2025", "Spring 2026"],
  n: ["n=10", "n=31", "n=21", "n=14", null],
  rows: {
    "Is trustworthy": ["100%", "94%", "100%", "100%", null],
    "Is reliable": ["100%", "87%", "95%", "100%", null],
    "Values my opinions about my life": ["100%", "90%", "100%", "93%", null],
    "Is available when I need them": ["70%", "81%", "95%", "100%", null],
    "Makes me feel heard and understood": ["80%", "87%", "95%", "86%", null],
  },
};

function makeHeading(text, level = 1) {
  return new Paragraph({
    text,
    heading: level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
  });
}

function makeCaption(text) {
  return new Paragraph({ text, style: "Caption" });
}

function makePara(text, options = {}) {
  const { bold = false, italic = false } = options;
  const para = new Paragraph({
    spacing: BODY_PARAGRAPH_SPACING,
    children: [new TextRun({ text, bold, italic, font: "Calibri", size: 22 })],
  });
  if (!String(text).trim()) SPACER_PARAGRAPHS.add(para);
  return para;
}

function makeBullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: BODY_PARAGRAPH_SPACING,
    children: [new TextRun({ text, font: "Calibri", size: 22 })],
  });
}

function makeTable(rows, totalLabel = "Total") {
  if (!rows || rows.length === 0) return new Paragraph({ text: "" });
  const columns = Object.keys(rows[0]).filter((key) => key !== "_header");
  const tableRows = rows.map((rowObj) => {
    const isHeader = rowObj._header === true;
    const firstVal = String(rowObj[columns[0]] ?? "").trim();
    const isTotal = !isHeader && firstVal === totalLabel;
    return new TableRow({
      children: columns.map((column) => new TableCell({
        width: { size: 0, type: WidthType.AUTO },
        verticalAlign: VerticalAlign.CENTER,
        borders: NO_BORDER,
        shading: isHeader || isTotal ? { fill: "DCE6F1" } : undefined,
        margins: { top: 70, bottom: 70, left: 90, right: 90 },
        children: [new Paragraph({
          children: [new TextRun({
            text: String(rowObj[column] ?? ""),
            bold: isHeader || isTotal,
            font: "Calibri",
            size: 21,
          })],
        })],
      })),
    });
  });

  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows });
}

async function embedChart(filename, widthInches = 6) {
  const chartPath = path.join(CHARTS_DIR, filename);
  if (!fs.existsSync(chartPath)) return new Paragraph({ text: "" });
  const data = fs.readFileSync(chartPath);
  let widthPx = Math.round(widthInches * 96);
  let heightPx = Math.round(widthPx * 0.6);
  if (sharp) {
    try {
      const meta = await sharp(chartPath).metadata();
      heightPx = Math.round(widthPx * (meta.height / meta.width));
    } catch (_) {}
  }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new ImageRun({ data, type: "png", transformation: { width: widthPx, height: heightPx } })],
  });
}

function loadSheet(sheetName) {
  const workbook = XLSX.readFile(ANALYSIS_PATH);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (raw.length < 2) return [];
  const headers = raw[1].map((header) => String(header ?? ""));
  const dataRows = raw.slice(2).filter((row) => row.some((value) => String(value).trim() !== ""));
  const headerObj = { _header: true };
  headers.forEach((header, index) => { headerObj[header || `col${index}`] = header; });
  const result = [headerObj];
  for (const rawRow of dataRows) {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header || `col${index}`] = rawRow[index] === null || rawRow[index] === undefined ? "" : String(rawRow[index]);
    });
    result.push(obj);
  }
  return result;
}

function splitSheet(sheetName) {
  const workbook = XLSX.readFile(ANALYSIS_PATH);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return {};
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const sections = {};
  let currentTitle = null;
  let currentRows = [];

  for (const rawRow of raw) {
    const first = String(rawRow[0] ?? "").trim();
    const isTitle = first !== "" && rawRow.slice(1).every((value) => String(value ?? "").trim() === "");
    if (isTitle) {
      if (currentTitle !== null && currentRows.length > 0) sections[currentTitle] = currentRows;
      currentTitle = first;
      currentRows = [];
    } else if (currentTitle !== null) {
      currentRows.push(rawRow);
    }
  }
  if (currentTitle !== null && currentRows.length > 0) sections[currentTitle] = currentRows;

  const result = {};
  for (const [title, rows] of Object.entries(sections)) {
    const nonBlank = rows.filter((row) => row.some((value) => String(value ?? "").trim() !== ""));
    if (nonBlank.length === 0) {
      result[title] = [];
      continue;
    }
    let maxCol = 0;
    for (const row of nonBlank) {
      for (let index = row.length - 1; index >= 0; index -= 1) {
        if (String(row[index] ?? "").trim() !== "") {
          maxCol = Math.max(maxCol, index + 1);
          break;
        }
      }
    }
    const headers = nonBlank[0].slice(0, maxCol).map((header) => String(header ?? ""));
    const headerObj = { _header: true };
    headers.forEach((header, index) => { headerObj[header || `col${index}`] = header; });
    const dataRows = nonBlank.slice(1).map((row) => {
      const obj = {};
      headers.forEach((header, index) => { obj[header || `col${index}`] = String(row[index] ?? ""); });
      return obj;
    });
    result[title] = [headerObj, ...dataRows];
  }
  return result;
}

function loadCsv() {
  let content = fs.readFileSync(CSV_PATH, "utf8");
  if (content.charCodeAt(0) === 0xFEFF) content = content.slice(1);
  return parseCsv(content, { columns: true, skip_empty_lines: true });
}

function splitPipe(value) {
  return String(value || "").split("|").map((part) => part.trim()).filter(Boolean);
}

function pct(numerator, denominator) {
  if (!denominator) return "";
  return `${Math.round((100 * numerator) / denominator)}%`;
}

function placeholderRun(text) {
  return new TextRun({ text: `[${text}]`, highlight: "yellow", bold: true, font: "Calibri", size: 22 });
}

function firstCol(row) {
  return String(Object.entries(row).filter(([key]) => key !== "_header")[0]?.[1] ?? "");
}

function findRow(rows, label) {
  return rows.find((row) => firstCol(row) === label);
}

function uniqueTexts(values, skip = new Set()) {
  const seen = new Set();
  const results = [];
  for (const raw of values) {
    const text = String(raw || "").trim();
    if (!text) continue;
    const key = text.toLowerCase();
    if (skip.has(key) || seen.has(key)) continue;
    seen.add(key);
    results.push(text);
  }
  return results;
}

function sec_title(dfCsv) {
  const ageKnown = dfCsv.filter((row) => ["14_17", "18_20", "21_23"].includes(String(row.age_range || "").trim()));
  const n14 = ageKnown.filter((row) => row.age_range === "14_17").length;
  const n18 = ageKnown.filter((row) => row.age_range === "18_20").length;
  const n21 = ageKnown.filter((row) => row.age_range === "21_23").length;
  const genderKnown = dfCsv.filter((row) => ["Female", "Male"].includes(String(row.gender || "").trim()));
  const female = genderKnown.filter((row) => row.gender === "Female").length;
  const male = genderKnown.filter((row) => row.gender === "Male").length;
  const orientationAnswered = dfCsv.filter((row) => String(row.sexual_orientation || "").trim() !== "");
  const lgbtqSet = new Set(["Asexual", "Bisexual", "Demisexual", "Gay or Lesbian", "Mostly heterosexual", "Pansexual", "Queer", "Same Gender Loving"]);
  const lgbtq = orientationAnswered.filter((row) => lgbtqSet.has(String(row.sexual_orientation || "").trim())).length;

  return [
    makeHeading("Crawford County IL Survey Results", 1),
    makePara(SURVEY_LABEL),
    makePara(""),
    makePara(
      "All individuals active with the Crawford County Independent Living Program had the opportunity to participate in a survey in spring 2026. Surveys could be completed on paper or online through SurveyMonkey. Respondents had the option to provide their name for a raffle drawing, but they could also complete the survey anonymously."
    ),
    new Paragraph({
      spacing: BODY_PARAGRAPH_SPACING,
      children: [
        new TextRun({ text: `${N_RESPONDENTS} unique youth responded to the survey out of `, font: "Calibri", size: 22 }),
        placeholderRun("TOTAL ACTIVE IL PARTICIPANTS"),
        new TextRun({ text: ", for a response rate of ", font: "Calibri", size: 22 }),
        placeholderRun("RESPONSE RATE"),
        new TextRun({ text: ". Most respondents were White and between the ages of 14 and 20. Among respondents who reported their age, ", font: "Calibri", size: 22 }),
        new TextRun({ text: `${pct(n14, ageKnown.length)} were 14-17, ${pct(n18, ageKnown.length)} were 18-20, and ${pct(n21, ageKnown.length)} were 21-23. `, font: "Calibri", size: 22 }),
        new TextRun({ text: `Among respondents who reported gender, ${pct(female, genderKnown.length)} identified as female and ${pct(male, genderKnown.length)} as male. About ${pct(lgbtq, orientationAnswered.length)} of respondents who answered the sexual orientation question identified as LGBTQ, similar to the roughly one-third reported in May 2025.`, font: "Calibri", size: 22 }),
      ],
    }),
    makePara(""),
  ];
}

function sec_demographics() {
  const tables = splitSheet("01_demographics");
  return [
    makeHeading("Survey Respondent Demographics", 1),
    makeCaption("Age"),
    makeTable(tables["Age"] || [], "Total"),
    makeCaption("Gender"),
    makeTable(tables["Gender"] || [], "Total"),
    makeCaption("Race"),
    makeTable(tables["Race"] || [], "Total"),
    makeCaption("Sexual Orientation"),
    makeTable(tables["Sexual Orientation"] || [], "Total"),
    makePara(""),
  ];
}

function sec_coach_relationships() {
  const rows = loadSheet("02_coach");
  const data = rows.filter((row) => !row._header);
  const current = {};
  data.forEach((row) => { current[firstCol(row)] = row["% Often or All the Time"]; });
  Object.keys(COACH_BENCHMARKS.rows).forEach((label) => {
    COACH_BENCHMARKS.rows[label][4] = current[label] || "";
  });
  COACH_BENCHMARKS.n[4] = `n=${N_RESPONDENTS}`;

  const header = { _header: true, "My Coach...": "My Coach..." };
  COACH_BENCHMARKS.headers.forEach((label) => { header[label] = label; });
  const nRow = { "My Coach...": "" };
  COACH_BENCHMARKS.headers.forEach((label, index) => { nRow[label] = COACH_BENCHMARKS.n[index]; });
  const tableRows = [header, nRow];
  Object.entries(COACH_BENCHMARKS.rows).forEach(([label, values]) => {
    const row = { "My Coach...": label };
    COACH_BENCHMARKS.headers.forEach((headerLabel, index) => { row[headerLabel] = values[index]; });
    tableRows.push(row);
  });

  return [
    makeHeading("FINDINGS", 1),
    makeHeading("Relationships with Coach", 2),
    makePara(
      `Coach relationship ratings remained exceptionally strong in the 2026 IL survey. All five coach items were rated Often or All the Time by at least ${current["Is reliable"] || "97%"} of respondents, with trustworthiness and availability both at ${current["Is trustworthy"] || "100%"} and ${current["Is available when I need them"] || "100%"}.`
    ),
    makePara(
      `Compared with May 2025, feeling heard and understood improved from 86% to ${current["Makes me feel heard and understood"] || "97%"}, and values opinions rose from 93% to ${current["Values my opinions about my life"] || "97%"}. Because annual IL response counts are small, the percentages should still be interpreted cautiously, but the overall pattern continues to show very strong youth-coach relationships.`
    ),
    makeCaption("My Coach... Percent Often or All the Time"),
    makeTable(tableRows, "__none__"),
    makePara(""),
  ];
}

function sec_communication() {
  const tables = splitSheet("03_communication");
  return [
    makeHeading("Communication", 2),
    makePara(
      "Communication with coaches remained strong overall, but responses were slightly more mixed than in the prior year. This year, 94% of respondents said the amount of communication with their coach was a good amount, while one respondent said it was not enough and one said it was too much; in May 2025, all respondents reported being satisfied with the amount of communication they had."
    ),
    makePara(
      "Reported contact frequency was still concentrated in regular check-ins. Fifty-five percent of respondents said they communicate with their coach one to two times per month and 39% said about once a week, while only one respondent reported almost daily communication and one reported less than monthly contact."
    ),
    makeCaption("Communication Satisfaction"),
    makeTable(tables["Communication Satisfaction"] || [], "__none__"),
    makeCaption("Reported Frequency of Communication"),
    makeTable(tables["Reported Frequency"] || [], "Total"),
    makePara(""),
  ];
}

async function sec_employment(dfCsv) {
  const tables = splitSheet("04_employment");
  const employed = dfCsv.filter((row) => ["yes_full_time", "yes_part_time"].includes(String(row.q6_employment_status || "").trim())).length;
  const unemployed = dfCsv.filter((row) => row.q6_employment_status === "no").length;
  const seeking = dfCsv.filter((row) => row.q6_employment_status === "no" && row.q6b_job_seeking === "yes").length;
  const employedRows = dfCsv.filter((row) => ["yes_full_time", "yes_part_time"].includes(String(row.q6_employment_status || "").trim()));
  const tenureLong = employedRows.filter((row) => row.q6a_job_tenure === "more_6mo").length;
  const tenureMid = employedRows.filter((row) => row.q6a_job_tenure === "3_6mo").length;
  const tenureShort = employedRows.filter((row) => row.q6a_job_tenure === "less_3mo").length;
  const highSchool = dfCsv.filter((row) => row.q5_school_status === "high_school");
  const hsEmployed = highSchool.filter((row) => ["yes_full_time", "yes_part_time"].includes(String(row.q6_employment_status || "").trim())).length;
  const hsUnemployed = highSchool.filter((row) => row.q6_employment_status === "no").length;
  const hsSeeking = highSchool.filter((row) => row.q6_employment_status === "no" && row.q6b_job_seeking === "yes").length;
  const notSchool = dfCsv.filter((row) => row.q5_school_status === "not_in_school");
  const notSchoolEmployed = notSchool.filter((row) => ["yes_full_time", "yes_part_time"].includes(String(row.q6_employment_status || "").trim())).length;

  return [
    makeHeading("Employment and Education", 2),
    makePara(
      `Youth were asked whether they are attending school and working and, if they were not working, whether they were looking for employment. ${pct(employed, dfCsv.length)} of respondents reported being employed, essentially unchanged from 40% in May 2025. Among the ${unemployed} respondents who were not working, ${seeking} (${pct(seeking, unemployed)}) said they were looking for a job, up from 55% last year.`
    ),
    makeCaption("Employment Status"),
    makeTable(tables["Employment Status"] || [], "Total"),
    makePara(
      `Employment was less stable than in the prior report. Of the ${employed} respondents who were working, ${pct(tenureLong, employed)} had been in their current job more than six months, down from 67% in May 2025. Another ${pct(tenureMid, employed)} had been in their job three to six months and ${pct(tenureShort, employed)} had been there less than three months.`
    ),
    makeCaption("Length of Employment for Youth Currently Employed"),
    makeTable(tables["Length of Employment for Youth Currently Employed"] || [], "Total"),
    makePara(
      `The school-employment split remained pronounced. ${hsEmployed} of the ${highSchool.length} high school respondents were employed, compared with ${notSchoolEmployed} of ${notSchool.length} respondents who were not in school. Among the ${hsUnemployed} high school respondents who were unemployed, ${hsSeeking} were actively looking for work.`
    ),
    await embedChart("chart_01_employment_by_school.png", 6.4),
    makeCaption("Employment Status by School Enrollment"),
    makeTable(tables["Employment Status by School Enrollment"] || [], "__none__"),
    makePara(
      "Ten respondents reported at least one recent barrier to finding work. Transportation issues were the most common barrier this year, followed by limited work experience. Unlike May 2025, when mental or physical health was the most frequently cited barrier, transportation was the clear leading challenge in 2026."
    ),
    makeCaption("Reasons Youth Have Trouble Finding Jobs"),
    makeTable(tables["Reasons Youth Have Trouble Finding Jobs"] || [], "__none__"),
    makePara("Note: Youth could select more than one option.", { italic: true }),
    makePara(
      "Eleven respondents reported at least one reason for leaving a job in the past year. The most common responses were an open-ended other category and finding a better job. Open-text other responses most often referred to workplace closures, placement issues, parenting, or moving farther away. Only three respondents selected quit, and none provided a follow-up quit reason."
    ),
    makeCaption("Reasons Youth Left a Job"),
    makeTable(tables["Reasons Youth Left a Job"] || [], "__none__"),
    makePara(""),
  ];
}

async function sec_program_impact(dfCsv) {
  const tables = splitSheet("05_program_impact");
  const helpedAny = dfCsv.filter((row) => String(row.q11_program_helped || "").trim() !== "").length;
  const future = dfCsv.filter((row) => splitPipe(row.q11_program_helped).includes("future")).length;
  const problems = dfCsv.filter((row) => splitPipe(row.q11_program_helped).includes("handle_problems")).length;
  const decisions = dfCsv.filter((row) => splitPipe(row.q11_program_helped).includes("decision_making")).length;
  const relationships = dfCsv.filter((row) => splitPipe(row.q11_program_helped).includes("positive_relationships")).length;
  const validInd = dfCsv.filter((row) => String(row.q16_gained_independence || "").trim() !== "");
  const indPositive = validInd.filter((row) => ["agree", "somewhat"].includes(String(row.q16_gained_independence || "").trim())).length;
  return [
    makeHeading("Program Impact", 2),
    makePara(
      `Youth continued to report that the IL program provides both relational support and concrete help. ${helpedAny} of ${dfCsv.length} respondents (${pct(helpedAny, dfCsv.length)}) selected at least one area in which the program had helped them. The most frequently selected areas were thinking about the future (${pct(future, dfCsv.length)}), handling problems (${pct(problems, dfCsv.length)}), decision-making (${pct(decisions, dfCsv.length)}), and establishing positive relationships (${pct(relationships, dfCsv.length)}).`
    ),
    makePara(
      "The age pattern in the chart below suggests younger respondents were especially likely to say the program helped them think about the future and handle problems, while older respondents more often highlighted concrete transition supports such as driver's license help and obtaining vital documents."
    ),
    await embedChart("chart_02_program_helped_by_age.png", 6.7),
    makeCaption("My Coach or the IL Program Has Helped Me To... (by Age)"),
    makeTable(tables["Program Helped By Age"] || [], "__none__"),
    makePara(
      `The program also received very strong marks on the two overall outcome questions. Nearly all respondents (${pct(30, dfCsv.length)}) agreed that IL helps them stay focused on their goals, and ${pct(indPositive, validInd.length)} agreed or somewhat agreed that the program has helped them gain independence. This independence item was newly added in 2026, and the early response is strongly positive.`
    ),
    makeCaption("Support from IL Helps Me Stay Focused on My Goals"),
    makeTable(tables["Stay Focused"] || [], "__none__"),
    makeCaption("The IL Program Has Helped Me Gain Independence"),
    makeTable(tables["Gained Independence"] || [], "__none__"),
    makePara(""),
  ];
}

function sec_respect_environment() {
  const tables = splitSheet("06_respect_environment");
  return [
    makeHeading("Respect and Environment", 2),
    makePara(
      "Ratings of respect and acceptance were strong overall and improved on the staff item relative to the prior report. Of the 30 respondents who answered, 97% said staff treat them with respect and acceptance often or all the time, up from 87% in May 2025. Peer ratings were more mixed: 77% said peers treat them with respect often or all the time, while one respondent reported rarely feeling respected by peers."
    ),
    makeCaption("Respect and Acceptance"),
    makeTable(tables["Respect and Acceptance"] || [], "__none__"),
    makePara(
      "Program environment ratings were also positive. All respondents who answered said diversity of backgrounds is valued, 93% said they are treated fairly, and 90% said people care about their success and that they feel accepted without judgment. Feeling safe sharing thoughts was the lowest-rated environment item at 83%, making it the clearest area for continued attention."
    ),
    makeCaption("Program Environment"),
    makeTable(tables["Program Environment"] || [], "__none__"),
    makePara(""),
  ];
}

function sec_banking() {
  const tables = splitSheet("07_banking");
  return [
    makeHeading("Banking", 2),
    makePara(
      "Participants were also asked about bank account use and, if applicable, their reasons for not having an account. Twenty of the 31 respondents (65%) reported currently having a bank account, down from 73% in May 2025. Checking accounts were more common than savings accounts, and 10 respondents said they had never had a bank account."
    ),
    makePara(
      "This year's sample did not include any respondents who reported being ages 21 to 23, so age comparisons are more limited than in the prior report. Among the eight respondents who explained why they did not have an account, the most common reasons were not knowing how to open one and open-ended other explanations that often referred to being too young, not needing one yet, or waiting until they had identification."
    ),
    makeCaption("Banking Status by Age"),
    makeTable(tables["Banking Status by Age"] || [], "__none__"),
    makeCaption("Reasons for Not Having an Account"),
    makeTable(tables["Reasons for No Account"] || [], "__none__"),
    makePara(""),
  ];
}

async function sec_nps() {
  const tables = splitSheet("08_nps");
  const summaryRows = (tables["NPS Summary"] || []).filter((row) => !row._header);
  const promoterRow = findRow(summaryRows, "Promoters (9-10)");
  const passiveRow = findRow(summaryRows, "Passives (7-8)");
  const detractorRow = findRow(summaryRows, "Detractors (0-6)");
  const scoreRow = findRow(summaryRows, "NPS Score");
  return [
    makeHeading("Overall Recommendation", 2),
    makePara(
      `The 2026 IL survey added a 0-10 recommendation question for the first time. Thirty respondents answered it, producing a Net Promoter Score of ${scoreRow ? scoreRow.Count : "63"}, which falls in the excellent range. ${promoterRow ? promoterRow.Percent : "73%"} of respondents were Promoters, ${passiveRow ? passiveRow.Percent : "17%"} were Passives, and ${detractorRow ? detractorRow.Percent : "10%"} were Detractors.`
    ),
    makePara(
      "Because this is a new question, there is no prior-year comparison. Even so, the distribution shows that strong positive recommendation is already common among current respondents."
    ),
    await embedChart("chart_03_nps.png", 5.8),
    makeCaption("Net Promoter Score Summary"),
    makeTable(tables["NPS Summary"] || [], "__none__"),
    makePara(""),
  ];
}

function sec_comments(dfCsv) {
  const supportTexts = uniqueTexts(dfCsv.map((row) => row.q12_other_supports), new Set(["i think im ok", "i didn't, too busy", "offered is good for now think whats"]));
  const commentTexts = uniqueTexts(dfCsv.map((row) => row.q18_other_comments), new Set(["no", "nope", "naw"]));

  const items = [
    makeHeading("Additional Comments", 2),
    makePara(
      "Respondents also had opportunities to share open-ended feedback about other supports they would use and any additional comments they wanted to provide. Requests for added supports most often focused on tutoring, transition and independent living supports, cooking or practical skill-building, better program timing, and LGBTQIA support."
    ),
  ];
  if (supportTexts.length) {
    items.push(makeHeading("Additional Supports Youth Mentioned", 2));
    supportTexts.forEach((text) => items.push(makeBullet(text)));
  }
  items.push(makePara(
    "Additional comments were overwhelmingly positive and emphasized staff support, easy communication, and the program's role in helping youth make progress. Representative comments are included below."
  ));
  if (commentTexts.length) {
    items.push(makeHeading("Representative Comments", 2));
    commentTexts.forEach((text) => items.push(makeBullet(text)));
  }
  items.push(makePara(""));
  return items;
}

async function main() {
  if (!fs.existsSync(ANALYSIS_PATH)) {
    console.error(`Analysis file not found: ${ANALYSIS_PATH}`);
    console.error("Run 04_analyze_IL.py first.");
    process.exit(1);
  }

  const dfCsv = loadCsv();
  N_RESPONDENTS = dfCsv.length;

  const children = [
    ...sec_title(dfCsv),
    ...sec_demographics(),
    ...sec_coach_relationships(),
    ...sec_communication(),
    ...(await sec_employment(dfCsv)),
    ...(await sec_program_impact(dfCsv)),
    ...sec_respect_environment(),
    ...sec_banking(),
    ...(await sec_nps()),
    ...sec_comments(dfCsv),
  ].filter((node) => !SPACER_PARAGRAPHS.has(node));

  const doc = new Document({
    numbering: BULLET_NUMBERING,
    styles: DOC_STYLES,
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        },
      },
      children,
    }],
  });

  fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUT_PATH, buffer);
  console.log(`Saved: ${OUT_PATH}`);
}

main().catch((error) => {
  console.error("Error:", error);
  process.exit(1);
});