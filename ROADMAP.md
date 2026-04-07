# Survey Pipeline — Roadmap
**Auberle 2026 Participant Survey · 412 Youth Zone (103 paper + 101 online) + IL (11 paper + 20 online)**
_Last updated: April 2026_

---

## Goal
AI-assisted extraction, QA, and report generation for annual participant surveys. Two survey types: IL (Auberle, s001–s011) and 412YZ (412 Youth Zone, s012–s114). Each type includes both paper (OCR pipeline) and online (SurveyMonkey export) respondents.

---

## Directory structure

```
survey_ocr/
├── data/
│   ├── raw/              # Scanned PDFs (s001–s114)
│   ├── extracted/        # Per-survey JSON from Claude (s001–s114)
│   └── online/           # SurveyMonkey Excel exports
│       ├── Youth Zone Survey - Feb 2026.xlsx
│       └── Crawford IL Participant Survey 2026.xlsx
├── output/
│   ├── IL/               # review_s00x.xlsx, survey_data_IL.csv
│   └── 412YZ/            # review_s0xx.xlsx, survey_data_412YZ.csv,
│                         # flagged_412YZ.csv, qa_questions_412YZ.xlsx,
│                         # analysis_412YZ.xlsx
├── report/
│   ├── example/          # Reference report from prior year
│   ├── IL/               # Generated IL report output
│   └── 412YZ/            # Generated 412YZ report output
├── scripts/
│   ├── 01_extract.py                    # [generic] Claude Vision API → JSON per survey
│   ├── review.py                        # [generic] Per-survey review workbook generator
│   ├── 02_compile.py                    # [generic] Merge JSONs → type-specific CSV
│   ├── 02b_ingest_online_412YZ.py       # [412YZ]   SurveyMonkey xlsx → merge into CSV
│   ├── 03_qa_412YZ.py                   # [412YZ]   Validation rules → flagged CSV + workbook
│   ├── 03b_apply_corrections_412YZ.py   # [412YZ]   Apply reviewer decisions → CSV
│   ├── 03c_standardize_fields_412YZ.py  # [412YZ]   Normalize DOB + coach name
│   ├── 04_analyze_412YZ.py              # [412YZ]   Descriptive stats → analysis xlsx
│   └── 05_report_412YZ.py              # [412YZ]   python-docx report generation
├── prompts/
│   ├── extraction_prompt_IL.txt
│   ├── extraction_prompt_412YZ.txt
│   └── report_prompt.txt
├── config.py
└── requirements.txt
```

---

## Complete Pipeline Flow (412YZ)

```
01_extract.py          (generic)   PDF → JSON
   └─ review.py        (generic)   human review workbook
02_compile.py          (generic)   JSON → survey_data_412YZ.csv  [paper rows only, s001-s114]
02b_ingest_online_412YZ.py         SurveyMonkey xlsx → append online rows [o001-o101]
03_qa_412YZ.py         (412YZ)    validate → flagged_412YZ.csv + qa_questions_412YZ.xlsx
03b_apply_corrections_412YZ.py     apply reviewer decisions → survey_data_412YZ.csv
03c_standardize_fields_412YZ.py    normalize DOB + coach_name_corrected
04_analyze_412YZ.py    (412YZ)    descriptive stats → analysis_412YZ.xlsx
05_report_412YZ.py     (412YZ)    generate report_412YZ.docx
```

Scripts marked **[generic]** work for both survey types. Scripts marked **[412YZ]** are specific to the 412 Youth Zone survey. IL counterparts (`03_qa_IL.py`, etc.) will be created separately.

---

## Phases

### Phase 1 — Extraction (01_extract.py) ✓
- Sends each PDF page as base64 JPEG (150 DPI) to Claude Sonnet 4.6
- Returns structured JSON with all survey fields + per-field confidence (0.0–1.0)
- Survey type auto-detected from filename (s001–s011 → IL; s012+ → 412YZ)
- Usage: `python scripts/01_extract.py s001` or run without args for all PDFs

### Phase 1b — Human Review (review.py) ✓
- Generates `output/<type>/review_{id}.xlsx` — one row per field in question order
- Confidence tiers: red < 75% (must review), yellow 75–89% (check), blank ≥ 90% (clean)
- Reviewer enters corrections in column E, notes in column F
- Usage: `python scripts/review.py s001`

### Phase 2 — Compile (02_compile.py) ✓
- Merges all JSONs → `output/IL/survey_data_IL.csv` and `output/412YZ/survey_data_412YZ.csv`
- Applies reviewer corrections from matching `review_{id}.xlsx` if present
- Array fields serialized as pipe-separated (`value1 | value2`); all fields get a `_conf` column
- Single-select fields: if reviewer enters multiple comma-separated values, first value is used

### Phase 2b — Ingest Online Data (02b_ingest_online_412YZ.py) — DONE
Reads `data/online/Youth Zone Survey - Feb 2026.xlsx` (SurveyMonkey export), normalizes all responses to the same schema as the paper CSV, and appends the rows to `survey_data_412YZ.csv`.

**SurveyMonkey export format:**
- Two-row header: row 0 = question text, row 1 = sub-label (option text for multi-select)
- Single-select fields: one column, English label value (e.g., `"Yes, full-time"`)
- Multi-select fields: one column per option, value = label if selected else blank
- Data starts row 2; 101 data rows in the 2026 file

**What the script does:**
1. Assigns survey IDs `o001`–`o101` (distinct from paper `s001`–`s114`)
2. Adds a `source` column — `"online"` for new rows; retroactively sets `"paper"` on existing CSV rows
3. Maps single-select English labels → internal codes (e.g., `"Yes, full-time"` → `"yes_full_time"`, `"All the time"` → `5`)
4. Collapses multi-select checkbox columns → pipe-separated token strings (same format as paper: `"token1 | token2"`)
5. Maps option text → internal tokens for all pipe-sep fields (race, gender, Q10/Q11/Q13/Q14/Q15a-b/Q17/Q24/Q25/Q26b)
6. Sets all `_conf` columns to `1.0` (no OCR uncertainty)
7. Leaves `dob`, `first_initial`, `last_name` blank (not collected online)
8. Writes `coach_name` from the free-text write-in field (col 10); `03c_standardize_fields_412YZ.py` handles canonicalization

**Key design decisions:**
- Merged file keeps a single `survey_data_412YZ.csv`; paper-only file is not preserved separately (back up before running if needed)
- Online rows with partial completions are included (consistent with paper approach)
- Rule G (low confidence) in `03_qa_412YZ.py` is skipped for `source == "online"` rows

### Phase 3 — QA (03_qa_412YZ.py) ✓
Validates `survey_data_412YZ.csv` (paper + online combined). Produces:
- `flagged_412YZ.csv` — machine-readable issue log
- `qa_questions_412YZ.xlsx` — reviewer workbook (tabs: Instructions · QA Questions · Accepted — No Action · Summary)

**Validation rules:**

| Rule | What is checked |
|---|---|
| A — Type/Range | Likert fields (q1, q20) integer 1–5; q22_nps integer 0–10 |
| B — Categorical | Single-select value must be in allowed code list |
| C — Array token | Each pipe-separated token must be in allowed code list |
| D — Conditional | Child fields checked against parent conditions (see below) |
| E — Required | Key fields flagged if blank |
| F — Free-text | Race tokens, gender, orientation non-standard labels flagged |
| G — Confidence | Fields with confidence < 0.75 flagged (skipped for online rows: `source == "online"`) |

**Correction decisions (April 2026):**

| Rule | Decision |
|---|---|
| E — Missing Required | Leave blank — respondents intentionally skipped |
| D — Conditional Missing | Leave blank — follow-up intentionally blank on survey |
| G — Low Confidence | Accept as-is — reviewer verified all values in Phase 1b |
| D — Conditional Violation | Auto-resolved — see rules below |

Rules E, D-missing, and G go to the **Accepted — No Action** tab. Active items needing input go to **QA Questions**.

**Conditional violation rules:**
- Child field non-blank when parent condition not met → **auto-clear child** (`D_auto_clear`)
- `q12_housing_stability` blank but `q13_sleeping_location` filled → **set `q12 = no_place`** (`D_infer_parent`)

### Phase 3b — Apply Corrections (03b_apply_corrections_412YZ.py) ✓
Reads reviewer decisions from `qa_questions_412YZ.xlsx` (QA Questions tab) and writes corrections to `survey_data_412YZ.csv`.

- Actions: `recode` · `clear` · `exclude` · `accept`
- Scope: `this_survey` (default when blank) · `all_surveys`
- Array fields: only the flagged token is modified; other tokens preserved

### Phase 3c — Standardize Fields (03c_standardize_fields_412YZ.py) ✓
Normalizes fields for both paper and online rows before analysis.

- **DOB**: normalizes to MM/DD/YYYY; handles YYYY-MM-DD, MM/DD/YY (→ MM/DD/20YY); explicit overrides for known bad values (s012, s106). Skipped for online rows (no DOB collected).
- **coach_name_corrected**: inserts canonical name column using `COACH_NAME_MAP`; 7 paper entries remain blank (OCR too ambiguous); online coach names go through the same map — map may need extension for new spelling variants.
- Usage: `python scripts/03c_standardize_fields_412YZ.py`

### Phase 4 — Analysis (04_analyze_412YZ.py) ✓
Produces `output/412YZ/analysis_412YZ.xlsx` — 22 sheets, one per reporting component. See Reporting Components below.

### Phase 5 — Report generation (05_report_412YZ.py) ✓
Fully code-driven python-docx script. No Claude API call. → `report/412YZ/report_412YZ.docx`

---

## Configuration (config.py)

| Setting | Value |
|---|---|
| `EXTRACTION_MODEL` | `claude-sonnet-4-6` |
| `REPORT_MODEL` | `claude-sonnet-4-6` |
| `CONFIDENCE_THRESHOLD` | `0.9` |
| `MUST_REVIEW_THRESHOLD` | `0.75` |

---

## Stack
- Python 3.12 (miniconda3) — use `python -m pip install` for packages
- `anthropic` SDK — Sonnet 4.6 for extraction and report
- `pandas` for compilation, QA, analysis
- `python-docx` for report generation
- `pdf2image` + `Pillow` for PDF → image (requires Poppler on PATH)
- `openpyxl` for review and QA workbooks

---

## Reporting Components — 412YZ (04_analyze_412YZ.py)

21 sections from the example report. Code→label mappings live in `extraction_prompt_412YZ.txt`.

### Demographics
1. **Age distribution** — `age_range`; blank = Unknown
2. **Gender × Sexual Orientation** — `gender` grouped (Female / Male / Trans-Non-binary); `sexual_orientation` exact label
3. **Race/Ethnicity (counted once)** — `race_ethnicity` pipe-sep; 1 token → high-level group; 2+ → Multi-Racial
4. **Race/Ethnicity (counted multiple times)** — same pipe-sep split, each token counted independently

   Race groups: Black · White · Multi-Racial · Hispanic or Latinx · East Asian · Native American or Native Hawaiian · Prefer not to answer

### Coach Relationship
5. **Q1 satisfaction** — `q1_*` Likert 1–5; % top-2 box (4–5)
6. **Communication** — `q2_communication_frequency` × `q3_communication_level`; % satisfied = `q3 == good_amount`; cross-tab not_enough by q2

### Housing
7. **Housing stability** — `q12` + `q13` sleeping location for non-stable rows; cross-tab by age
8. **Housing instability reasons** — `q14` (pipe-sep) × `age_range`

### Education & Employment
9. **Education** — `q5_school_status` + `q5a_highest_education` combined into one table
10. **Employment status** — `q8` × `age_range`; tenure table (`q8a` × full/part-time); derived "Seeking Employment" = `q8b == yes`
11. **Job barriers** — `q10_job_barriers` (pipe-sep); denominator = non-full-time employed
12. **Reasons left a job** — `q11_left_job_reasons` + `q11a_quit_reasons` (pipe-sep); q11a are sub-reasons under "quit"

### Transportation & Civic
13. **Transportation** — `q9_primary_transport`; derived "No-Car Combination" = public_transit / rideshare / rides_from_others / active_transport; `q6_drivers_license` + `q6a_vehicle_access` × age
14. **Voter registration** — `q7_registered_to_vote` × age (18–20 and 21–23 only); `q7a_not_registered_reasons` × age

### Engagement & Impact
15. **Visit frequency** — `q15_visit_frequency`; `q15a_visit_reasons` for frequent; `q15b_visit_barriers` for infrequent; × age
16. **Program impact** — `q17_program_helped` (pipe-sep) × age; `q16_stay_focused`; `q21_gained_independence`
17. **Staff & peer respect** — `q18_staff_respect`, `q19_peer_respect`; % top-2 box (often + all_the_time)
18. **Program environment** — `q20_*` Likert 1–5; % top-2 box (4–5)

### Financial
19. **Banking** — `q25_bank_account` × age; `q24_money_methods` × age; `q26b_account_usage` × age; derived `has_account` = q25 includes checking/savings
20. **NPS** — `q22_nps` 0–10; NPS = % Promoters (9–10) − % Detractors (0–6)
21. **Additional comments** — `q23_other_comments` verbatim quotes

---

## Phase 5 Plan — Report Generation (05_report_412YZ.py)

### Approach
Fully code-driven `python-docx` script. No Claude API call — all narrative text is templated directly from the analysis data. Numbers are pulled from `analysis_412YZ.xlsx` (one sheet per section) and injected into f-string templates. This is deterministic and matches the prior year report structure exactly.

**Inputs:**
- `output/412YZ/analysis_412YZ.xlsx` — pre-computed stats (22 sheets)
- `output/412YZ/survey_data_412YZ.csv` — raw CSV for cross-sheet inline stats

**Output:** `report/412YZ/report_412YZ.docx`

### Config / knowns
- Survey month: March 2026
- N respondents: 103
- Total active participants: `[PLACEHOLDER — highlight yellow]`
- Q1 coach satisfaction: single-year + prior year columns from hardcoded benchmark dict

### Styling (matched to prior report)
- Page: 8.5 × 11 in, 1 in margins all sides
- Body: Normal style, default Calibri
- Section headings (FINDINGS, TRANSPORTATION, etc.): bold, black
- Table captions: bold, not all-caps
- Table header row: fill `DCE6F1`, bold text
- Table total row: fill `DCE6F1`, bold text
- No table borders (matches prior report)
- Placeholder highlight: yellow (`FFFF00`) character highlight

### Document sections (19 sections, 21 tables)

| # | Section heading | Tables | Analysis sheet(s) |
|---|---|---|---|
| 0 | Title + intro | — | `01_age` |
| 1 | Survey Respondents by Age | Table 1 | `01_age` |
| 2 | Survey Respondents by Gender and Sexual Orientation | Table 2 | `02_gender_orient` |
| 3 | Youth by Race and Gender (counted once) | Table 3 | `03_race_once` |
| 4 | Youth with Full/Partial Racial Identities (counted multiple times) | Table 4 | `04_race_multi` |
| 5 | Relationships with Coach — satisfaction | Table 5 (multi-year) | `05_q1` + hardcoded benchmarks |
| 6 | Communication frequency/level | inline only | `06_communication` |
| 7 | Stable Housing — status + sleeping arrangements | Table 6 | `07_housing` |
| 8 | Reasons for Unstable Housing by Age | Table 7 | `08_housing_reasons` |
| 9 | Employment and Education — bullets + education table | Table 8 | `09_education`, `10_employment`, raw CSV |
| 10 | Length of Employment (job tenure) | Table 9 | `11_job_tenure` |
| 11 | Employment Status by Age | Table 10 | `10_employment` |
| 12 | Job Barriers | Table 11 | `12_job_barriers` |
| 13 | Reasons Lost or Quit a Job | Table 12 | `13_left_job` |
| 14 | Driver's License Status by Age | Table 13 | `14_transport` |
| 14b | Drivers' Access to Reliable Vehicle by Age | Table 14 | `14_transport` |
| 14c | Primary Way Youth Get to Work | Table 15 | `14_transport` |
| 15 | Voter Registration by Age + reasons | Tables 16–17 | `15_voter_reg` |
| 16 | Zone Visit Frequency + Barriers | Table 18 | `16_visit` |
| 17 | Program Impact (Q17 by age) | Table 19 | `17_impact` |
| 18 | Staff/Peer Respect + Environment (Q20) | inline | `18_respect`, `19_environment` |
| 19 | Banking status + money methods + account usage | Tables 20–22 | `20_banking` |
| 20 | NPS | inline | `21_nps` |
| 21 | Additional Comments | bullet list | `22_comments` |

### Code structure
```
05_report_412YZ.py
├── Constants
│   ├── Q1_BENCHMARKS       dict of prior-year % top-2 per Q1 field
│   └── PLACEHOLDER_COLOR   yellow highlight for manual fill-ins
├── load_sheets()           reads all analysis xlsx sheets into a dict of DataFrames
├── Styling helpers
│   ├── para()              add paragraph with optional bold / style
│   ├── caption()           bold table-caption paragraph
│   ├── heading()           section heading (bold)
│   ├── bullet()            List Paragraph style
│   └── placeholder()       highlighted yellow run for manual fill-ins
├── Table helpers
│   ├── add_table()         write DataFrame to doc table; header + total row shading
│   └── shade_row()         apply DCE6F1 fill to a table row
├── Section functions       one per section: sec_demographics() … sec_comments()
└── main()                  instantiate Document, call sections, save .docx
```

### Placeholders (highlighted yellow in output)
- Response rate denominator (total active participants)
- Any inline cross-year comparison stats not computable from current data alone
