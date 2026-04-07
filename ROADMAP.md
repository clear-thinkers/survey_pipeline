# Survey Pipeline — Roadmap
**Auberle 2026 Participant Survey · 412 Youth Zone (103 surveys) + IL (11 surveys)**
_Last updated: April 2026_

---

## Goal
AI-assisted extraction, QA, and report generation for annual participant surveys. Two survey types: IL (Auberle, s001–s011) and 412YZ (412 Youth Zone, s012–s114).

---

## Directory structure

```
survey_ocr/
├── data/
│   ├── raw/              # Scanned PDFs (s001–s114)
│   └── extracted/        # Per-survey JSON from Claude (s001–s114)
├── output/
│   ├── IL/               # review_s00x.xlsx, survey_data_IL.csv
│   └── 412YZ/            # review_s0xx.xlsx, survey_data_412YZ.csv,
│                         # flagged_412YZ.csv, qa_questions_412YZ.xlsx
├── report/
│   ├── example/          # Reference report from prior year
│   ├── IL/               # Generated IL report output
│   └── 412YZ/            # Generated 412YZ report output
├── scripts/
│   ├── 01_extract.py         # Claude Vision API → JSON per survey
│   ├── review.py             # Per-survey review workbook generator
│   ├── 02_compile.py         # Merge JSONs → type-specific CSV + apply reviewer corrections
│   ├── 03_qa.py              # Validation rules → flagged CSV + reviewer question workbook
│   ├── apply_corrections.py  # Apply reviewer decisions from qa_questions xlsx → CSV
│   ├── 04_analyze.py         # Descriptive stats → tables
│   └── 05_report.py          # Claude API → report.docx
├── prompts/
│   ├── extraction_prompt_IL.txt
│   ├── extraction_prompt_412YZ.txt
│   └── report_prompt.txt
├── config.py
└── requirements.txt
```

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

### Phase 3 — QA (03_qa.py) ✓
Validates `survey_data_412YZ.csv`. Produces:
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
| G — Confidence | Fields with confidence < 0.75 flagged |

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

### Phase 3b — Apply Corrections (apply_corrections.py) ✓
Reads reviewer decisions from `qa_questions_412YZ.xlsx` (QA Questions tab) and writes corrections to `survey_data_412YZ.csv`.

- Actions: `recode` · `clear` · `exclude` · `accept`
- Scope: `this_survey` (default when blank) · `all_surveys`
- Array fields: only the flagged token is modified; other tokens preserved

### Phase 4 — Analysis (04_analyze.py) ✓
Produces `output/412YZ/analysis_412YZ.xlsx` — 22 sheets, one per reporting component. See Reporting Components below.

### Phase 5 — Report generation (05_report.py)
Feed summary stats to Claude Sonnet 4.6 → `report/412YZ/report_412YZ.docx`

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

## Reporting Components — 412YZ (04_analyze.py)

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
